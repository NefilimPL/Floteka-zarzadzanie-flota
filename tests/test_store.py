import sqlite3
from concurrent.futures import ThreadPoolExecutor
from threading import Barrier

import pytest

from fleet.store import DraftConfirmationRequired, Store


@pytest.fixture
def fleet(tmp_path):
    store = Store(tmp_path / 'flota.sqlite3')
    truck = store.add_vehicle('truck', 'WX 12345')
    driver = store.add_driver('Łukasz Kowalski', truck)
    yield store, driver
    store.close()


def complete(store, driver, slot=1, **changes):
    data = dict(trailer='PO 23456', info='Firma Żółć', load_date='2026-11-17', unload_date='2026-11-20')
    data.update(changes)
    return store.save_slot(driver, slot, **data)


def test_more_than_four_slots_survive_restart_and_restore(fleet):
    store, driver = fleet
    for expected in range(5, 13):
        assert store.add_slot(driver) == expected
        complete(store, driver, expected, info=f'Zlecenie {expected}')
    store.save_slot(driver, 5, info='szkic')
    history = store.history()
    other = Store(store.path)
    try:
        assert len(other.slots(driver)) == 12
        assert other.slots(driver)[4]['info'] == 'szkic'
        assert other.history() == history
        assert other.restore(history[0]['id']) == 13
        assert other.slots(driver)[12]['status'] == 'Kompletny'
        assert other.history() == history
        assert other.clear_oldest(driver, allow_draft=True) == 5
    finally:
        other.close()


def test_migration_preserves_v1_drafts_history_and_order(fleet):
    store, driver = fleet
    complete(store, driver, 3)
    store.save_slot(driver, 2, info='szkic ze starej bazy')
    before, history = store.slots(driver), store.history()
    # Rzeczywisty schemat starego wydania, niezależny od nowego SCHEMA.
    store.conn.executescript('''
        CREATE TABLE legacy_slots (
            driver_id INTEGER NOT NULL REFERENCES drivers(id),
            position INTEGER NOT NULL CHECK(position BETWEEN 1 AND 4),
            trailer TEXT NOT NULL DEFAULT '', info TEXT NOT NULL DEFAULT '',
            load_date TEXT NOT NULL DEFAULT '', unload_date TEXT NOT NULL DEFAULT '',
            started_order INTEGER, PRIMARY KEY(driver_id, position));
        INSERT INTO legacy_slots SELECT driver_id,position,trailer,info,load_date,unload_date,started_order FROM slots;
        DROP TABLE slots;
        ALTER TABLE legacy_slots RENAME TO slots;
        PRAGMA user_version=1;
    ''')
    other = Store(store.path)
    try:
        assert other.slots(driver) == before
        assert other.history() == history
        assert other.add_slot(driver) == 5
        complete(other, driver, 5)
        assert other.clear_oldest(driver) == 3
        assert other.conn.execute('PRAGMA integrity_check').fetchone()[0] == 'ok'
        assert other.conn.execute('PRAGMA foreign_key_check').fetchall() == []
        assert other.conn.execute('PRAGMA user_version').fetchone()[0] == 3
        with pytest.raises(sqlite3.IntegrityError):
            other.conn.execute('DELETE FROM history')
    finally:
        other.close()


def test_simultaneous_first_starts_recheck_schema_inside_transaction(fleet, monkeypatch):
    store, driver = fleet
    store.save_slot(driver, 1, info='szkic sprzed aktualizacji')
    store.conn.executescript('''
        CREATE TABLE legacy_slots(driver_id INTEGER REFERENCES drivers(id),position INTEGER CHECK(position BETWEEN 1 AND 4),
          trailer TEXT DEFAULT '',info TEXT DEFAULT '',load_date TEXT DEFAULT '',unload_date TEXT DEFAULT '',
          started_order INTEGER,PRIMARY KEY(driver_id,position));
        INSERT INTO legacy_slots SELECT driver_id,position,trailer,info,load_date,unload_date,started_order FROM slots;
        DROP TABLE slots;
        ALTER TABLE legacy_slots RENAME TO slots;
        PRAGMA user_version=1;
    ''')
    barrier = Barrier(2)
    connect = sqlite3.connect

    class SynchronizedConnection(sqlite3.Connection):
        initial_version_read = False

        def execute(self, sql, *args, **kwargs):
            result = super().execute(sql, *args, **kwargs)
            if sql == 'PRAGMA user_version' and not self.initial_version_read:
                self.initial_version_read = True
                barrier.wait(timeout=10)
            return result

    monkeypatch.setattr('fleet.store.sqlite3.connect', lambda *a, **k: connect(*a, **k, factory=SynchronizedConnection))

    def start():
        other = Store(store.path)
        try:
            return other.slots(driver)[0]['info']
        finally:
            other.close()

    with ThreadPoolExecutor(max_workers=2) as pool:
        assert list(pool.map(lambda _: start(), range(2))) == ['szkic sprzed aktualizacji'] * 2


@pytest.mark.parametrize('position', [0, -1, 1.5, True, 99])
def test_invalid_or_missing_slot_never_archives(fleet, position):
    store, driver = fleet
    with pytest.raises(ValueError):
        complete(store, driver, position)
    assert not store.history()


def test_four_slots_and_draft_survives_restart(fleet):
    store, driver = fleet
    assert [s['position'] for s in store.slots(driver)] == [1, 2, 3, 4]
    store.save_slot(driver, 2, trailer='PO23456', info='czeka', load_date='', unload_date='')
    other = Store(store.path)
    try:
        assert other.slots(driver)[1]['info'] == 'czeka'
        assert other.slots(driver)[1]['status'] == 'Szkic'
        assert other.history() == []
    finally:
        other.close()


def test_complete_active_record_is_already_in_history(fleet):
    store, driver = fleet
    complete(store, driver)
    assert store.slots(driver)[0]['status'] == 'Kompletny'
    assert len(store.history()) == 1
    assert store.history()[0]['driver_name'] == 'Łukasz Kowalski'


def test_identical_content_deduplicated_but_other_driver_is_distinct(fleet):
    store, driver = fleet
    complete(store, driver)
    complete(store, driver)
    complete(store, driver, 2, trailer='po23456')
    assert len(store.history()) == 1
    second = store.add_driver('Łukasz Kowalski', store.drivers()[0]['truck_id'])
    complete(store, second)
    assert len(store.history()) == 2


def test_edit_appends_version_and_clearing_preserves_both(fleet):
    store, driver = fleet
    complete(store, driver)
    complete(store, driver, info='Nowa firma')
    store.clear_slot(driver, 1)
    assert {h['info'] for h in store.history()} == {'Firma Żółć', 'Nowa firma'}
    assert store.slots(driver)[0]['status'] == 'Puste'


def test_optional_info_and_partial_dates(fleet):
    store, driver = fleet
    complete(store, driver, info='', unload_date='')
    assert not store.history()
    complete(store, driver, info='')
    assert len(store.history()) == 1


@pytest.mark.parametrize('changes', [dict(load_date='31.02.2026'), dict(unload_date='2026-11-16')])
def test_invalid_dates_leave_previous_data_untouched(fleet, changes):
    store, driver = fleet
    complete(store, driver)
    before = store.slots(driver)
    with pytest.raises(ValueError):
        complete(store, driver, **changes)
    assert store.slots(driver) == before
    assert len(store.history()) == 1


def test_polish_date_input_and_duplicate_plates(fleet):
    store, driver = fleet
    complete(store, driver, load_date='17.11.2026', unload_date='20.11.2026')
    assert store.history()[0]['load_date'] == '2026-11-17'
    assert store.add_vehicle('truck', 'wx12345') == store.drivers()[0]['truck_id']
    assert len(store.vehicles('truck')) == 1


def test_draft_clear_needs_explicit_confirmation(fleet):
    store, driver = fleet
    store.save_slot(driver, 1, trailer='PO1', info='', load_date='', unload_date='')
    with pytest.raises(DraftConfirmationRequired):
        store.clear_slot(driver, 1)
    assert store.slots(driver)[0]['trailer'] == 'PO1'
    store.clear_slot(driver, 1, allow_draft=True)
    assert store.slots(driver)[0]['status'] == 'Puste'
    assert not store.history()


def test_saving_empty_fields_over_draft_needs_confirmation(fleet):
    store, driver = fleet
    store.save_slot(driver, 1, trailer='PO1')
    with pytest.raises(DraftConfirmationRequired):
        store.save_slot(driver, 1)
    assert store.slots(driver)[0]['trailer'] == 'PO1'
    store.save_slot(driver, 1, allow_draft=True)
    assert store.slots(driver)[0]['status'] == 'Puste'


def test_oldest_is_occupancy_order_not_date_or_position(fleet):
    store, driver = fleet
    complete(store, driver, 3)
    complete(store, driver, 1, trailer='PO2', load_date='2026-01-01')
    complete(store, driver, 3, info='poprawka')
    assert store.clear_oldest(driver) == 3
    assert store.slots(driver)[2]['status'] == 'Puste'
    assert store.slots(driver)[0]['trailer'] == 'PO2'
    assert len(store.history()) == 3


def test_history_cannot_be_deleted_or_updated(fleet):
    store, driver = fleet
    complete(store, driver)
    for sql in ['DELETE FROM history', "UPDATE history SET info='znikło'"]:
        with pytest.raises(sqlite3.IntegrityError):
            store.conn.execute(sql)
    assert len(store.history()) == 1


def test_archive_failure_cannot_clear_slot(fleet):
    store, driver = fleet
    complete(store, driver)
    store.conn.execute("CREATE TRIGGER fail_insert BEFORE INSERT ON history BEGIN SELECT RAISE(ABORT, 'test failure'); END")
    with pytest.raises(sqlite3.IntegrityError):
        store.clear_slot(driver, 1)
    assert store.slots(driver)[0]['status'] == 'Kompletny'


def test_restore_only_into_empty_slot_and_current_truck(fleet):
    store, driver = fleet
    complete(store, driver)
    hid = store.history()[0]['id']
    with pytest.raises(ValueError):
        store.restore(hid, 1)
    store.clear_slot(driver, 1)
    store.restore(hid, 2)
    assert len(store.history()) == 1
    new_truck = store.add_vehicle('truck', 'KR999')
    store.assign_truck(driver, new_truck)
    store.restore(hid, 1)
    assert store.slots(driver)[0]['truck'] == 'KR999'
    assert len(store.history()) == 2
    assert {h['truck'] for h in store.history()} == {'WX12345', 'KR999'}


def test_deactivation_rename_and_backup_preserve_history(fleet, tmp_path):
    store, driver = fleet
    complete(store, driver)
    store.rename_driver(driver, 'Nowe nazwisko')
    assert {h['driver_name'] for h in store.history()} == {'Łukasz Kowalski', 'Nowe nazwisko'}
    store.set_driver_active(driver, False)
    assert not store.drivers()
    assert len(store.history()) == 2
    dest = tmp_path / 'backup.sqlite3'
    store.backup(dest)
    other = Store(dest)
    assert len(other.history()) == 2
    other.close()


def test_search_filter_and_invalid_slot(fleet):
    store, driver = fleet
    complete(store, driver)
    assert len(store.history(driver, 'Żółć')) == 1
    assert not store.history(driver, 'nieistniejące')
    with pytest.raises(ValueError):
        complete(store, driver, 5)
