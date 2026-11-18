import sqlite3

import pytest

from fleet.store import DraftConfirmationRequired, Store


@pytest.fixture
def fleet(tmp_path):
    store = Store(tmp_path / 'optional.sqlite3')
    truck = store.add_vehicle('truck', 'WX1')
    driver = store.add_driver('Anna', truck, phone='+48 123 456 789')
    yield store, driver, truck
    store.close()


def test_optional_only_is_draft_and_requires_confirmation_before_clear(fleet):
    store, driver, truck = fleet
    store.save_slot(driver, 1, route='Poznań → Berlin', order_number='ZL-1')
    assert store.slots(driver)[0]['status'] == 'Szkic'
    assert not store.history()
    with pytest.raises(DraftConfirmationRequired):
        store.clear_slot(driver, 1)
    store.clear_slot(driver, 1, allow_draft=True)
    assert store.slots(driver)[0]['route'] == ''
    assert store.slots(driver)[0]['order_number'] == ''


def test_optional_entry_versions_restore_and_restart(fleet):
    store, driver, truck = fleet
    data = dict(trailer='PO1', load_date='17.11.2026', unload_date='18.11.2026', route='Łódź → Gdańsk')
    store.save_slot(driver, 1, **data, order_number='ZL-1')
    first = store.history()[0]
    store.save_slot(driver, 1, **data, order_number='ZL-2')
    assert len(store.history()) == 2
    store.clear_slot(driver, 1)
    store.restore(first['id'], 2)
    assert store.slots(driver)[1]['route'] == 'Łódź → Gdańsk'
    assert store.slots(driver)[1]['order_number'] == 'ZL-1'
    assert len(store.history()) == 2
    other = Store(store.path)
    try:
        assert other.driver(driver)['phone'] == '+48 123 456 789'
        assert other.slots(driver)[1]['order_number'] == 'ZL-1'
        assert other.conn.execute('PRAGMA user_version').fetchone()[0] == 3
        with pytest.raises(sqlite3.IntegrityError):
            other.conn.execute('DELETE FROM history')
    finally:
        other.close()


def test_truck_optional_dates_and_odometer_validation_preserve_previous(fleet):
    store, driver, truck = fleet
    store.save_truck_details(truck, odometer_km='123456', inspection_date='18.11.2026',
                             insurance_date='', service_date='2027-01-01', notes='Opony zimowe')
    before = store.vehicles('truck')[0]
    assert before['odometer_km'] == 123456
    assert before['inspection_date'] == '2026-11-18'
    assert before['insurance_date'] == ''
    for bad in ('-1', '12.5', 'abc'):
        with pytest.raises(ValueError):
            store.save_truck_details(truck, odometer_km=bad)
        assert store.vehicles('truck')[0] == before
    with pytest.raises(ValueError):
        store.save_truck_details(truck, inspection_date='31.02.2026')
    assert store.vehicles('truck')[0] == before


def test_version_two_migration_keeps_immutable_history_and_drafts(tmp_path):
    path = tmp_path / 'old-v2.sqlite3'
    with sqlite3.connect(path) as old:
        old.executescript('''
            PRAGMA user_version=2;
            CREATE TABLE vehicles(id INTEGER PRIMARY KEY,kind TEXT,plate TEXT,active INTEGER DEFAULT 1, UNIQUE(kind,plate));
            CREATE TABLE drivers(id INTEGER PRIMARY KEY,name TEXT,truck_id INTEGER REFERENCES vehicles(id),active INTEGER DEFAULT 1);
            CREATE TABLE counter(id INTEGER PRIMARY KEY,value INTEGER);
            INSERT INTO counter VALUES(1,9);
            CREATE TABLE slots(driver_id INTEGER REFERENCES drivers(id),position INTEGER CHECK(position>=1),
              trailer TEXT DEFAULT '',info TEXT DEFAULT '',load_date TEXT DEFAULT '',unload_date TEXT DEFAULT '',
              started_order INTEGER,PRIMARY KEY(driver_id,position));
            CREATE TABLE history(id INTEGER PRIMARY KEY,driver_id INTEGER REFERENCES drivers(id),driver_name TEXT,
              truck TEXT,trailer TEXT,info TEXT,load_date TEXT,unload_date TEXT,fingerprint TEXT,recorded_at TEXT,
              UNIQUE(driver_id,fingerprint));
            CREATE TRIGGER history_no_update BEFORE UPDATE ON history BEGIN SELECT RAISE(ABORT,'immutable'); END;
            CREATE TRIGGER history_no_delete BEFORE DELETE ON history BEGIN SELECT RAISE(ABORT,'immutable'); END;
            INSERT INTO vehicles VALUES(8,'truck','WX8',1);
            INSERT INTO drivers VALUES(7,'Dawny kierowca',8,1);
            INSERT INTO slots VALUES(7,1,'PO8','zachowany szkic','','',9);
            INSERT INTO history VALUES(44,7,'Dawny kierowca','WX8','PO9','dawna historia',
              '2026-11-17','2026-11-18','dawny-klucz','2026-11-17T10:00:00+00:00');
        ''')
    store = Store(path)
    try:
        assert store.slots(7)[0]['info'] == 'zachowany szkic'
        assert store.slots(7)[0]['started_order'] == 9
        assert store.history()[0]['id'] == 44
        assert store.history()[0]['fingerprint'] == 'dawny-klucz'
        assert store.history()[0]['route'] == ''
        assert store.driver(7)['phone'] == ''
        store.save_truck_details(8, odometer_km='100')
        assert store.vehicles('truck')[0]['odometer_km'] == 100
        assert store.conn.execute('PRAGMA foreign_key_check').fetchall() == []
        with pytest.raises(sqlite3.IntegrityError):
            store.conn.execute('UPDATE history SET info=\'modified\'')
    finally:
        store.close()


def test_two_windows_saving_same_empty_place_keep_both_drafts(fleet):
    store, driver, truck = fleet
    snapshot = store.slots(driver)[0]
    other = Store(store.path)
    try:
        assert store.save_slot(driver, 1, info='pierwszy szkic', expected=snapshot) == 1
        assert other.save_slot(driver, 1, info='drugi szkic', expected=snapshot) == 2
        assert [slot['info'] for slot in store.slots(driver)][:2] == ['pierwszy szkic', 'drugi szkic']
        assert not store.history()
    finally:
        other.close()
