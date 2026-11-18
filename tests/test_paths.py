import sqlite3

import pytest

from fleet.paths import database_path


def test_legacy_database_migration_includes_uncheckpointed_wal(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    old = tmp_path / 'FlotaAdriana' / 'flota.sqlite3'
    old.parent.mkdir()
    connection = sqlite3.connect(old)
    try:
        connection.execute('PRAGMA journal_mode=WAL')
        connection.execute('CREATE TABLE records(value TEXT)')
        connection.execute("INSERT INTO records VALUES ('dane i szkice')")
        connection.commit()
        target = database_path()
        assert target == tmp_path / 'Floteka' / 'flota.sqlite3'
        with sqlite3.connect(target) as migrated:
            assert migrated.execute('SELECT * FROM records').fetchall() == [('dane i szkice',)]
            assert migrated.execute('PRAGMA integrity_check').fetchone()[0] == 'ok'
        assert connection.execute('SELECT * FROM records').fetchall() == [('dane i szkice',)]
    finally:
        connection.close()


def test_existing_new_database_is_never_replaced(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    target = tmp_path / 'Floteka' / 'flota.sqlite3'
    target.parent.mkdir()
    target.write_bytes(b'new database, preserve this')
    old = tmp_path / 'FlotaAdriana' / 'flota.sqlite3'
    old.parent.mkdir()
    old.write_bytes(b'old database')
    assert database_path() == target
    assert target.read_bytes() == b'new database, preserve this'


def test_corrupt_legacy_database_does_not_leave_empty_target(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    old = tmp_path / 'FlotaAdriana' / 'flota.sqlite3'
    old.parent.mkdir()
    old.write_bytes(b'not SQLite')
    with pytest.raises(sqlite3.DatabaseError):
        database_path()
    assert not (tmp_path / 'Floteka' / 'flota.sqlite3').exists()
    assert old.read_bytes() == b'not SQLite'


def test_explicit_database_bypasses_legacy_migration(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    target = tmp_path / 'portable.sqlite3'
    assert database_path(target) == target
    assert not (tmp_path / 'Floteka').exists()
