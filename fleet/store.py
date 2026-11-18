"""Transakcyjne dane bieżące i historia tylko do dopisywania."""

from __future__ import annotations

import hashlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


class DraftConfirmationRequired(ValueError):
    pass


def plate(value: str) -> str:
    return ''.join(str(value).upper().split())


def date_value(value: str) -> str:
    value = str(value or '').strip()
    if not value:
        return ''
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError('Nieprawidłowa data. Użyj DD.MM.RRRR, np. 17.11.2026.')


SCHEMA = """
CREATE TABLE IF NOT EXISTS vehicles (
    id INTEGER PRIMARY KEY, kind TEXT NOT NULL CHECK(kind IN ('truck','trailer')),
    plate TEXT NOT NULL CHECK(length(plate)>0), active INTEGER NOT NULL DEFAULT 1,
    UNIQUE(kind, plate)
);
CREATE TABLE IF NOT EXISTS drivers (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL CHECK(length(trim(name))>0),
    truck_id INTEGER NOT NULL REFERENCES vehicles(id), active INTEGER NOT NULL DEFAULT 1
);
CREATE TABLE IF NOT EXISTS counter (id INTEGER PRIMARY KEY CHECK(id=1), value INTEGER NOT NULL);
INSERT OR IGNORE INTO counter VALUES (1,0);
CREATE TABLE IF NOT EXISTS slots (
    driver_id INTEGER NOT NULL REFERENCES drivers(id),
    position INTEGER NOT NULL CHECK(position >= 1),
    trailer TEXT NOT NULL DEFAULT '', info TEXT NOT NULL DEFAULT '',
    load_date TEXT NOT NULL DEFAULT '', unload_date TEXT NOT NULL DEFAULT '',
    started_order INTEGER, PRIMARY KEY(driver_id, position)
);
CREATE TABLE IF NOT EXISTS history (
    id INTEGER PRIMARY KEY, driver_id INTEGER NOT NULL REFERENCES drivers(id),
    driver_name TEXT NOT NULL CHECK(length(trim(driver_name))>0),
    truck TEXT NOT NULL CHECK(length(truck)>0), trailer TEXT NOT NULL CHECK(length(trailer)>0),
    info TEXT NOT NULL, load_date TEXT NOT NULL CHECK(length(load_date)=10),
    unload_date TEXT NOT NULL CHECK(length(unload_date)=10 AND unload_date>=load_date),
    fingerprint TEXT NOT NULL, recorded_at TEXT NOT NULL,
    UNIQUE(driver_id, fingerprint)
);
CREATE TRIGGER IF NOT EXISTS history_no_update BEFORE UPDATE ON history
BEGIN SELECT RAISE(ABORT, 'Historia jest niezmienna'); END;
CREATE TRIGGER IF NOT EXISTS history_no_delete BEFORE DELETE ON history
BEGIN SELECT RAISE(ABORT, 'Nie wolno usuwac historii'); END;
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
"""

OPTIONAL_COLUMNS = {
    'drivers': {'phone': "TEXT NOT NULL DEFAULT ''"},
    'vehicles': {
        'odometer_km': 'INTEGER', 'inspection_date': "TEXT NOT NULL DEFAULT ''",
        'insurance_date': "TEXT NOT NULL DEFAULT ''", 'service_date': "TEXT NOT NULL DEFAULT ''",
        'notes': "TEXT NOT NULL DEFAULT ''",
    },
    'slots': {'route': "TEXT NOT NULL DEFAULT ''", 'order_number': "TEXT NOT NULL DEFAULT ''"},
    'history': {'route': "TEXT NOT NULL DEFAULT ''", 'order_number': "TEXT NOT NULL DEFAULT ''"},
}


class Store:
    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path, timeout=10)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute('PRAGMA foreign_keys=ON')
        self.conn.execute('PRAGMA journal_mode=WAL')
        self.conn.execute('PRAGMA synchronous=FULL')
        version = self.conn.execute('PRAGMA user_version').fetchone()[0]
        if version > 3:
            self.close()
            raise ValueError('Baza pochodzi z nowszej wersji programu.')
        self.conn.executescript(SCHEMA)
        try:
            # DDL i kopiowanie w jednej transakcji: nie zostawiamy częściowej migracji.
            self.conn.execute('BEGIN IMMEDIATE')
            # Inne uruchomienie mogło zakończyć migrację podczas oczekiwania na blokadę.
            version = self.conn.execute('PRAGMA user_version').fetchone()[0]
            if version > 3:
                raise ValueError('Baza pochodzi z nowszej wersji programu.')
            if version == 1:
                self.conn.execute('''CREATE TABLE slots_v2 (
                    driver_id INTEGER NOT NULL REFERENCES drivers(id),
                    position INTEGER NOT NULL CHECK(position >= 1),
                    trailer TEXT NOT NULL DEFAULT '', info TEXT NOT NULL DEFAULT '',
                    load_date TEXT NOT NULL DEFAULT '', unload_date TEXT NOT NULL DEFAULT '',
                    started_order INTEGER, PRIMARY KEY(driver_id, position))''')
                self.conn.execute('INSERT INTO slots_v2 SELECT * FROM slots')
                self.conn.execute('DROP TABLE slots')
                self.conn.execute('ALTER TABLE slots_v2 RENAME TO slots')
            if version < 3:
                for table, definitions in OPTIONAL_COLUMNS.items():
                    existing = {row['name'] for row in self.conn.execute(f'PRAGMA table_info({table})')}
                    for column, definition in definitions.items():
                        if column not in existing:
                            self.conn.execute(f'ALTER TABLE {table} ADD COLUMN {column} {definition}')
            self.conn.execute('PRAGMA user_version=3')
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            self.close()
            raise

    def close(self):
        self.conn.close()

    def get_setting(self, key, default=''):
        row = self.conn.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
        return row['value'] if row else default

    def set_setting(self, key, value):
        with self.conn:
            self.conn.execute('INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',
                              (key, str(value)))

    def save_truck_details(self, truck_id, odometer_km='', inspection_date='', insurance_date='', service_date='', notes=''):
        self._truck(truck_id)
        text = str(odometer_km if odometer_km is not None else '').strip()
        if text and (not text.isascii() or not text.isdecimal() or int(text) > 9_223_372_036_854_775_807):
            raise ValueError('Przebieg musi być nieujemną liczbą całkowitą kilometrów.')
        notes = str(notes or '').strip()
        if len(notes) > 4000:
            raise ValueError('Uwagi tira: maksymalnie 4000 znaków.')
        values = (int(text) if text else None, date_value(inspection_date), date_value(insurance_date),
                  date_value(service_date), notes, truck_id)
        with self.conn:
            self.conn.execute('UPDATE vehicles SET odometer_km=?,inspection_date=?,insurance_date=?,service_date=?,notes=? WHERE id=?', values)

    def add_vehicle(self, kind: str, number: str) -> int:
        number = plate(number)
        if kind not in ('truck', 'trailer') or not number:
            raise ValueError('Podaj numer rejestracyjny i prawidłowy rodzaj pojazdu.')
        with self.conn:
            self.conn.execute(
                'INSERT INTO vehicles(kind,plate) VALUES (?,?) ON CONFLICT(kind,plate) DO UPDATE SET active=1',
                (kind, number),
            )
        return self.conn.execute('SELECT id FROM vehicles WHERE kind=? AND plate=?', (kind, number)).fetchone()[0]

    def vehicles(self, kind: str, include_inactive=False) -> list[dict]:
        sql = 'SELECT * FROM vehicles WHERE kind=?'
        if not include_inactive:
            sql += ' AND active=1'
        return [dict(r) for r in self.conn.execute(sql + ' ORDER BY plate', (kind,))]

    def _truck(self, truck_id):
        truck = self.conn.execute("SELECT * FROM vehicles WHERE id=? AND kind='truck' AND active=1", (truck_id,)).fetchone()
        if not truck:
            raise ValueError('Wybierz aktywnego tira z rejestru.')
        return truck

    def add_driver(self, name: str, truck_id: int, phone='') -> int:
        name = name.strip()
        if not name:
            raise ValueError('Podaj imię i nazwisko kierowcy.')
        phone = str(phone or '').strip()
        if len(phone) > 100:
            raise ValueError('Telefon: maksymalnie 100 znaków.')
        self._truck(truck_id)
        with self.conn:
            cursor = self.conn.execute('INSERT INTO drivers(name,truck_id,phone) VALUES (?,?,?)', (name, truck_id, str(phone or '').strip()))
            driver = cursor.lastrowid
            self.conn.executemany('INSERT INTO slots(driver_id,position) VALUES (?,?)', [(driver, n) for n in range(1, 5)])
        return driver

    def drivers(self, include_inactive=False) -> list[dict]:
        sql = 'SELECT d.*, v.plate AS truck FROM drivers d JOIN vehicles v ON v.id=d.truck_id'
        if not include_inactive:
            sql += ' WHERE d.active=1'
        return [dict(r) for r in self.conn.execute(sql + ' ORDER BY d.name COLLATE NOCASE,d.id')]

    def _add_slot(self, driver_id):
        if not self.driver(driver_id)['active']:
            raise ValueError('Najpierw aktywuj kierowcę w rejestrze.')
        position = self.conn.execute(
            'SELECT COALESCE(MAX(position),0)+1 FROM slots WHERE driver_id=?', (driver_id,)
        ).fetchone()[0]
        self.conn.execute('INSERT INTO slots(driver_id,position) VALUES (?,?)', (driver_id, position))
        return position

    def add_slot(self, driver_id):
        with self.conn:
            self.conn.execute('BEGIN IMMEDIATE')
            return self._add_slot(driver_id)

    def driver(self, driver_id: int) -> dict:
        row = self.conn.execute(
            'SELECT d.*, v.plate AS truck FROM drivers d JOIN vehicles v ON v.id=d.truck_id WHERE d.id=?', (driver_id,)
        ).fetchone()
        if row is None:
            raise ValueError('Nie znaleziono kierowcy.')
        return dict(row)

    def slots(self, driver_id: int) -> list[dict]:
        driver = self.driver(driver_id)
        result = []
        for row in self.conn.execute('SELECT * FROM slots WHERE driver_id=? ORDER BY position', (driver_id,)):
            data = dict(row)
            data.update(driver_name=driver['name'], truck=driver['truck'])
            data['status'] = self._status(data)
            result.append(data)
        return result

    @staticmethod
    def _status(data):
        if not any(data.get(k) for k in ('trailer', 'info', 'load_date', 'unload_date', 'route', 'order_number')):
            return 'Puste'
        if all(data[k] for k in ('driver_name', 'truck', 'trailer', 'load_date', 'unload_date')):
            return 'Kompletny'
        return 'Szkic'

    def _archive(self, data):
        if self._status(data) != 'Kompletny':
            return
        keys = ('driver_name', 'truck', 'trailer', 'info', 'load_date', 'unload_date')
        values = [data[k] for k in keys]
        optional = [data.get(k, '') for k in ('route', 'order_number')]
        # Puste nowe pola zachowują klucze poprzednich wydań.
        fingerprint = hashlib.sha256(json.dumps(values + (optional if any(optional) else []), ensure_ascii=False).encode('utf-8')).hexdigest()
        self.conn.execute(
            '''INSERT INTO history(driver_id,driver_name,truck,trailer,info,load_date,unload_date,fingerprint,recorded_at,route,order_number)
               VALUES (?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(driver_id,fingerprint) DO NOTHING''',
            (data['driver_id'], *values, fingerprint, datetime.now(timezone.utc).isoformat(timespec='seconds'), *optional),
        )

    @staticmethod
    def _position(position):
        if type(position) is not int or position < 1:
            raise ValueError('Numer miejsca musi być dodatnią liczbą całkowitą.')

    def _slot(self, driver_id, position):
        self._position(position)
        row = self.conn.execute('SELECT * FROM slots WHERE driver_id=? AND position=?', (driver_id, position)).fetchone()
        if row is None:
            raise ValueError('Nie znaleziono miejsca. Dodaj nowe miejsce przyciskiem „+ Dodaj wpis”.')
        return row

    def _save_slot(self, driver_id, position, trailer, info, load_date, unload_date, allow_draft=False, route='', order_number=''):
        self._position(position)
        driver = self.driver(driver_id)
        if not driver['active']:
            raise ValueError('Najpierw aktywuj kierowcę w rejestrze.')
        trailer, info = plate(trailer), str(info or '').strip()
        route, order_number = str(route or '').strip(), str(order_number or '').strip()
        if max(len(route), len(order_number)) > 4000:
            raise ValueError('Trasa i numer zlecenia: maksymalnie 4000 znaków w polu.')
        load_date, unload_date = date_value(load_date), date_value(unload_date)
        if load_date and unload_date and unload_date < load_date:
            raise ValueError('Data rozładunku nie może być wcześniejsza od załadunku.')
        old = self._slot(driver_id, position)
        occupied = any((trailer, info, load_date, unload_date, route, order_number))
        previous = dict(old, driver_name=driver['name'], truck=driver['truck'])
        if not occupied and self._status(previous) == 'Szkic' and not allow_draft:
            raise DraftConfirmationRequired('Ten szkic nie ma kompletnej kopii w historii. Czy go wyczyścić?')
        order = old['started_order']
        if occupied and order is None:
            self.conn.execute('UPDATE counter SET value=value+1 WHERE id=1')
            order = self.conn.execute('SELECT value FROM counter WHERE id=1').fetchone()[0]
        if not occupied:
            order = None
        data = dict(driver_id=driver_id, driver_name=driver['name'], truck=driver['truck'],
                    trailer=trailer, info=info, load_date=load_date, unload_date=unload_date, route=route, order_number=order_number)
        self._archive(data)
        if trailer:
            self.conn.execute("INSERT INTO vehicles(kind,plate) VALUES ('trailer',?) ON CONFLICT(kind,plate) DO NOTHING", (trailer,))
        self.conn.execute(
            'UPDATE slots SET trailer=?,info=?,load_date=?,unload_date=?,started_order=?,route=?,order_number=? WHERE driver_id=? AND position=?',
            (trailer, info, load_date, unload_date, order, route, order_number, driver_id, position),
        )

    def save_slot(self, driver_id, position, trailer='', info='', load_date='', unload_date='', allow_draft=False, *, route='', order_number='', expected=None):
        with self.conn:
            self.conn.execute('BEGIN IMMEDIATE')
            if expected is not None:
                fields = ('trailer', 'info', 'load_date', 'unload_date', 'route', 'order_number')
                previous = self._slot(driver_id, position)
                if any(previous[key] != expected.get(key, '') for key in fields):
                    if any(expected.get(key) for key in fields):
                        raise ValueError('Wpis zmienił się w innym oknie. Twoje dane pozostają w formularzu; skopiuj je przed ponownym wczytaniem wpisu.')
                    # Dwa okna mogły wybrać to samo wolne miejsce. Nowy wpis
                    # zapisujemy w innym wolnym miejscu pod tą samą blokadą.
                    empty = next((slot['position'] for slot in self.slots(driver_id) if slot['status'] == 'Puste'), None)
                    position = empty if empty is not None else self._add_slot(driver_id)
            self._save_slot(driver_id, position, trailer, info, load_date, unload_date, allow_draft, route, order_number)
            return position

    def _clear_slot(self, driver_id, position, allow_draft=False):
        self._slot(driver_id, position)
        data = self.slots(driver_id)[position - 1]
        if data['status'] == 'Szkic' and not allow_draft:
            raise DraftConfirmationRequired('Ten szkic nie ma kompletnej kopii w historii. Czy go wyczyścić?')
        self._archive(data)
        self.conn.execute(
            "UPDATE slots SET trailer='',info='',load_date='',unload_date='',route='',order_number='',started_order=NULL WHERE driver_id=? AND position=?",
            (driver_id, position),
        )

    def clear_slot(self, driver_id, position, allow_draft=False):
        with self.conn:
            self._clear_slot(driver_id, position, allow_draft)

    def clear_oldest(self, driver_id, allow_draft=False):
        with self.conn:
            self.driver(driver_id)
            row = self.conn.execute(
                'SELECT position FROM slots WHERE driver_id=? AND started_order IS NOT NULL ORDER BY started_order LIMIT 1', (driver_id,)
            ).fetchone()
            if row is None:
                return None
            self._clear_slot(driver_id, row['position'], allow_draft)
            return row['position']

    def history(self, driver_id=None, query='') -> list[dict]:
        sql, params = 'SELECT * FROM history', []
        if driver_id is not None:
            sql += ' WHERE driver_id=?'
            params.append(driver_id)
        rows = [dict(r) for r in self.conn.execute(sql + ' ORDER BY id DESC', params)]
        if query.strip():
            needle = query.strip().casefold()
            rows = [r for r in rows if needle in ' '.join(str(v) for v in r.values()).casefold()]
        return rows

    def restore(self, history_id, position=None):
        with self.conn:
            self.conn.execute('BEGIN IMMEDIATE')
            row = self.conn.execute('SELECT * FROM history WHERE id=?', (history_id,)).fetchone()
            if row is None:
                raise ValueError('Nie znaleziono wpisu historycznego.')
            if position is None:
                position = self._add_slot(row['driver_id'])
            self._slot(row['driver_id'], position)
            if self.slots(row['driver_id'])[position - 1]['status'] != 'Puste':
                raise ValueError('Wybrane miejsce jest zajęte. Najpierw je zwolnij.')
            self._save_slot(row['driver_id'], position, row['trailer'], row['info'], row['load_date'], row['unload_date'],
                            route=row['route'], order_number=row['order_number'])
            return position

    def assign_truck(self, driver_id, truck_id):
        self._truck(truck_id)
        with self.conn:
            self.driver(driver_id)
            self.conn.execute('UPDATE drivers SET truck_id=? WHERE id=?', (truck_id, driver_id))
            for data in self.slots(driver_id):
                self._archive(data)

    def rename_driver(self, driver_id, name):
        name = name.strip()
        if not name:
            raise ValueError('Podaj imię i nazwisko kierowcy.')
        with self.conn:
            self.driver(driver_id)
            self.conn.execute('UPDATE drivers SET name=? WHERE id=?', (name, driver_id))
            for data in self.slots(driver_id):
                self._archive(data)

    def set_driver_active(self, driver_id, active):
        with self.conn:
            self.driver(driver_id)
            self.conn.execute('UPDATE drivers SET active=? WHERE id=?', (bool(active), driver_id))

    def set_vehicle_active(self, vehicle_id, active):
        with self.conn:
            if not active and self.conn.execute('SELECT 1 FROM drivers WHERE active=1 AND truck_id=?', (vehicle_id,)).fetchone():
                raise ValueError('Ten tir jest przypisany do aktywnego kierowcy. Najpierw zmień przypisanie.')
            self.conn.execute('UPDATE vehicles SET active=? WHERE id=?', (bool(active), vehicle_id))

    def backup(self, target):
        target = Path(target).resolve()
        if target == self.path:
            raise ValueError('Wybierz inny plik niż używana baza.')
        target.parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(target) as destination:
            self.conn.backup(destination)
