"""Lokalizacja Floteki i jednorazowe przejęcie spójnej kopii dawnej bazy."""

import os
import sqlite3
import tempfile
from contextlib import closing
from pathlib import Path


def default_data_folder():
    return Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'Floteka'


def database_path(explicit=None):
    if explicit is not None:
        return Path(explicit)
    folder = default_data_folder()
    target = folder / 'flota.sqlite3'
    legacy = folder.parent / 'FlotaAdriana' / 'flota.sqlite3'
    if target.exists() or not legacy.is_file():
        return target
    folder.mkdir(parents=True, exist_ok=True)
    handle, name = tempfile.mkstemp(prefix='migracja-', suffix='.sqlite3', dir=folder)
    os.close(handle)
    temporary = Path(name)
    try:
        # backup obejmuje także zatwierdzone dane pozostające w WAL.
        with closing(sqlite3.connect(legacy.resolve().as_uri() + '?mode=ro', uri=True)) as source:
            with closing(sqlite3.connect(temporary)) as destination:
                source.backup(destination)
                if destination.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
                    raise sqlite3.DatabaseError('Nie udało się sprawdzić kopii dotychczasowej bazy.')
        try:
            # Atomowe utworzenie bez nadpisania bazy innego uruchomienia.
            os.link(temporary, target)
        except FileExistsError:
            pass
    finally:
        temporary.unlink(missing_ok=True)
    return target
