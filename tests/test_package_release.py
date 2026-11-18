import json
from pathlib import Path
import shutil
from zipfile import ZipFile

from tools import package_release


def test_public_package_excludes_individual_workbooks_and_historical_reports(tmp_path, monkeypatch):
    root = tmp_path / 'sources'
    shutil.copytree(Path(__file__).resolve().parents[1], root,
                    ignore=shutil.ignore_patterns('.git', 'build', 'dist', 'local-archive',
                                                 '__pycache__', '.venv', '*.json', '*.sqlite3', '*.xlsm', '*.exe', '*.zip'))
    reports = ['excel-test.json', 'excel-regressions.json', 'excel-deadlines.json',
               'excel-optional.json', 'excel-export.json', 'exe-test.json']
    for name in reports:
        (root / 'docs' / name).write_text(json.dumps({'ok': True, 'frozen_export_checked': True}), encoding='utf-8')
    (root / 'docs' / 'excel-migration.json').write_text('private migration report', encoding='utf-8')
    (root / 'docs' / 'ci-local-check.json').write_text('outdated local report', encoding='utf-8')
    planning = root / 'docs' / 'superpowers' / 'plans'
    planning.mkdir(parents=True, exist_ok=True)
    (planning / 'completed.md').write_text('outdated implementation notes', encoding='utf-8')
    output = tmp_path / 'release'
    output.mkdir()
    for name in ('Floteka.exe', 'Floteka.xlsm', 'Floteka-DEMO.xlsm', 'Floteka-DEMO.sqlite3'):
        (output / name).write_bytes(b'public artifact')
    private = output / 'Floteka-TWOJE-DANE.xlsm'
    private.write_bytes(b'private workbook')
    monkeypatch.setattr(package_release, 'ROOT', root)
    monkeypatch.setattr(package_release, 'OUT', output)

    package_release.package()

    with ZipFile(output / 'Floteka-pakiet.zip') as archive:
        assert 'Floteka/Floteka.xlsm' in archive.namelist()
        assert 'Floteka/Floteka-TWOJE-DANE.xlsm' not in archive.namelist()
    with ZipFile(output / 'Zrodla.zip') as archive:
        assert 'README.md' in archive.namelist()
        assert 'docs/excel-test.json' in archive.namelist()
        assert 'docs/excel-migration.json' not in archive.namelist()
        assert 'docs/ci-local-check.json' not in archive.namelist()
        assert not any(name.startswith('docs/superpowers/') for name in archive.namelist())
    assert private.read_bytes() == b'private workbook'
