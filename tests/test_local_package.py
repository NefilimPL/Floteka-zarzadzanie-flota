"""Local delivery rejects stale evidence and never includes private files."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
from zipfile import ZipFile

import pytest

from tools import package_local


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.fixture
def local_build(tmp_path):
    root = tmp_path / 'source with spaces'
    names = ['main.py', 'README.md', 'requirements.txt', 'requirements-build.txt',
             'Buduj.bat', 'Uruchom.bat', '.gitignore', 'docs/verification.md',
             'fleet/branding.py', 'tools/build_local.ps1', 'tools/package_local.py',
             'tools/generate_excel.py', 'tools/refresh_excel_templates.py',
             'tools/build_exe.ps1', 'tools/release_version.py', 'tools/make_demo.py',
             'excel/Floteka.bas', 'excel/templates/Floteka.xlsm',
             'excel/templates/Floteka-DEMO.xlsm', 'excel/templates/manifest.json']
    for name in names:
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text('public source', encoding='utf-8')
    output = tmp_path / 'release with spaces'
    output.mkdir()
    for name in ['Floteka.exe', 'Floteka.xlsm', 'Floteka-DEMO.xlsm', 'Floteka-DEMO.sqlite3']:
        (output / name).write_bytes(('generated ' + name).encode())
    reports = {
        'exe-test.json': {'ok': True, 'frozen': True, 'app': 'Floteka', 'version': '1.2.3.4-dev'},
        'build-info.json': {'app': 'Floteka', 'version': '1.2.3.4-dev', 'tag': 'v1.2.3.4-dev',
                            'commit': 'a' * 40, 'sha256': digest(output / 'Floteka.exe'), 'self_test': True},
        'excel-build.json': {'ok': True, 'version': '1.2.3.4-dev', 'tag': 'v1.2.3.4-dev',
                             'commit': 'a' * 40, 'engine': 'template', 'macro_tests_run': False,
                             'files': ['Floteka.xlsm', 'Floteka-DEMO.xlsm'],
                             'hashes': {name: digest(output / name) for name in ['Floteka.xlsm', 'Floteka-DEMO.xlsm']}},
    }
    for name, value in reports.items():
        (output / name).write_text(json.dumps(value), encoding='utf-8-sig')
    return root, output


def test_local_package_contains_verified_artifacts_and_rebuildable_sources(local_build):
    root, output = local_build
    for base in [root, output]:
        (base / 'Floteka-TWOJE-DANE.xlsm').write_bytes(b'private workbook')
        (base / 'flota.sqlite3').write_bytes(b'private database')
    (root / 'docs/excel-migration.json').write_text('private report')
    (root / 'tools/password.txt').write_text('private secret')
    package_local.package(output, root=root)
    with ZipFile(output / 'Floteka-pakiet.zip') as archive:
        assert 'Floteka/Floteka.exe' in archive.namelist()
        assert 'Floteka/Floteka.xlsm' in archive.namelist()
        assert 'Floteka/Floteka-DEMO.xlsm' in archive.namelist()
        assert 'Floteka/Floteka-DEMO.sqlite3' in archive.namelist()
        assert 'Floteka/SUMY-SHA256.txt' in archive.namelist()
        assert 'Floteka/excel-build.json' in archive.namelist()
        report = json.loads(archive.read('Floteka/excel-build.json').decode('utf-8-sig'))
        assert report['macro_tests_run'] is False
        assert not any('TWOJE-DANE' in name or name.endswith('/flota.sqlite3') for name in archive.namelist())
    with ZipFile(output / 'Zrodla.zip') as archive:
        for name in ['Buduj.bat', 'tools/build_local.ps1', 'tools/generate_excel.py',
                     'tools/package_local.py', 'excel/templates/manifest.json', 'excel/templates/Floteka.xlsm']:
            assert name in archive.namelist()
        assert not any(name.endswith('.sqlite3') or name.startswith('build/') or 'migration' in name
                       or name.endswith('password.txt') for name in archive.namelist())
    for line in (output / 'SUMY-SHA256.txt').read_text().splitlines():
        checksum, name = line.split('  ')
        assert digest(output / name) == checksum
    assert (output / 'flota.sqlite3').read_bytes() == b'private database'


@pytest.mark.parametrize('report,field,value', [
    ('exe-test.json', 'ok', False), ('exe-test.json', 'frozen', False),
    ('exe-test.json', 'app', 'Other'), ('exe-test.json', 'version', '9.9.9.9'),
    ('build-info.json', 'app', 'Other'), ('build-info.json', 'self_test', False),
    ('build-info.json', 'version', '9.9.9.9'), ('build-info.json', 'tag', 'v9.9.9.9'),
    ('build-info.json', 'commit', 'b' * 40), ('build-info.json', 'sha256', '0' * 64),
    ('excel-build.json', 'ok', False), ('excel-build.json', 'engine', 'com'),
    ('excel-build.json', 'macro_tests_run', True), ('excel-build.json', 'hashes', {}),
    ('excel-build.json', 'files', ['Floteka.xlsm']),
])
def test_local_package_rejects_unverified_or_mismatched_reports(local_build, report, field, value):
    root, output = local_build
    path = output / report
    data = json.loads(path.read_text(encoding='utf-8-sig'))
    data[field] = value
    path.write_text(json.dumps(data), encoding='utf-8')
    with pytest.raises(ValueError):
        package_local.package(output, root=root)
    assert not (output / 'Floteka-pakiet.zip').exists()


@pytest.mark.parametrize('name', ['exe-test.json', 'build-info.json', 'excel-build.json', 'Floteka.exe', 'Floteka.xlsm'])
def test_local_package_rejects_missing_artifact_or_report(local_build, name):
    root, output = local_build
    (output / name).unlink()
    with pytest.raises((ValueError, FileNotFoundError)):
        package_local.package(output, root=root)
    assert not (output / 'Zrodla.zip').exists()


@pytest.mark.parametrize('name', ['Floteka.xlsm', 'Floteka-DEMO.xlsm', 'Floteka.exe'])
def test_local_package_rejects_changed_binary(local_build, name):
    root, output = local_build
    (output / name).write_bytes(b'modified after testing')
    with pytest.raises(ValueError):
        package_local.package(output, root=root)


def test_local_package_preserves_existing_package(local_build):
    root, output = local_build
    previous = output / 'Floteka-pakiet.zip'
    previous.write_bytes(b'previous delivery')
    with pytest.raises((ValueError, FileExistsError)):
        package_local.package(output, root=root)
    assert previous.read_bytes() == b'previous delivery'


def test_local_package_final_publication_failure_never_leaves_partial_zip(local_build, monkeypatch):
    root, output = local_build
    final_zip = output / 'Floteka-pakiet.zip'
    original_copy = package_local.shutil.copyfileobj
    original_rename = os.rename
    original_link = os.link

    def fail_copy(source, target, *args, **kwargs):
        if Path(target.name) == final_zip:
            target.write(source.read(32))
            raise OSError('injected disk full during final publication')
        return original_copy(source, target, *args, **kwargs)

    def fail_publish(operation):
        def publish(source, target, *args, **kwargs):
            if Path(target) == final_zip:
                raise OSError('injected disk full during final publication')
            return operation(source, target, *args, **kwargs)
        return publish

    monkeypatch.setattr(package_local.shutil, 'copyfileobj', fail_copy)
    monkeypatch.setattr(os, 'rename', fail_publish(original_rename))
    monkeypatch.setattr(os, 'link', fail_publish(original_link))
    with pytest.raises(OSError, match='injected disk full'):
        package_local.package(output, root=root)
    assert not final_zip.exists()


def test_local_package_atomic_publication_preserves_concurrent_destination(local_build, monkeypatch):
    root, output = local_build
    final_zip = output / 'Floteka-pakiet.zip'

    def create_destination_before_publish(operation):
        def publish(source, target, *args, **kwargs):
            if Path(target) == final_zip:
                final_zip.write_bytes(b'concurrent delivery')
            return operation(source, target, *args, **kwargs)
        return publish

    monkeypatch.setattr(os, 'rename', create_destination_before_publish(os.rename))
    monkeypatch.setattr(os, 'link', create_destination_before_publish(os.link))
    with pytest.raises(FileExistsError):
        package_local.package(output, root=root)
    assert final_zip.read_bytes() == b'concurrent delivery'


def test_local_build_rejects_nonempty_output_before_dependencies(tmp_path):
    output = tmp_path / 'user files'
    output.mkdir()
    sentinel = output / 'flota.sqlite3'
    sentinel.write_bytes(b'user data')
    script = Path(__file__).resolve().parents[1] / 'tools/build_local.ps1'
    result = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                             '-File', str(script), '-OutputDirectory', str(output)],
                            capture_output=True, text=True, timeout=30)
    assert result.returncode != 0
    assert 'Output directory must be empty' in result.stdout + result.stderr
    assert list(output.iterdir()) == [sentinel]
    assert sentinel.read_bytes() == b'user data'

