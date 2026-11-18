"""Generate real macro workbooks without importing any Office automation."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from zipfile import ZipFile

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
NAMES = ('Floteka.xlsm', 'Floteka-DEMO.xlsm')


@pytest.fixture
def source(tmp_path, monkeypatch):
    # These cases simulate a local source ZIP, including subprocess generation.
    # Do not inherit the host runner's requirement for a verified Git commit.
    monkeypatch.delenv('GITHUB_ACTIONS', raising=False)
    root = tmp_path / 'sources with spaces'
    shutil.copytree(ROOT / 'excel', root / 'excel')
    shutil.copytree(ROOT / 'fleet', root / 'fleet',
                    ignore=shutil.ignore_patterns('__pycache__', '_build_version.py'))
    (root / 'tools').mkdir()
    shutil.copyfile(ROOT / 'tools' / 'build_excel.py', root / 'tools' / 'build_excel.py')
    return root


def test_generator_can_be_imported_without_office():
    script = '''
import sys
class NoOffice:
    def find_spec(self, fullname, *args):
        if fullname.split('.')[0] in ('winreg', 'win32com', 'pythoncom'):
            raise ImportError('Office automation is unavailable')
sys.meta_path.insert(0, NoOffice())
from tools.generate_excel import generate
assert callable(generate)
'''
    result = subprocess.run([sys.executable, '-c', script], cwd=ROOT,
                            capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


@pytest.mark.parametrize('tag, version', [('V2.3.4.5', '2.3.4.5'), ('v2.3.4.6-dev', '2.3.4.6-dev')])
def test_generated_files_keep_all_macros_and_layout(source, tmp_path, tag, version):
    from tools.generate_excel import generate
    output = tmp_path / 'output with spaces'
    report = generate(output, tag, root=source)
    assert report['ok'] and report['engine'] == 'template'
    assert report['macro_tests_run'] is False
    assert report['version'] == version and report['tag'] == tag
    assert report['commit'] is None  # A downloaded source ZIP has no .git.
    assert json.loads((output / 'excel-build.json').read_text(encoding='utf-8')) == report
    for name in NAMES:
        target = output / name
        assert report['hashes'][name] == hashlib.sha256(target.read_bytes()).hexdigest()
        with ZipFile(source / 'excel' / 'templates' / name) as before, ZipFile(target) as after:
            assert before.namelist() == after.namelist()
            assert after.testzip() is None
            for member in before.namelist():
                if member not in ('docProps/core.xml', 'docProps/custom.xml', 'xl/sharedStrings.xml') and not member.startswith('xl/worksheets/sheet'):
                    assert before.read(member) == after.read(member), member
            assert before.read('xl/vbaProject.bin') == after.read('xl/vbaProject.bin')
        book = load_workbook(target, keep_vba=True)
        try:
            assert book.custom_doc_props['FlotekaVersion'].value == version
            assert f'Wersja {version} |' in book['START']['A2'].value
            assert book['DANE'].oddFooter.left.text == f'Floteka {version}'
            assert book['_STAN'].sheet_state == 'veryHidden'
            assert book['DANE']['K5'].value == 'Trasa'
            assert book['ARCHIWUM']['L5'].value == 'Numer zlecenia'
            if name == 'Floteka.xlsm':
                assert all(value is None for sheet in ('DANE', 'ARCHIWUM')
                           for row in book[sheet].iter_rows(min_row=6, values_only=True) for value in row)
                assert book['KIEROWCY']['B6'].value is None
            else:
                assert book['KIEROWCY']['B6'].value == 'Kierowca pokazowy A'
                assert book['KIEROWCY']['B7'].value == 'Kierowca pokazowy B'
        finally:
            book.close()


@pytest.mark.parametrize('changed', ['excel/Flota.bas', 'excel/ThisWorkbook.cls', 'tools/build_excel.py', 'excel/templates/Floteka.xlsm'])
def test_stale_sources_or_modified_templates_fail_before_output(source, tmp_path, changed):
    from tools.generate_excel import generate
    with (source / changed).open('ab') as file:
        file.write(b'\nchanged')
    output = tmp_path / 'must not be created'
    with pytest.raises(ValueError, match='(?i)template|szablon'):
        generate(output, 'V2.3.4.5', root=source)
    assert not output.exists()


def test_existing_workbook_is_never_replaced(source, tmp_path):
    from tools.generate_excel import generate
    output = tmp_path / 'existing'
    output.mkdir()
    original = b'user workbook and history'
    (output / 'Floteka.xlsm').write_bytes(original)
    with pytest.raises(FileExistsError):
        generate(output, 'V2.3.4.5', root=source)
    assert (output / 'Floteka.xlsm').read_bytes() == original
    assert not (output / 'Floteka-DEMO.xlsm').exists()


def test_generation_runs_with_office_imports_blocked(source, tmp_path):
    script = '''
import sys
from pathlib import Path
class NoOffice:
    def find_spec(self, fullname, *args):
        if fullname.split('.')[0] in ('winreg', 'win32com', 'pythoncom'):
            raise ImportError('Office automation is unavailable')
sys.meta_path.insert(0, NoOffice())
from tools.generate_excel import generate
generate(Path(sys.argv[1]), 'V2.3.4.5', root=Path(sys.argv[2]))
'''
    output = tmp_path / 'without Office'
    result = subprocess.run([sys.executable, '-c', script, str(output), str(source)],
                            cwd=ROOT, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert all((output / name).exists() for name in NAMES)


def test_ci_generation_requires_a_verified_source_commit(source, tmp_path, monkeypatch):
    from tools.generate_excel import generate
    monkeypatch.setenv('GITHUB_ACTIONS', 'true')
    output = tmp_path / 'unverified CI build'
    with pytest.raises(ValueError, match='verified source commit'):
        generate(output, 'V2.3.4.5', root=source)
    assert not output.exists()
    assert not (source / 'fleet' / '_build_version.py').exists()
