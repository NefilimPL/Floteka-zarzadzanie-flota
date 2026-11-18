import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

from tools.release_version import parse_tag, prepare_version


@pytest.mark.parametrize('tag, version, numbers, dev, ref', [
    ('V1.0.0.0', '1.0.0.0', (1, 0, 0, 0), False, 'refs/tags/V1.0.0.0'),
    ('v12.3.4.5', '12.3.4.5', (12, 3, 4, 5), False, 'refs/tags/v12.3.4.5'),
    ('v1.0.0.0-dev', '1.0.0.0-dev', (1, 0, 0, 0), True, 'dev'),
    ('V0.65535.0.1-dev', '0.65535.0.1-dev', (0, 65535, 0, 1), True, 'dev'),
])
def test_tag_selects_version_windows_numbers_and_source(tag, version, numbers, dev, ref):
    result = parse_tag(tag)
    assert result.version == version
    assert result.windows_version == numbers
    assert result.prerelease is dev
    assert result.source_ref == ref


@pytest.mark.parametrize('tag', [
    '1.0.0.0', 'v1.0.0', 'v1.0.0.0.1', 'v1.0.0.0-beta', 'v1.0.0.0-dev-extra',
    'v65536.0.0.0', 'v1.-1.0.0', 'v01.0.0.0', 'v1.0.0.0\n', 'v1.0.0.0;echo bad',
])
def test_invalid_tags_are_rejected(tag):
    with pytest.raises(ValueError):
        parse_tag(tag)


@pytest.fixture
def source_copy(tmp_path):
    root = Path(__file__).resolve().parents[1]
    shutil.copytree(root / 'fleet', tmp_path / 'fleet', ignore=shutil.ignore_patterns('__pycache__', '_build_version.py'))
    shutil.copy(root / 'main.py', tmp_path / 'main.py')
    return tmp_path


@pytest.mark.parametrize('tag, expected', [('V2.3.4.5', '2.3.4.5'), ('v2.3.4.5-dev', '2.3.4.5-dev')])
def test_prepared_version_reaches_cli_windows_resource_and_excel(source_copy, tag, expected):
    prepare_version(source_copy, tag)
    result = subprocess.run([sys.executable, 'main.py', '--version'], cwd=source_copy,
                            capture_output=True, text=True, check=True)
    assert result.stdout.strip() == 'Floteka ' + expected
    # Evaluate the real PyInstaller version resource, including numeric fields.
    from PyInstaller.utils.win32.versioninfo import load_version_info_from_text_file
    resource = load_version_info_from_text_file(str(source_copy / 'build' / 'windows-version.txt'))
    assert (resource.ffi.fileVersionMS, resource.ffi.fileVersionLS) == (131075, 262149)
    assert (resource.ffi.productVersionMS, resource.ffi.productVersionLS) == (131075, 262149)
    strings = resource.kids[0].kids[0].kids
    assert {item.name: item.val for item in strings}['ProductVersion'] == expected
    assert bool(resource.ffi.fileFlags & 2) == tag.endswith('-dev')
    metadata = json.loads((source_copy / 'build' / 'version.json').read_text(encoding='utf-8'))
    assert metadata['version'] == expected
    assert metadata['tag'] == tag
    script = (
        "from fleet.store import Store; from fleet.export import export_xlsx; "
        "store=Store('probe.sqlite3'); export_xlsx(store,'probe.xlsx'); store.close()"
    )
    subprocess.run([sys.executable, '-c', script], cwd=source_copy, check=True)
    book = load_workbook(source_copy / 'probe.xlsx')
    try:
        assert book.custom_doc_props['FlotekaVersion'].value == expected
        assert book.properties.version == expected
        assert book['DANE'].oddFooter.left.text == 'Floteka ' + expected
        assert book.sheetnames == ['DANE', 'ARCHIWUM']
    finally:
        book.close()


def test_new_build_replaces_cached_version_and_untagged_build_resets_it(source_copy):
    for tag, expected in [('v1.2.3.4-dev', '1.2.3.4-dev'), ('v1.2.3.5-dev', '1.2.3.5-dev'), (None, '1.1.2')]:
        prepare_version(source_copy, tag)
        result = subprocess.run([sys.executable, 'main.py', '--version'], cwd=source_copy,
                                capture_output=True, text=True, check=True)
        assert result.stdout.strip() == 'Floteka ' + expected


def test_invalid_tag_leaves_prepared_version_untouched(source_copy):
    prepare_version(source_copy, 'v1.2.3.4')
    previous = (source_copy / 'fleet' / '_build_version.py').read_bytes()
    with pytest.raises(ValueError):
        prepare_version(source_copy, 'v1.2.3')
    assert (source_copy / 'fleet' / '_build_version.py').read_bytes() == previous


def test_source_archive_build_does_not_require_git(source_copy, monkeypatch):
    monkeypatch.setenv('PATH', '')
    prepare_version(source_copy, 'v1.2.3.4')
    metadata = json.loads((source_copy / 'build' / 'version.json').read_text(encoding='utf-8'))
    assert metadata['commit'] is None


def test_ci_build_requires_a_verified_source_commit(source_copy):
    with pytest.raises(ValueError, match='commit'):
        prepare_version(source_copy, 'v1.2.3.4', require_commit=True)
    assert not (source_copy / 'fleet' / '_build_version.py').exists()
