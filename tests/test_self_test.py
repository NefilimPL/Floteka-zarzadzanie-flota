import json
import subprocess
import sys

from openpyxl import load_workbook

from main import self_test


def test_default_self_test_does_not_overwrite_sibling_workbook(tmp_path):
    report = tmp_path / 'Raport.json'
    sibling = tmp_path / 'Raport.xlsx'
    sibling.write_bytes(b'istniejacy skoroszyt uzytkownika')
    self_test(report)
    assert sibling.read_bytes() == b'istniejacy skoroszyt uzytkownika'
    assert json.loads(report.read_text(encoding='utf-8'))['ok']


def test_explicit_self_test_export_uses_only_requested_path(tmp_path):
    report = tmp_path / 'Raport.json'
    sibling = tmp_path / 'Raport.xlsx'
    sibling.write_bytes(b'nie zmieniaj tego pliku')
    exported = tmp_path / 'test-only.xlsx'
    result = subprocess.run([sys.executable, 'main.py', '--self-test', str(report),
                             '--self-test-export', str(exported)], capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stderr
    assert sibling.read_bytes() == b'nie zmieniaj tego pliku'
    workbook = load_workbook(exported)
    try:
        assert len(workbook['DANE'].tables) == 1
        assert workbook['DANE'].auto_filter.ref is None
    finally:
        workbook.close()


def test_report_and_explicit_export_cannot_be_the_same_file(tmp_path):
    path = tmp_path / 'same.xlsx'
    path.write_bytes(b'zachowaj dane')
    result = subprocess.run([sys.executable, 'main.py', '--self-test', str(path),
                             '--self-test-export', str(path)], capture_output=True, text=True, timeout=30)
    assert result.returncode == 2
    assert path.read_bytes() == b'zachowaj dane'
