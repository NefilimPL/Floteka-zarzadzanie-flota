import subprocess
import sys

from main import default_data_folder


def test_default_database_uses_floteka_folder(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    assert default_data_folder() / 'flota.sqlite3' == tmp_path / 'Floteka' / 'flota.sqlite3'


def test_command_line_reports_floteka_version():
    result = subprocess.run([sys.executable, 'main.py', '--version'], capture_output=True, text=True)
    assert result.returncode == 0
    assert result.stdout.strip() == 'Floteka 1.1.2'
