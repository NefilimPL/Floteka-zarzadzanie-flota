"""Wybór skoroszytu Floteki do testów."""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def default_excel_file():
    override = os.environ.get('FLOTA_EXCEL_FILE') or os.environ.get('FLOTA_LAYOUT_FILE')
    if override:
        return Path(override).resolve()
    return ROOT / 'dist' / 'Floteka' / 'Floteka.xlsm'
