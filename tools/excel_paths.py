"""Wybór aktualnego pliku do testów; zachowana obsługa starszych nazw."""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def default_excel_file():
    override = os.environ.get('FLOTA_EXCEL_FILE') or os.environ.get('FLOTA_LAYOUT_FILE')
    if override:
        return Path(override).resolve()
    candidates = [ROOT / 'dist' / 'Floteka' / 'Floteka.xlsm',
                  ROOT / 'dist' / 'poprawione' / 'Flota-Adriana.xlsm',
                  ROOT / 'dist' / 'Flota-Adriana.xlsm']
    return next((p for p in candidates if p.exists()), candidates[0])
