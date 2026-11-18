"""Kontrola gotowego pliku, w tym faktycznych kotwic przycisków OOXML."""
import os
import posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import pytest
from openpyxl import load_workbook
from tools.excel_paths import default_excel_file

ROOT = Path(__file__).resolve().parents[1]
FILE = default_excel_file()
pytestmark = pytest.mark.skipif(not FILE.exists(), reason='Zbuduj XLSM poleceniem python tools/build_excel.py')
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'p': 'http://schemas.openxmlformats.org/package/2006/relationships',
      'd': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
HEADERS = ['Kierowca', 'TIR — rejestracja', 'Przyczepa — rejestracja', 'Informacje dodatkowe', 'Data załadunku', 'Data rozładunku', 'Status', 'Wyczyść']


def test_all_headers_are_in_the_correct_columns():
    wb = load_workbook(FILE, keep_vba=True)
    try:
        assert [wb['DANE'].cell(5, col).value for col in range(1, 9)] == HEADERS
        assert [wb['KIEROWCY'].cell(5, col).value for col in range(1, 5)] == ['ID', 'Kierowca', 'TIR', 'Aktywny']
        assert [wb['ARCHIWUM'].cell(5, col).value for col in range(1, 7)] == HEADERS[:6]
        for name in ('DANE', 'ARCHIWUM'):
            assert [wb[name][cell].value for cell in ('K5', 'L5')] == ['Trasa', 'Numer zlecenia']
            assert wb[name].max_row < 10_000
        assert wb['ARCHIWUM'].column_dimensions['J'].hidden
        assert wb['KIEROWCY']['E5'].value == 'Telefon'
        assert [cell.value for cell in wb['TIRY'][5][:6]] == [
            'Rejestracja', 'Przebieg (km)', 'Data przeglądu', 'Data ubezpieczenia', 'Data serwisu', 'Uwagi',
        ]
        assert wb['PRZYCZEPY']['B5'].value is None
    finally:
        wb.close()


def test_toolbar_buttons_fit_inside_their_own_row():
    with ZipFile(FILE) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        relationships = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        mapping = {r.attrib['Id']: r.attrib['Target'] for r in relationships}
        node = next(s for s in workbook.find('s:sheets', NS) if s.attrib['name'] == 'DANE')
        sheet_path = posixpath.normpath('xl/' + mapping[node.attrib['{' + NS['r'] + '}id']])
        sheet = ET.fromstring(archive.read(sheet_path))
        drawing = sheet.find('s:drawing', NS)
        rel_path = posixpath.dirname(sheet_path) + '/_rels/' + posixpath.basename(sheet_path) + '.rels'
        rels = ET.fromstring(archive.read(rel_path))
        target = next(r.attrib['Target'] for r in rels if r.attrib['Id'] == drawing.attrib['{' + NS['r'] + '}id'])
        path = posixpath.normpath(posixpath.dirname(sheet_path) + '/' + target)
        anchors = ET.fromstring(archive.read(path))
        toolbar = [a for a in anchors if a.find('d:sp', NS) is not None and a.find('d:sp', NS).attrib.get('macro') in ('DodajKierowce', 'ZwolnijNajstarsze', 'ZapiszWszystko', 'DodajPrzyczepe')]
        assert len(toolbar) == 4
        for anchor in toolbar:
            assert int(anchor.find('d:from/d:row', NS).text) == 2, 'Przycisk nachodzi na opis w wierszu 2'
            assert int(anchor.find('d:to/d:row', NS).text) == 2, 'Przycisk wystaje poza własny wiersz'
