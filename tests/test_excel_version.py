"""Version checks for freshly built workbooks, separate from legacy layout checks."""
import os
from pathlib import Path

import pytest
from openpyxl import load_workbook

from fleet.branding import VERSION


FOLDER = os.environ.get('FLOTEKA_EXCEL_VERSION_DIR')
pytestmark = pytest.mark.skipif(not FOLDER, reason='Set FLOTEKA_EXCEL_VERSION_DIR to the new XLSM build directory.')


@pytest.mark.parametrize('name', ['Floteka.xlsm', 'Floteka-DEMO.xlsm'])
def test_built_workbook_contains_the_application_version(name):
    wb = load_workbook(Path(FOLDER) / name, keep_vba=True)
    try:
        assert wb.custom_doc_props['FlotekaVersion'].value == VERSION
        assert f'Wersja {VERSION}' in wb['START']['A2'].value
        assert wb['DANE'].oddFooter.left.text == f'Floteka {VERSION}'
    finally:
        wb.close()
