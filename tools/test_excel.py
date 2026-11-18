"""Scenariusz integracyjny — rzeczywisty Excel/VBA, wyłącznie kopia testowa."""

import json
import os
import shutil
import sys
from pathlib import Path

import win32com.client
from excel_paths import default_excel_file

ROOT = Path(__file__).resolve().parents[1]


def run(source=None):
    source = Path(source or default_excel_file()).resolve()
    test_file = ROOT / 'build' / 'excel-test.xlsm'
    test_file.parent.mkdir(exist_ok=True)
    shutil.copy2(source, test_file)
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 1  # Tylko nasz lokalnie wygenerowany skoroszyt testowy.
    book = None
    checks = []
    try:
        book = app.Workbooks.Open(str(test_file))
        macro = lambda name, *args: app.Run(f"'{book.Name}'!{name}", *args)
        driver = int(macro('AddDriverRecord', 'Test Żółć', 'WX100'))
        data, archive = book.Worksheets('DANE'), book.Worksheets('ARCHIWUM')
        first = int(macro('DriverFirstRow', driver))
        assert first == 6
        assert data.Cells(first + 3, 1).Value == 'Test Żółć'
        checks.append('cztery miejsca i polskie znaki')
        data.Cells(first, 3).Value = 'PO100'
        assert macro('HistoryCount') == 0
        assert data.Cells(first, 7).Value == 'Szkic'
        checks.append('szkic poza historia')
        data.Range(data.Cells(first, 4), data.Cells(first, 6)).Value = (('Firma', '17.11.2026', '20.11.2026'),)
        assert macro('HistoryCount') == 1
        data.Cells(first, 4).Value = 'Firma'
        assert macro('HistoryCount') == 1
        checks.append('pelny wpis archiwizowany od razu i bez duplikatow')
        data.Cells(first, 4).Value = 'Poprawiona firma'
        assert macro('HistoryCount') == 2
        assert archive.Cells(6, 4).Value == 'Firma'
        checks.append('poprawka zachowuje poprzednia wersje')
        macro('ClearRow', first, True)
        assert data.Cells(first, 3).Value in (None, '')
        assert macro('HistoryCount') == 2
        macro('RestoreRow', 6, first + 1)
        assert data.Cells(first + 1, 3).Value == 'PO100'
        assert macro('HistoryCount') == 2
        checks.append('czyszczenie i przywrocenie nie usuwa historii ani nie duplikuje')
        data.Range(data.Cells(first + 2, 3), data.Cells(first + 2, 6)).Value = (('PO200', '', '01.01.2026', '02.01.2026'),)
        assert macro('HistoryCount') == 3
        macro('ClearOldestForDriver', driver, True)
        assert data.Cells(first + 1, 3).Value in (None, '')
        assert data.Cells(first + 2, 3).Value == 'PO200'
        checks.append('najstarsze wedlug kolejnosci zajecia miejsca')
        reg = book.Worksheets('KIEROWCY')
        reg.Cells(6, 3).Value = 'KR999'
        assert data.Cells(first + 2, 2).Value == 'KR999'
        assert macro('HistoryCount') == 4
        assert archive.Cells(8, 2).Value == 'WX100'
        checks.append('zmiana tira zachowuje kopie')
        assert archive.ProtectContents
        archive.Unprotect('flota')
        archive.Protect('flota', True, True, True, False)
        try:
            app.EnableEvents = False
            archive.Cells(6, 4).Value = 'Niedozwolone'
            raise AssertionError('Ochrona historii nie dziala')
        except win32com.client.pywintypes.com_error:
            pass
        finally:
            app.EnableEvents = True
        macro('Initialize')
        checks.append('ochrona arkusza historii')
        book.Save()
        book.Close(False)
        book = app.Workbooks.Open(str(test_file))
        assert book.Worksheets('DANE').Cells(first + 2, 3).Value == 'PO200'
        assert app.Run(f"'{book.Name}'!HistoryCount") == 4
        checks.append('zapis i ponowne otwarcie')
        report = {'ok': True, 'checks': checks, 'excel': app.Version}
        (ROOT / 'docs' / 'excel-test.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(report, ensure_ascii=True))
    finally:
        if book:
            book.Close(False)
        app.Quit()


if __name__ == '__main__':
    run(sys.argv[1] if len(sys.argv) > 1 else None)
