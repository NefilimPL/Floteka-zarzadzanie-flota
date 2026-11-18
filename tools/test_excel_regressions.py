"""Rzeczywiste scenariusze awarii; modyfikuje wyłącznie kopie w build/."""
from contextlib import contextmanager
import os
import json
import shutil
import threading
import time
from pathlib import Path

import win32con
import win32gui
import win32process
import win32com.client

from build_excel import ROOT, vba_access
from excel_paths import default_excel_file


@contextmanager
def acknowledge_test_dialog(app):
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    stop = threading.Event()
    seen = []
    def worker():
        while not stop.wait(0.05):
            def visit(hwnd, _):
                if win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.GetClassName(hwnd) == '#32770' and win32gui.GetWindowText(hwnd) in ('Floteka', 'Flota Adriana'):
                    seen.append(hwnd)
                    def click_ok(child, __):
                        if win32gui.GetWindowText(child) == 'OK':
                            win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                    win32gui.EnumChildWindows(hwnd, click_ok, None)
            win32gui.EnumWindows(visit, None)
    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    try:
        yield seen
    finally:
        stop.set()
        thread.join(2)


def run():
    results, failures = [], []
    with vba_access():
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 1
        try:
            for test in ('failed_driver_creation', 'multirow_registry_paste', 'invalid_dates_restore_previous', 'missing_history_reconciled', 'missing_key_deduplicated'):
                path = ROOT / 'build' / f'{test}.xlsm'
                source = default_excel_file()
                shutil.copy2(source, path)
                book = app.Workbooks.Open(str(path))
                try:
                    macro = lambda name, *args: app.Run(f"'{book.Name}'!{name}", *args)
                    module = book.VBProject.VBComponents.Add(1)
                    module.Name = 'TestAutomation'
                    module.CodeModule.AddFromString('''Public Function TryAdd() As String
On Error GoTo Failed
TryAdd = "OK:" & CStr(Flota.AddDriverRecord("Przerwany", "WX999"))
Exit Function
Failed:
TryAdd = "ERROR:" & Err.Description
End Function''')
                    data, reg = book.Worksheets('DANE'), book.Worksheets('KIEROWCY')
                    if test == 'failed_driver_creation':
                        book.Names('ListaPrzyczep').Delete()
                        result = macro('TryAdd')
                        assert str(result).startswith('ERROR:'), result
                        assert reg.Cells(6, 1).Value in (None, ''), 'Częściowy kierowca pozostał w rejestrze'
                        assert data.Cells(6, 1).Value in (None, ''), 'Pozostał niepełny blok DANE'
                        book.Names.Add('ListaPrzyczep', "='PRZYCZEPY'!$A$6:$A$6")
                        driver = int(macro('AddDriverRecord', 'Poprawny', 'WX123'))
                        assert int(macro('DriverFirstRow', driver)) == 6
                        assert data.Cells(9, 1).Value == 'Poprawny'
                    elif test == 'multirow_registry_paste':
                        macro('AddDriverRecord', 'Pierwszy', 'WX100')
                        macro('AddDriverRecord', 'Drugi', 'WX200')
                        data.Range('C10:F10').Value = (('PO100', 'Firma', '17.11.2026', '18.11.2026'),)
                        with acknowledge_test_dialog(app) as dialogs:
                            reg.Range('C6:C7').Value = (('',), ('PO777',))
                        assert dialogs, 'Nie zgłoszono pustego tira pierwszego kierowcy'
                        assert reg.Cells(6, 3).Value == 'WX100'
                        assert reg.Cells(7, 3).Value == 'PO777'
                        assert data.Cells(10, 2).Value == 'PO777', 'Rejestr i DANE są niespójne po wklejeniu'
                        assert macro('HistoryCount') == 2, 'Brak kopii z nowym tirem'
                    elif test == 'invalid_dates_restore_previous':
                        macro('AddDriverRecord', 'Pierwszy', 'WX100')
                        data.Range('C6:F6').Value = (('PO100', 'Firma', '17.11.2026', '18.11.2026'),)
                        before = data.Range('C6:F6').Value2
                        with acknowledge_test_dialog(app) as dialogs:
                            data.Cells(6, 5).Value = '31.02.2026'
                        assert dialogs
                        assert data.Range('C6:F6').Value2 == before
                        assert macro('HistoryCount') == 1
                    elif test == 'missing_history_reconciled':
                        macro('AddDriverRecord', 'Pierwszy', 'WX100')
                        data.Range('C6:F6').Value = (('PO100', 'Firma', '17.11.2026', '18.11.2026'),)
                        archive = book.Worksheets('ARCHIWUM')
                        archive.Unprotect('flota')
                        archive.Range('A6:J6').ClearContents()  # Wyłącznie testowa kopia: symulacja dawnego błędu.
                        macro('Initialize')
                        assert macro('HistoryCount') == 0
                        assert macro('FlushAll')
                        assert macro('HistoryCount') == 1, 'Niezmieniony kompletny wpis nie odzyskał brakującej kopii historii'
                    else:
                        macro('AddDriverRecord', 'Pierwszy', 'WX100')
                        data.Range('C6:F6').Value = (('PO100', 'Firma', '17.11.2026', '18.11.2026'),)
                        archive = book.Worksheets('ARCHIWUM')
                        archive.Unprotect('flota')
                        archive.Range('J6').ClearContents()
                        macro('Initialize')
                        assert macro('FlushAll')
                        assert macro('HistoryCount') == 1, 'Brak technicznego klucza utworzył duplikat pełnych danych'
                    results.append(test)
                    print('PASS', test, flush=True)
                except Exception as exc:
                    failures.append(f'{test}: {exc}')
                    print('FAIL', failures[-1], flush=True)
                finally:
                    app.EnableEvents = False
                    book.Close(False)
                    app.EnableEvents = True
        finally:
            app.Quit()
    report = {'ok': not failures, 'checks': results, 'failures': failures}
    (ROOT / 'docs' / 'excel-regressions.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    assert not failures, failures


if __name__ == '__main__':
    run()
