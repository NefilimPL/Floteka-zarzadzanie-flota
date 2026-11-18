"""Pomiar i próby filtrów na syntetycznych danych, wyłącznie w kopii build/."""
import argparse
import json
import shutil
import statistics
import time
from pathlib import Path

import win32com.client

from build_excel import ROOT, vba_access

SEED = '''Public Function SeedPerformance() As String
On Error GoTo Failed
Dim d(1 To 1000, 1 To 10) As Variant, h(1 To 1000, 1 To 10) As Variant
Dim s(1 To 1000, 1 To 9) As Variant, r(1 To 250, 1 To 4) As Variant
Dim t(1 To 250, 1 To 1) As Variant, p(1 To 1000, 1 To 1) As Variant
Dim i As Long, j As Long, person As Long, key As String, part As String
Application.EnableEvents = False
For i = 1 To 1000
    person = ((i - 1) \\ 4) + 1
    d(i, 1) = "D" & person: d(i, 2) = "T" & person: d(i, 3) = "P" & i
    d(i, 4) = "Informacje " & i: d(i, 5) = CDbl(DateSerial(2026,11,17)): d(i, 6) = CDbl(DateSerial(2026,11,19))
    d(i, 7) = "Kompletny": d(i, 9) = person: d(i, 10) = ((i-1) Mod 4)+1
    key = CStr(person) & "|"
    For j = 1 To 6
        h(i,j) = d(i,j): part = CStr(d(i,j)): key = key & Len(part) & ":" & part
    Next j
    h(i,7) = CDbl(Now): h(i,8) = person: h(i,9) = i: h(i,10) = key
    s(i,1) = person: s(i,2) = d(i,10): s(i,3) = i
    For j = 3 To 6: s(i,j+1) = d(i,j): Next j
    s(i,8) = d(i,1): s(i,9) = d(i,2)
    p(i,1) = "P" & i
Next i
For i = 1 To 250
    r(i,1)=i: r(i,2)="D" & i: r(i,3)="T" & i: r(i,4)="TAK": t(i,1)="T" & i
Next i
With ThisWorkbook.Worksheets("DANE")
    .Unprotect "flota": .Range("A6:J1005").Value2 = d: .Range("C6:F1005").Locked=False
End With
ThisWorkbook.Worksheets("ARCHIWUM").Range("A6:J1005").Value2 = h
ThisWorkbook.Worksheets("_STAN").Range("A6:I1005").Value2 = s
ThisWorkbook.Worksheets("_STAN").Range("B1").Value2 = 250
ThisWorkbook.Worksheets("_STAN").Range("B2").Value2 = 1000
ThisWorkbook.Worksheets("KIEROWCY").Range("A6:D255").Value2 = r
ThisWorkbook.Worksheets("TIRY").Range("A6:A255").Value2 = t
ThisWorkbook.Worksheets("PRZYCZEPY").Range("A6:A1005").Value2 = p
Application.EnableEvents = True
Flota.Initialize
SeedPerformance="OK": Exit Function
Failed:
SeedPerformance="ERROR:" & Err.Description
Application.EnableEvents=True
End Function'''


def filter_state(item):
    result = [item.Operator, item.Criteria1]
    try:
        result.append(item.Criteria2)
    except Exception:
        result.append(None)
    return result


def run(source, baseline=False, trace=False):
    report_path = ROOT / 'docs' / ('excel-performance-before.json' if baseline else 'excel-performance-after.json')
    target = ROOT / 'build' / ('performance-before.xlsm' if baseline else 'performance-after.xlsm')
    shutil.copy2(source, target)
    with vba_access():
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 1
        book = None
        try:
            book = app.Workbooks.Open(str(target))
            if trace:
                code_module = book.VBProject.VBComponents('Flota').CodeModule
                code = code_module.Lines(1, code_module.CountOfLines).replace('\r\n', '\n')
                code = code.replace('    keys(key) = True\n    On Error Resume Next', '    arc.Cells(destination, 11).Value2 = "key=" & key & ";record=" & CStr(record(1, 10)) & ";range=" & arc.Cells(destination, 1).Resize(1, 10).Address & ";j=" & CStr(arc.Cells(destination, 10).Value2)\n    keys(key) = True\n    On Error Resume Next')
                code = code.replace('    RefreshArchiveFilter arc, destination', '    RefreshArchiveFilter arc, destination\n    arc.Cells(destination, 12).Value2 = "after=" & CStr(arc.Cells(destination, 10).Value2)')
                code_module.DeleteLines(1, code_module.CountOfLines)
                code_module.AddFromString(code)
            mod = book.VBProject.VBComponents.Add(1)
            mod.Name = 'PerformanceSeed'
            mod.CodeModule.AddFromString(SEED)
            macro = lambda name, *args: app.Run(f"'{book.Name}'!{name}", *args)
            assert macro('SeedPerformance') == 'OK'
            print('Seed: 250 kierowcow, 1000 wpisow, 1000 historii', flush=True)
            history_original = book.Worksheets('ARCHIWUM').Range('A6:J1005').Value2
            times = []
            for _ in range(3):
                start = time.perf_counter()
                book.Save()
                elapsed = time.perf_counter() - start
                times.append(elapsed)
                print('Save seconds:', round(elapsed, 4), 'history:', macro('HistoryCount'), flush=True)
            arc, data = book.Worksheets('ARCHIWUM'), book.Worksheets('DANE')
            data.Unprotect('flota')
            data.Range('A5:J1005').AutoFilter(1, 'D1')
            arc.Unprotect('flota')
            arc.Range('A5:I1005').AutoFilter(3, 'P1')
            macro('Initialize')
            assert arc.FilterMode and data.FilterMode
            print('Filtered, history:', macro('HistoryCount'), flush=True)
            book.Save()
            filtered_count = int(macro('HistoryCount'))
            print('After filtered save:', filtered_count, flush=True)
            book.Close(False)
            book = app.Workbooks.Open(str(target))
            arc, data = book.Worksheets('ARCHIWUM'), book.Worksheets('DANE')
            assert macro('FlushAll')
            reopened_count = int(macro('HistoryCount'))
            print('After reopen/flush:', reopened_count, flush=True)
            start = time.perf_counter()
            data.Cells(6, 4).Value = 'Nowa wersja pod filtrem'
            edit_seconds = time.perf_counter() - start
            edited_count = int(macro('HistoryCount'))
            print('New history key present:', bool(arc.Cells(1006, 10).Value2), flush=True)
            if trace:
                print('TRACE', arc.Range('J1006:L1006').Value2, flush=True)
            history_unchanged = arc.Range('A6:J1005').Value2 == history_original
            filter_kept = bool(arc.FilterMode and arc.AutoFilter.Filters(3).On)
            if filter_kept:
                assert str(arc.AutoFilter.Filters(3).Criteria1).lstrip('=') == 'P1'
            extra_tests = []
            if not baseline:
                arc.Range('A5:I1006').AutoFilter(3, 'BRAK_PASUJACYCH')
                macro('Initialize')
                book.Save()
                assert macro('HistoryCount') == 1001, f"History after empty-filter save: {macro('HistoryCount')}; key: {arc.Cells(1006, 10).Value2}"
                data.Cells(7, 4).Value = 'Poprawka przy pustym wyniku filtra'
                assert macro('HistoryCount') == 1002
                assert str(arc.AutoFilter.Filters(3).Criteria1).lstrip('=') == 'BRAK_PASUJACYCH'
                assert arc.Range('A6:J1005').Value2 == history_original
                extra_tests.append('zero visible rows: no overwrite or duplicates')
                driver = int(macro('AddDriverRecord', 'Nowy pod filtrem', 'T1'))
                assert int(macro('DriverFirstRow', driver)) == 1006
                assert data.Cells(1009, 1).Value == 'Nowy pod filtrem'
                extra_tests.append('driver added after all hidden existing slots')
                arc.Range('A5:I1007').AutoFilter(3, ('P1', 'P2'), 7)
                multiple_before = filter_state(arc.AutoFilter.Filters(3))
                assert not arc.Rows(6).Hidden and not arc.Rows(7).Hidden and arc.Rows(8).Hidden
                data.Cells(7, 4).Value = 'Poprawka przy wielokrotnym wyborze'
                assert macro('HistoryCount') == 1003
                assert filter_state(arc.AutoFilter.Filters(3)) == multiple_before
                extra_tests.append('multiple selected filter values retained')
                arc.Range('A5:I1008').AutoFilter(5, '>=46343', 1, '<=46345')
                data.Cells(9, 4).Value = 'Poprawka przy zakresie dat'
                assert macro('HistoryCount') == 1004
                assert arc.AutoFilter.Filters(5).Criteria1 == '>=46343'
                assert arc.AutoFilter.Filters(5).Criteria2 == '<=46345'
                assert arc.Range('A6:J1005').Value2 == history_original
                actual_keys = [row[0] for row in arc.Range('J6:J1009').Value2]
                assert len(set(actual_keys)) == 1004
                extra_tests.append('date range retained; every key unique; original history unchanged')
            no_duplicates = filtered_count == reopened_count == 1000 and edited_count == 1001 and history_unchanged
            result = {'ok': filter_kept and no_duplicates, 'drivers': 250, 'visible_slots': 1000, 'history_initial': 1000,
                      'save_seconds': times, 'save_median_seconds': statistics.median(times), 'edit_seconds': edit_seconds,
                      'no_duplicates_after_filter_save_reopen': no_duplicates, 'filter_preserved_after_append': filter_kept,
                      'history_after_edit': edited_count, 'excel': app.Version, 'extra_tests': extra_tests}
            result['all_previous_history_cells_unchanged'] = history_unchanged
            report_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
            print(json.dumps(result), flush=True)
            if not baseline:
                assert filter_kept, 'Dopisanie historii skasowalo filtr'
                assert no_duplicates, 'Wykryto dodatkowe kopie historii'
        finally:
            if book is not None:
                app.EnableEvents = False
                book.Close(False)
            app.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('source', type=Path)
    parser.add_argument('--baseline', action='store_true')
    parser.add_argument('--trace', action='store_true')
    args = parser.parse_args()
    run(args.source.resolve(), args.baseline, args.trace)
