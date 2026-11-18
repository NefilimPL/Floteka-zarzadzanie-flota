"""Otwiera eksporty źródeł i EXE w Excelu bez użycia trybu naprawy."""

import argparse
import json
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import win32com.client
from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel
from datetime import date, datetime

from fleet.export import export_xlsx
from fleet.store import Store


def snapshot_values(path):
    workbook = load_workbook(path)
    try:
        return {sheet.title: [[to_excel(cell.value, workbook.epoch) if isinstance(cell.value, (date, datetime)) else cell.value
                               for cell in row] for row in sheet] for sheet in workbook}
    finally:
        workbook.close()


def check_open(app, path, *, populated, expected_names=None, expected_values=None):
    """CorruptLoad=0: błąd tabel ma przerwać test, bez odzyskiwania danych."""
    book = app.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        if expected_names is not None:
            assert [sheet.Name for sheet in book.Worksheets] == expected_names
        tables = 0
        for sheet in book.Worksheets:
            if expected_values is not None:
                rows = expected_values[sheet.Name]
                cells = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(rows), len(rows[0])))
                assert [list(row) for row in cells.Value2] == rows, sheet.Name
                assert cells.HasFormula is False, sheet.Name
            expected_tables = int(populated and sheet.Name in ('DANE', 'ARCHIWUM'))
            assert sheet.ListObjects.Count == expected_tables, sheet.Name
            tables += sheet.ListObjects.Count
            if sheet.Name in ('KIEROWCY', 'TIRY'):
                assert sheet.AutoFilterMode, sheet.Name
            if expected_tables:
                table = sheet.ListObjects(1)
                assert table.ShowAutoFilter
                assert table.AutoFilter.Filters.Count == (10 if sheet.Name == 'DANE' else 11)
                assert table.DataBodyRange.Rows.Count >= 1
        return tables
    finally:
        book.Close(False)


def run(exe=None):
    work = ROOT / 'build'
    work.mkdir(exist_ok=True)
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.AutomationSecurity = 3
    checks = []
    try:
        with TemporaryDirectory(prefix='xlsx-test-', dir=work) as directory:
            folder = Path(directory)
            store = Store(folder / 'source.sqlite3')
            try:
                truck = store.add_vehicle('truck', 'TEST1')
                driver = store.add_driver('Test Żółć', truck, phone='+48 000 000 000')
                store.save_truck_details(truck, odometer_km='123456')
                store.save_slot(driver, 1, 'TEST2', '=tekst użytkownika', '2026-11-17', '2026-11-18',
                                route='Łódź → Berlin', order_number='ZL-1')
                for scope, names in (
                    ('both', ['DANE', 'ARCHIWUM', 'KIEROWCY', 'TIRY']),
                    ('current', ['DANE', 'KIEROWCY', 'TIRY']),
                    ('history', ['ARCHIWUM']),
                ):
                    path = folder / (scope + '.xlsx')
                    export_xlsx(store, path, scope=scope)
                    values = snapshot_values(path)
                    count = check_open(app, path, populated=True, expected_names=names, expected_values=values)
                    book = app.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
                    try:
                        sheet = book.Worksheets(names[0])
                        assert sheet.Cells(2, 1).Value2 == 'Test Żółć'
                        assert sheet.Cells(2, 4).Value2 == '=tekst użytkownika'
                        assert not sheet.Cells(2, 4).HasFormula
                        route_col = 10 if scope == 'history' else 9
                        assert sheet.Cells(2, route_col).Value2 == 'Łódź → Berlin'
                        assert sheet.Cells(2, route_col + 1).Value2 == 'ZL-1'
                        assert sheet.Cells(2, 5).Value2 == 46343  # 17.11.2026, system dat 1900.
                        roundtrip = folder / (scope + '-saved.xlsx')
                        book.SaveCopyAs(str(roundtrip))
                    finally:
                        book.Close(False)
                    assert check_open(app, roundtrip, populated=True, expected_names=names, expected_values=values) == count
                    checks.append(f'{scope}: normalne otwarcie, tabele/filtry i dane zachowane po zapisie Excela')
            finally:
                store.close()
            empty = Store(folder / 'empty.sqlite3')
            try:
                path = folder / 'empty.xlsx'
                export_xlsx(empty, path)
                assert check_open(app, path, populated=False, expected_names=['DANE', 'ARCHIWUM']) == 0
                driver = empty.add_driver('Szkic', empty.add_vehicle('truck', 'TEST3'))
                empty.save_slot(driver, 1, info='Tylko szkic')
                path = folder / 'draft.xlsx'
                export_xlsx(empty, path)
                book = app.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
                try:
                    assert book.Worksheets('DANE').ListObjects.Count == 1
                    assert book.Worksheets('ARCHIWUM').ListObjects.Count == 0
                finally:
                    book.Close(False)
                checks.append('pusty eksport i sam szkic: normalne otwarcie bez naprawy')
            finally:
                empty.close()
            if exe is not None:
                report = folder / 'frozen.json'
                exported = folder / 'frozen-export.xlsx'
                subprocess.run([str(Path(exe).resolve()), '--self-test', str(report), '--self-test-export', str(exported)], check=True, timeout=60,
                               creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                result = json.loads(report.read_text(encoding='utf-8'))
                assert result['ok'] and result['frozen']
                assert check_open(app, exported, populated=True, expected_values=snapshot_values(exported),
                                  expected_names=['DANE', 'ARCHIWUM', 'TIRY']) == 2
                checks.append('XLSX utworzony przez gotowy EXE: normalne otwarcie, dwie tabele i filtry zachowane')
        result = {'ok': True, 'excel': app.Version, 'frozen_export_checked': exe is not None, 'checks': checks}
        (ROOT / 'docs' / 'excel-export.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(result, ensure_ascii=False), flush=True)
    finally:
        app.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--exe', type=Path)
    run(parser.parse_args().exe)
