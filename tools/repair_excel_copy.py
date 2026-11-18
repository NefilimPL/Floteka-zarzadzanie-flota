"""Poprawia wyłącznie nową kopię skoroszytu; sprawdza zachowanie jego danych."""
import argparse
import hashlib
import json
import shutil
from pathlib import Path

import win32com.client
from openpyxl import load_workbook

from build_excel import ROOT, NAVY, WHITE, vba_access, button, setup_optional_fields

HEADERS = {
    'DANE': ['Kierowca', 'TIR — rejestracja', 'Przyczepa — rejestracja', 'Informacje dodatkowe', 'Data załadunku', 'Data rozładunku', 'Status', 'Wyczyść'],
    'ARCHIWUM': ['Kierowca', 'TIR — rejestracja', 'Przyczepa — rejestracja', 'Informacje dodatkowe', 'Data załadunku', 'Data rozładunku', 'Zapisano (lokalnie)', 'ID kierowcy', 'ID historii'],
    'KIEROWCY': ['ID', 'Kierowca', 'TIR', 'Aktywny'],
}


def content(path):
    wb = load_workbook(path, keep_vba=True)
    try:
        result = {}
        for name, columns in (('DANE', 12), ('ARCHIWUM', 12), ('KIEROWCY', 5),
                              ('TIRY', 6), ('PRZYCZEPY', 1), ('_STAN', 17)):
            ws = wb[name]
            first = 1 if name == '_STAN' else 6
            # Porównujemy dane z adresami. Nowe puste, sformatowane komórki
            # nie wpływają na wynik, a ukryte lub częściowe rekordy pozostają.
            rows = ws.iter_rows(min_row=first, max_col=max(columns, ws.max_column), values_only=True)
            result[name] = tuple(
                (row, col, value)
                for row, values in enumerate(rows, first)
                for col, value in enumerate(values, 1)
                if value is not None and not (name == '_STAN' and row == 3 and col <= 3)
            )
        return result
    finally:
        wb.close()


def repair(source, target):
    source, target = Path(source).resolve(), Path(target).resolve()
    if source == target:
        raise ValueError('Wymagana jest osobna kopia docelowa.')
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    before = content(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    with vba_access():
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        book = None
        try:
            book = app.Workbooks.Open(str(target))
            for sheet in book.Worksheets:
                if sheet.Name == '_STAN':
                    continue
                sheet.Unprotect('flota')
                sheet.Cells(2, 1).WrapText = True
                sheet.Rows(2).RowHeight = 36
                sheet.Rows(3).RowHeight = 44
                if sheet.Name in HEADERS:
                    headers = HEADERS[sheet.Name]
                    head = sheet.Range(sheet.Cells(5, 1), sheet.Cells(5, len(headers)))
                    head.Value = (tuple(headers),)
                    head.Interior.Color = NAVY
                    head.Font.Color = WHITE
                    head.Font.Bold = True
                    head.WrapText = True
                    head.RowHeight = 38
                elif sheet.Name in ('TIRY', 'PRZYCZEPY'):
                    sheet.Range('A5').Value = 'Rejestracja'
                elif sheet.Name == 'START':
                    sheet.Range('A1').Value = 'FLOTEKA'
                    sheet.Range('A5').Value = 'Szybki start'
                setup_optional_fields(sheet)
                if sheet.Name in ('START', 'DANE'):
                    existing = [shape.Name for shape in sheet.Shapes]
                    if 'toolbar_PrzelaczKoloryTerminow' not in existing:
                        button(sheet, 'Kolory terminów: WŁ.', 'PrzelaczKoloryTerminow',
                               608 if sheet.Name == 'START' else 653, width=200)
                for shape in sheet.Shapes:
                    # Excel dodaje do Shapes także własne kontrolki filtrów.
                    # Nie mają one zwykłej właściwości Placement.
                    toolbar_actions = ('DodajKierowce', 'DodajTir', 'DodajPrzyczepe',
                                       'ZwolnijNajstarsze', 'ZapiszWszystko', 'KopiaPliku', 'PrzywrocWpis', 'PrzelaczKoloryTerminow')
                    if shape.Type == 1 and shape.OnAction in toolbar_actions:
                        shape.Top = sheet.Range('A3').Top + (sheet.Rows(3).Height - 29) / 2
                        shape.Height = 29
                        if shape.Placement != 2:
                            shape.Placement = 2
            module = book.VBProject.VBComponents('Flota').CodeModule
            module.DeleteLines(1, module.CountOfLines)
            module.AddFromString((ROOT / 'excel' / 'Flota.bas').read_text(encoding='utf-8'))
            # Zachowaj również techniczne wartości kopii. Zwykłe otwarcie
            # odświeży punkty przywracania rejestru z komórek użytkownika.
            app.Run(f"'{book.Name}'!Initialize", False)
            app.EnableEvents = False  # Zachowanie każdego rekordu bez zapisu nowych wersji.
            book.Worksheets('DANE').Activate()
            book.Worksheets('DANE').Range('C6').Select()
            app.DisplayAlerts = False
            book.Save()
        finally:
            if book is not None:
                app.EnableEvents = False
                book.Close(False)
            app.Quit()
    after = content(target)
    assert before == after, 'Dane kopii różnią się od oryginału — nie przekazuj tej kopii.'
    assert hashlib.sha256(source.read_bytes()).hexdigest() == original_hash, 'Oryginał został zmieniony.'
    report = {'ok': True, 'original_untouched': True, 'all_data_and_history_preserved': True,
              'source_sha256': original_hash, 'target': target.name}
    (ROOT / 'docs' / 'excel-migration.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(json.dumps(report), flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('source', type=Path)
    parser.add_argument('target', type=Path)
    args = parser.parse_args()
    repair(args.source, args.target)
