"""Buduje XLSM w lokalnym Excelu; przywraca poprzednie ustawienie AccessVBOM."""

from contextlib import contextmanager
from pathlib import Path
import argparse
import json
import os
import sys
import winreg

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tools.release_version import prepare_version
OUT = ROOT / 'dist' / 'Floteka'
NAVY = 0x4F3A17  # Excel używa BGR
TEAL = 0x807D14
WHITE = 0xFFFFFF
REGISTRY_LAST_ROW = 5005


@contextmanager
def vba_access():
    key_path = r'Software\Microsoft\Office\16.0\Excel\Security'
    key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ | winreg.KEY_SET_VALUE)
    previous = None
    try:
        try:
            previous = winreg.QueryValueEx(key, 'AccessVBOM')
        except FileNotFoundError:
            pass
        winreg.SetValueEx(key, 'AccessVBOM', 0, winreg.REG_DWORD, 1)
        yield
    finally:
        if previous is None:
            try:
                winreg.DeleteValue(key, 'AccessVBOM')
            except FileNotFoundError:
                pass
        else:
            winreg.SetValueEx(key, 'AccessVBOM', 0, previous[1], previous[0])
        winreg.CloseKey(key)


def button(sheet, label, action, left, top=None, width=150):
    if top is None:
        top = sheet.Range('A3').Top + (sheet.Rows(3).Height - 29) / 2
    shape = sheet.Shapes.AddShape(5, left, top, width, 29)
    shape.Fill.ForeColor.RGB = TEAL
    shape.Line.Visible = False
    shape.TextFrame.Characters().Text = label
    shape.TextFrame.Characters().Font.Color = WHITE
    shape.TextFrame.Characters().Font.Size = 10
    shape.TextFrame.HorizontalAlignment = -4108
    shape.TextFrame.VerticalAlignment = -4108
    shape.OnAction = action
    shape.Name = 'toolbar_' + action
    shape.Placement = 2
    return shape


def setup_sheet(sheet, title, subtitle, headers, widths):
    sheet.Cells.Font.Name = 'Calibri'
    sheet.Cells.Font.Size = 11
    sheet.Cells.Locked = True
    for index, width in enumerate(widths, 1):
        sheet.Columns(index).ColumnWidth = width
    count = len(headers)
    sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, count)).Merge()
    sheet.Cells(1, 1).Value = title
    sheet.Cells(1, 1).Interior.Color = NAVY
    sheet.Cells(1, 1).Font.Color = WHITE
    sheet.Cells(1, 1).Font.Size = 22
    sheet.Cells(1, 1).Font.Bold = True
    sheet.Rows(1).RowHeight = 42
    sheet.Range(sheet.Cells(2, 1), sheet.Cells(2, count)).Merge()
    sheet.Cells(2, 1).Value = subtitle
    sheet.Cells(2, 1).Font.Color = 0x806040
    sheet.Cells(2, 1).WrapText = True
    sheet.Rows(2).RowHeight = 36
    sheet.Rows(3).RowHeight = 44
    sheet.Rows(4).RowHeight = 8
    # Resize jest indeksowaną właściwością COM. Wywołanie jej jak funkcji
    # w późnym wiązaniu wybierało jedną komórkę (np. H5 zamiast A5:H5).
    head = sheet.Range(sheet.Cells(5, 1), sheet.Cells(5, count))
    head.Value = (tuple(headers),)
    head.Interior.Color = NAVY
    head.Font.Color = WHITE
    head.Font.Bold = True
    head.WrapText = True
    head.RowHeight = 38
    sheet.Activate()
    sheet.Application.ActiveWindow.DisplayGridlines = False
    sheet.Range('C6' if sheet.Name == 'DANE' else 'A6').Select()
    sheet.Application.ActiveWindow.FreezePanes = True
    sheet.Application.ActiveWindow.Zoom = 90
    sheet.PageSetup.Orientation = 2
    try:
        sheet.PageSetup.PaperSize = 9
    except Exception:
        # Bez zainstalowanej drukarki Excel odrzuca format papieru.
        # Widok i eksport pliku pozostają dostępne.
        pass
    sheet.PageSetup.Zoom = False
    sheet.PageSetup.FitToPagesWide = 1
    sheet.PageSetup.FitToPagesTall = False
    sheet.PageSetup.PrintTitleRows = '$5:$5'
    sheet.PageSetup.LeftMargin = 18
    sheet.PageSetup.RightMargin = 18


def setup_optional_fields(sheet):
    name = sheet.Name
    if name in ('DANE', 'ARCHIWUM'):
        sheet.Range('K5:L5').Value2 = (('Trasa', 'Numer zlecenia'),)
        sheet.Columns('K').ColumnWidth = 30
        sheet.Columns('L').ColumnWidth = 23
        sheet.Columns('K:L').NumberFormat = '@'
        sheet.Columns('K:L').Hidden = False
        sheet.Columns('I:J' if name == 'DANE' else 'J').Hidden = True
        head = sheet.Range('K5:L5')
        last = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1
        if last >= 6:
            details = sheet.Range(f'K6:L{last}')
            details.WrapText = True
            details.VerticalAlignment = -4108
            details.Interior.Color = 0xF7F6ED if name == 'DANE' else 0xF9F7F0
            if name == 'DANE':
                details.Locked = False
                sheet.PageSetup.PrintArea = f'$A$5:$L${last}'
    elif name == 'KIEROWCY':
        head = sheet.Range('E5')
        head.Value2 = 'Telefon'
        sheet.Columns('E').ColumnWidth = 23
        phone = sheet.Range(f'E6:E{REGISTRY_LAST_ROW}')
        phone.Locked = False
        phone.NumberFormat = '@'
    elif name == 'TIRY':
        if sheet.Range('A5').MergeCells:
            sheet.Range('A5').MergeArea.UnMerge()
        head = sheet.Range('A5:F5')
        head.Value2 = (('Rejestracja', 'Przebieg (km)', 'Data przeglądu', 'Data ubezpieczenia', 'Data serwisu', 'Uwagi'),)
        for index, width in enumerate((18, 20, 23, 20, 42), 2):
            sheet.Columns(index).ColumnWidth = width
        sheet.Range(f'B6:F{REGISTRY_LAST_ROW}').Locked = False
        sheet.Range(f'B6:B{REGISTRY_LAST_ROW}').NumberFormat = '0'
        sheet.Range(f'C6:E{REGISTRY_LAST_ROW}').NumberFormat = 'dd.mm.yyyy'
        sheet.Range(f'F6:F{REGISTRY_LAST_ROW}').NumberFormat = '@'
        sheet.Range(f'F6:F{REGISTRY_LAST_ROW}').WrapText = True
    else:
        return
    head.Interior.Color = NAVY
    head.Font.Color = WHITE
    head.Font.Bold = True
    head.WrapText = True
    head.RowHeight = 38


def build(output=None, tag=None):
    global OUT
    if output is not None:
        OUT = Path(output).resolve()
    OUT.mkdir(parents=True, exist_ok=True)
    release = prepare_version(ROOT, tag, require_commit=os.environ.get('GITHUB_ACTIONS') == 'true')
    with vba_access():
        app = win32com.client.DispatchEx('Excel.Application')
        book = None
        try:
            app.Visible = False
            app.DisplayAlerts = False
            app.EnableEvents = False
            book = app.Workbooks.Add()
            book.CustomDocumentProperties.Add('FlotekaVersion', False, 4, release.version)
            book.BuiltinDocumentProperties('Title').Value = f'Floteka {release.version}'
            while book.Worksheets.Count > 1:
                book.Worksheets(book.Worksheets.Count).Delete()
            book.Worksheets(1).Name = 'START'
            for name in ('DANE', 'ARCHIWUM', 'KIEROWCY', 'TIRY', 'PRZYCZEPY', '_STAN'):
                book.Worksheets.Add(pythoncom.Missing, book.Worksheets(book.Worksheets.Count)).Name = name
            start = book.Worksheets('START')
            setup_sheet(start, 'FLOTEKA', f'Wersja {release.version} | Kierowcy • TIR-y • przyczepy | cztery miejsca i pełna historia',
                        ['Szybki start', '', '', '', '', ''], [30, 26, 26, 26, 26, 26])
            start.Range('A5:F5').Merge()
            help_rows = [
                (7, '1  DODAJ TIRA I KIEROWCĘ', 'Kliknij „Dodaj kierowcę” i wpisz nazwę oraz tira. Powstaną cztery miejsca w DANE.'),
                (10, '2  UZUPEŁNIAJ DANE', 'W DANE wpisz przyczepę, uwagi i daty DD.MM.RRRR w C:F. Trasę oraz numer zlecenia możesz dodać w K:L.'),
                (13, '3  HISTORIA POWSTAJE OD RAZU', 'Przyczepa i obie daty oznaczają kompletny wpis. Jego kopia od razu trafia do ARCHIWUM.'),
                (16, '4  ZWALNIAJ MIEJSCA PRZYCISKIEM', 'Zaznacz kierowcę i kliknij „Zwolnij najstarsze” albo x przy wierszu. Historia zostaje.'),
                (19, '5  POPRAWKI ZOSTAJĄ W HISTORII', 'Poprawiony kompletny wpis tworzy nową wersję. Identyczne dane nie są dopisywane drugi raz.'),
                (22, '6  ZAPISUJ PLIK  Ctrl+S', 'Szkice są zachowywane w DANE po zapisaniu pliku. Niekompletne wpisy nie trafiają do archiwum.'),
                (25, 'EDYCJA KIEROWCY / TIRA', 'W KIEROWCY możesz podać telefon, a w TIRY dane pojazdu. Aktywny = NIE ukrywa kierowcę bez kasowania danych.'),
                (28, 'POWRÓT DO DAWNEGO WPISU', 'W ARCHIWUM zaznacz wpis i kliknij „Przywróć wpis”. Wybierz puste miejsce. Użyty będzie obecny tir.'),
                (31, 'MAKRA I BEZPIECZNY ZAPIS', 'Włącz makra tylko dla tej sprawdzonej kopii. Bez makr przyciski i historia nie działają.'),
                (34, 'DWA NIEZALEŻNE NARZĘDZIA', 'Skoroszyt i aplikacja SQLite nie synchronizują się. Z aplikacji można wyeksportować XLSX.'),
                (37, 'KOLORY TERMINÓW ROZŁADUNKU', 'Czerwony: data przed dzisiaj. Pomarańczowy: dziś lub jutro. Przycisk „Kolory terminów” włącza i wyłącza oznaczenia.'),
            ]
            for row, title, text in help_rows:
                start.Range(f'A{row}:F{row}').Merge()
                start.Cells(row, 1).Value = title
                start.Cells(row, 1).Font.Bold = True
                start.Cells(row, 1).Font.Color = TEAL
                start.Rows(row).RowHeight = 24
                start.Range(f'A{row+1}:F{row+1}').Merge()
                start.Cells(row + 1, 1).Value = text
                start.Cells(row + 1, 1).WrapText = True
                start.Rows(row + 1).RowHeight = 28
            button(start, '+ Dodaj kierowcę', 'DodajKierowce', 8, width=170)
            button(start, '+ TIR', 'DodajTir', 188, width=110)
            button(start, '+ Przyczepa', 'DodajPrzyczepe', 308, width=140)
            button(start, 'Kopia pliku', 'KopiaPliku', 458, width=140)
            button(start, 'Kolory terminów: WŁ.', 'PrzelaczKoloryTerminow', 608, width=200)
            data = book.Worksheets('DANE')
            headers = ['Kierowca', 'TIR — rejestracja', 'Przyczepa — rejestracja', 'Informacje dodatkowe', 'Data załadunku', 'Data rozładunku', 'Status', 'Wyczyść']
            setup_sheet(data, 'DANE BIEŻĄCE', 'Edytuj C:F oraz opcjonalnie K:L. Każdy kierowca ma 4 miejsca. Zaznacz wiersz kierowcy przed zwolnieniem najstarszego.',
                        headers, [28, 20, 24, 45, 19, 19, 15, 11])
            data.Columns('I:J').Hidden = True
            data.Columns('C:D').NumberFormat = '@'
            data.Columns('E:F').NumberFormat = 'dd.mm.yyyy'
            button(data, '+ Kierowca', 'DodajKierowce', 8, width=130)
            button(data, 'Zwolnij najstarsze', 'ZwolnijNajstarsze', 148, width=180)
            button(data, 'Zapisz wszystko', 'ZapiszWszystko', 338, width=155)
            button(data, '+ Przyczepa', 'DodajPrzyczepe', 503, width=140)
            button(data, 'Kolory terminów: WŁ.', 'PrzelaczKoloryTerminow', 653, width=200)
            archive = book.Worksheets('ARCHIWUM')
            setup_sheet(archive, 'PEŁNA HISTORIA', 'Pełne wpisy, również nadal widoczne. Poprawki tworzą kolejne kopie. Historia nie jest usuwana.',
                        headers[:6] + ['Zapisano (lokalnie)', 'ID kierowcy', 'ID historii'], [28, 20, 24, 45, 19, 19, 25, 14, 14])
            archive.Columns('A:D').NumberFormat = '@'
            archive.Columns('J:J').NumberFormat = '@'
            archive.Columns('J:J').Hidden = True
            button(archive, 'Przywróć wpis', 'PrzywrocWpis', 8, width=170)
            reg = book.Worksheets('KIEROWCY')
            setup_sheet(reg, 'REJESTR KIEROWCÓW', 'Dodaj kierowcę przyciskiem lub wpisz nazwę i tira w nowym wierszu. Wyłączenie z listy: Aktywny = NIE.',
                        ['ID', 'Kierowca', 'TIR', 'Aktywny'], [10, 35, 25, 18])
            reg.Range('B6:D5005').Locked = False
            reg.Range('B6:D5005').NumberFormat = '@'
            button(reg, '+ Dodaj kierowcę', 'DodajKierowce', 8, width=180)
            for name, title, action in [('TIRY', 'REJESTR TIRÓW', 'DodajTir'), ('PRZYCZEPY', 'REJESTR PRZYCZEP', 'DodajPrzyczepe')]:
                sheet = book.Worksheets(name)
                setup_sheet(sheet, title, 'Wpisuj kolejne numery poniżej. Nowe przyczepy z DANE są dopisywane automatycznie.',
                            ['Rejestracja', ''], [28, 80])
                sheet.Range('A5:B5').Merge()
                sheet.Range('A6:A5005').Locked = False
                sheet.Range('A6:A5005').NumberFormat = '@'
                button(sheet, '+ Dodaj pojazd', action, 8, width=180)
            for sheet in (data, archive, reg, book.Worksheets('TIRY')):
                setup_optional_fields(sheet)
            for sheet in book.Worksheets:
                sheet.PageSetup.LeftFooter = f'Floteka {release.version}'
            book.Names.Add(Name='ListaTirow', RefersTo="='TIRY'!$A$6:$A$6")
            book.Names.Add(Name='ListaPrzyczep', RefersTo="='PRZYCZEPY'!$A$6:$A$6")
            reg.Range('C6:C5005').Validation.Add(3, 2, 1, '=ListaTirow')
            reg.Range('C6:C5005').Validation.ShowError = False
            reg.Range('D6:D5005').Validation.Add(3, 1, 1, 'TAK,NIE')
            state = book.Worksheets('_STAN')
            state.Range('A1:B2').Value = (('Następny ID', 0), ('Kolejność miejsc', 0))
            state.Visible = 2  # xlSheetVeryHidden
            # Jawna kolejność, niezależna od zachowania opcjonalnego argumentu Add.
            order = ['START', 'DANE', 'ARCHIWUM', 'KIEROWCY', 'TIRY', 'PRZYCZEPY', '_STAN']
            for name in reversed(order):
                if book.Worksheets(1).Name != name:
                    book.Worksheets(name).Move(book.Worksheets(1))
            module = book.VBProject.VBComponents.Add(1)
            module.Name = 'Flota'
            module.CodeModule.AddFromString((ROOT / 'excel' / 'Flota.bas').read_text(encoding='utf-8'))
            book.VBProject.VBComponents(book.CodeName).CodeModule.AddFromString((ROOT / 'excel' / 'ThisWorkbook.cls').read_text(encoding='utf-8'))
            app.Run(f"'{book.Name}'!Initialize")
            start.Activate()
            start.Range('A1').Select()
            app.ActiveWindow.FreezePanes = False
            clean = OUT / 'Floteka.xlsm'
            book.SaveAs(str(clean), FileFormat=52)
            print('XLSM clean:', clean.name, flush=True)
            automation = book.VBProject.VBComponents.Add(1)
            automation.Name = 'BuildAutomation'
            automation.CodeModule.AddFromString('''Public Function SafeAdd(ByVal person As String, ByVal truck As String) As String
On Error GoTo Failed
SafeAdd = CStr(Flota.AddDriverRecord(person, truck))
Exit Function
Failed:
SafeAdd = "ERROR: " & Err.Description
End Function''')
            app.EnableEvents = False
            print('ProtectionMode:', data.ProtectionMode, flush=True)
            result = app.Run(f"'{book.Name}'!SafeAdd", 'Kierowca pokazowy A', 'WX10001')
            print('Add driver:', result, flush=True)
            driver = int(result)
            result = app.Run(f"'{book.Name}'!SafeAdd", 'Kierowca pokazowy B', 'PO20002')
            if str(result).startswith('ERROR'):
                raise RuntimeError(result)
            app.EnableEvents = True
            data.Range('C6:F6').Value = (('WX30003', 'PRZYKŁAD — firma A', '17.11.2026', '19.11.2026'),)
            data.Range('C7:F7').Value = (('WX40004', 'PRZYKŁAD — firma B', '20.11.2026', '23.11.2026'),)
            data.Range('C8:D8').Value = (('WX50005', 'PRZYKŁAD — szkic, uzupełnij daty'),)
            start.Activate()
            book.VBProject.VBComponents.Remove(automation)
            # Usunięcie modułu resetuje projekt VBA i może przywrócić alerty Excela.
            app.DisplayAlerts = False
            book.SaveAs(str(OUT / 'Floteka-DEMO.xlsm'), FileFormat=52)
            print('XLSM demo: OK', flush=True)
            (ROOT / 'docs' / 'excel-build.json').write_text(json.dumps({'ok': True, 'excel': app.Version,
                'version': release.version, 'tag': release.tag,
                'commit': release.commit,
                'macro_lines': module.CodeModule.CountOfLines, 'files': [clean.name, 'Floteka-DEMO.xlsm'],
                'AccessVBOM': 'przywracany w finally'}, ensure_ascii=False, indent=2), encoding='utf-8')
        finally:
            if book is not None:
                app.EnableEvents = False
                book.Close(False)
            app.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path)
    parser.add_argument('--tag', default=os.environ.get('FLOTEKA_BUILD_TAG') or None)
    args = parser.parse_args()
    build(args.output, args.tag)
