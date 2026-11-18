"""Sprawdza pola opcjonalne i historię w Excelu na tymczasowej kopii XLSM."""

from contextlib import contextmanager
from datetime import date, timedelta
import json
from pathlib import Path
import shutil
from tempfile import TemporaryDirectory
import threading

import win32com.client
import win32con
import win32gui
import win32process

from build_excel import ROOT
from excel_paths import default_excel_file
from test_excel_regressions import acknowledge_test_dialog


@contextmanager
def decline_draft_clear(app):
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    stop = threading.Event()
    seen = []

    def visit(hwnd, _):
        if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
            return
        if win32gui.GetClassName(hwnd) == '#32770' and win32gui.GetWindowText(hwnd) == 'Wyczyścić szkic?':
            button = win32gui.GetDlgItem(hwnd, win32con.IDNO)
            if button:
                seen.append(hwnd)
                win32gui.PostMessage(button, win32con.BM_CLICK, 0, 0)

    def worker():
        while not stop.wait(0.05):
            win32gui.EnumWindows(visit, None)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    try:
        yield seen
    finally:
        stop.set()
        thread.join(2)


def row(sheet, address, *values):
    sheet.Range(address).Value2 = (values,)


def serial(day):
    return (day - date(1899, 12, 30)).days


def rgb(red, green, blue):
    return red + green * 256 + blue * 65536


def assert_colors(sheet, line, fill, font):
    for col in range(1, 13):
        display = sheet.Cells(line, col).DisplayFormat
        assert int(display.Interior.Color) == fill, (sheet.Name, line, col, 'wypełnienie')
        assert int(display.Font.Color) == font, (sheet.Name, line, col, 'tekst')


def check_optional_data(app, path):
    book = None
    checks = []
    try:
        book = app.Workbooks.Open(str(path))

        def macro(name, *args):
            return app.Run(f"'{book.Name}'!{name}", *args)

        driver = int(macro('AddDriverRecord', 'Kierowca testowy', 'TEST100'))
        assert macro('DriverFirstRow', driver) == 6
        data, archive = book.Worksheets('DANE'), book.Worksheets('ARCHIWUM')
        state = book.Worksheets('_STAN')
        assert data.Range('I:J').EntireColumn.Hidden and archive.Range('J:J').EntireColumn.Hidden
        assert not data.Range('K6:L9').Locked and archive.Range('K6:L9').Locked
        row(data, 'K6:L6', 'Gdańsk – Łódź', '000123')
        assert data.Range('G6').Value2 == 'Szkic' and macro('HistoryCount') == 0
        assert state.Range('J6:K6').Value2 == (('Gdańsk – Łódź', '000123'),)
        with decline_draft_clear(app) as declined:
            data.Range('K6:L6').ClearContents()
        assert declined and data.Range('K6:L6').Value2 == (('Gdańsk – Łódź', '000123'),)
        with decline_draft_clear(app) as declined:
            macro('ClearRow', 6)
        assert declined and data.Range('K6:L6').Value2 == (('Gdańsk – Łódź', '000123'),)
        checks.append('szkic z samą trasą/zleceniem i odmowa czyszczenia przez Delete lub przycisk')

        today = date.today()
        row(data, 'C6:F6', 'PO123', 'Ładunek', serial(today - timedelta(days=5)), serial(today - timedelta(days=1)))
        assert data.Range('G6').Value2 == 'Kompletny' and macro('HistoryCount') == 1
        assert archive.Range('K6:L6').Value2 == (('Gdańsk – Łódź', '000123'),)
        archive.Range('A5:L6').AutoFilter(11, 'Gdańsk – Łódź')
        data.Range('L6').Value2 = '000124'
        assert macro('HistoryCount') == 2
        assert archive.Range('K6:L7').Value2 == (('Gdańsk – Łódź', '000123'), ('Gdańsk – Łódź', '000124'))
        assert archive.AutoFilter.Filters(11).On
        assert str(archive.AutoFilter.Filters(11).Criteria1).lstrip('=') == 'Gdańsk – Łódź'
        macro('ClearRow', 6)
        assert data.Range('K6:L6').Value2 == ((None, None),) and data.Range('G6').Value2 == 'Puste'
        macro('RestoreRow', 6, 6)
        assert data.Range('K6:L6').Value2 == (('Gdańsk – Łódź', '000123'),) and macro('HistoryCount') == 2
        checks.append('historia wersji opcjonalnych, filtr trasy, zwolnienie i przywrócenie bez duplikatu')

        row(data, 'K7:L7', 'Późniejszy szkic', '000125')
        data.Unprotect('flota')
        assert not data.ProtectContents
        data.Range('K7').NumberFormat = '0'
        macro('Initialize')
        with acknowledge_test_dialog(app) as dialogs:
            data.Range('K7').Formula = '=1+1'
        assert dialogs and data.Range('K7:L7').Value2 == (('Późniejszy szkic', '000125'),)
        row(data, 'C8:F8', 'PO124', '', serial(today - timedelta(days=5)), serial(today + timedelta(days=1)))
        assert macro('HistoryCount') == 3 and archive.Range('K8:L8').Value2 == ((None, None),)
        red, orange = rgb(229, 57, 53), rgb(255, 167, 38)
        white, dark = rgb(255, 255, 255), rgb(33, 19, 0)
        app.CalculateFull()
        assert_colors(data, 6, red, white)
        assert_colors(data, 8, orange, dark)
        assert_colors(archive, 6, red, white)
        archive.ShowAllData()
        assert_colors(archive, 8, orange, dark)
        macro('PrzelaczKoloryTerminow')
        assert int(data.Range('L6').DisplayFormat.Interior.Color) != red
        checks.append('odrzucenie formuły w szkicu oraz kolory i kontrast na całej szerokości obu arkuszy')

        reg, trucks = book.Worksheets('KIEROWCY'), book.Worksheets('TIRY')
        assert not reg.Range('E6').Locked and not trucks.Range('B6:F6').Locked
        reg.Range('E6').Value2 = '+48 123 456 789'
        row(trucks, 'B6:F6', 0, '2026-11-16', '17.11.2026', None, 'Serwis planowany')
        expected_truck = ((0, serial(date(2026, 11, 16)), serial(date(2026, 11, 17)), None, 'Serwis planowany'),)
        assert trucks.Range('B6:F6').Value2 == expected_truck
        for cell, invalid in [('B6', -1), ('B6', 1.5), ('D6', '31.02.2026'), ('E6', '1900-02-29')]:
            with acknowledge_test_dialog(app) as dialogs:
                trucks.Range(cell).Value2 = invalid
            assert dialogs and trucks.Range('B6:F6').Value2 == expected_truck, (cell, invalid)
        with acknowledge_test_dialog(app) as dialogs:
            trucks.Range('B6').Formula = '=1+1'
        assert dialogs and trucks.Range('B6:F6').Value2 == expected_truck
        assert reg.Range('E6').Value2 == '+48 123 456 789' and macro('HistoryCount') == 3
        checks.append('telefon, nieujemny całkowity przebieg, daty i uwagi jako bieżący rejestr')

        book.Save()
        book.Close(False)
        book = app.Workbooks.Open(str(path))
        data, archive = book.Worksheets('DANE'), book.Worksheets('ARCHIWUM')
        assert macro('FlushAll') and macro('HistoryCount') == 3
        assert data.Range('K6:L7').Value2 == (('Gdańsk – Łódź', '000123'), ('Późniejszy szkic', '000125'))
        assert archive.Range('K6:L8').Value2 == (('Gdańsk – Łódź', '000123'), ('Gdańsk – Łódź', '000124'), (None, None))
        assert book.Worksheets('KIEROWCY').Range('E6').Value2 == '+48 123 456 789'
        assert book.Worksheets('TIRY').Range('B6:F6').Value2 == expected_truck
        assert int(data.Range('L6').DisplayFormat.Interior.Color) != red
        macro('PrzelaczKoloryTerminow')
        assert_colors(data, 6, red, white)
        assert data.UsedRange.Rows.Count < 100 and archive.UsedRange.Rows.Count < 100
        checks.append('ponowne otwarcie zachowuje szkic, historię bez duplikatów, rejestry i wyłączone kolory')
        return checks
    finally:
        if book is not None:
            app.EnableEvents = False
            book.Close(False)


def check_registry_after_disabled_macros(app, path):
    book = None
    try:
        app.EnableEvents = True
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        book = app.Workbooks.Open(str(path))
        reg, trucks = book.Worksheets('KIEROWCY'), book.Worksheets('TIRY')
        phone = '+48 600 123 456'
        saved = (135, serial(date(2026, 11, 16)), 'Dawna nieprawidłowa data', None, 'Zapisane bez makr')
        reg.Range('E6').Value2 = phone
        row(trucks, 'B6:F6', *saved)
        trucks.Range('B7').Formula = '=333+1'
        book.Save()
        book.Close(False)
        book = None

        app.AutomationSecurity = 1
        with acknowledge_test_dialog(app) as dialogs:
            book = app.Workbooks.Open(str(path))
        assert not dialogs, 'Błąd podczas otwierania pliku zapisanego bez makr'
        reg, trucks = book.Worksheets('KIEROWCY'), book.Worksheets('TIRY')
        assert trucks.Range('B6:F6').Value2 == (saved,)
        with acknowledge_test_dialog(app) as dialogs:
            trucks.Range('B6').Value2 = -1
        assert dialogs and trucks.Range('B6:F6').Value2 == (saved,), 'Utrata danych zapisanych bez makr'
        with acknowledge_test_dialog(app) as dialogs:
            reg.Range('E6').Value2 = 'x' * 101
        assert dialogs and reg.Range('E6').Value2 == phone
        with acknowledge_test_dialog(app) as dialogs:
            trucks.Range('B7').Value2 = -1
        assert dialogs and trucks.Range('B7').HasFormula and trucks.Range('B7').Formula == '=333+1'
        service = serial(date(2026, 11, 18))
        trucks.Range('E6').Value2 = service
        assert trucks.Range('B6:F6').Value2 == ((saved[0], saved[1], saved[2], service, saved[4]),)
        return 'rejestry zachowują dane zapisane bez makr, również wcześniejszą błędną wartość w innym polu'
    finally:
        if book is not None:
            app.EnableEvents = False
            book.Close(False)


def run():
    work = ROOT / 'build'
    work.mkdir(exist_ok=True)
    with TemporaryDirectory(prefix='excel-optional-', dir=work) as directory:
        path = Path(directory) / 'Floteka-test.xlsm'
        shutil.copy2(default_excel_file(), path)
        app = win32com.client.DispatchEx('Excel.Application')
        try:
            app.Visible = False
            app.DisplayAlerts = False
            app.AutomationSecurity = 1
            checks = check_optional_data(app, path)
            checks.append(check_registry_after_disabled_macros(app, path))
            report = {'ok': True, 'checks': checks, 'excel': app.Version}
            (ROOT / 'docs' / 'excel-optional.json').write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8',
            )
            print(json.dumps(report, ensure_ascii=False), flush=True)
        finally:
            app.Quit()


if __name__ == '__main__':
    run()
