"""Kolory wyświetlane przez rzeczywisty Excel, wyłącznie na kopii testowej."""

import json
import shutil
from datetime import date, timedelta

import win32com.client

from build_excel import ROOT
from excel_paths import default_excel_file


def run():
    target = ROOT / 'build' / 'excel-deadlines-test.xlsm'
    target.parent.mkdir(exist_ok=True)
    shutil.copy2(default_excel_file(), target)
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 1
    book = None
    checks = []
    red = 229 + 57 * 256 + 53 * 65536
    orange = 255 + 167 * 256 + 38 * 65536
    try:
        book = app.Workbooks.Open(str(target))
        macro = lambda name, *args: app.Run(f"'{book.Name}'!{name}", *args)
        driver = int(macro('AddDriverRecord', 'Test terminów', 'TEST100'))
        first = int(macro('DriverFirstRow', driver))
        sheet = book.Worksheets('DANE')
        for index, days in enumerate((-1, 0, 1, 2)):
            load = (date.today() - timedelta(days=10)).strftime('%d.%m.%Y')
            unload = (date.today() + timedelta(days=days)).strftime('%d.%m.%Y')
            sheet.Range(sheet.Cells(first + index, 3), sheet.Cells(first + index, 6)).Value = (
                (f'TEST{index}', 'Test kolorów', load, unload),)
        app.CalculateFull()
        for index, expected in enumerate((red, orange, orange, None)):
            for col in range(1, 13):
                color = int(sheet.Cells(first + index, col).DisplayFormat.Interior.Color)
                assert color == expected if expected else color not in (red, orange), (index, col, color)
        archive = book.Worksheets('ARCHIWUM')
        assert int(archive.Cells(6, 1).DisplayFormat.Interior.Color) == red
        checks.append('całe wiersze: wczoraj czerwone, dziś/jutro pomarańczowe, pojutrze zwykłe; także historia')
        macro('PrzelaczKoloryTerminow')
        assert int(sheet.Cells(first, 1).DisplayFormat.Interior.Color) not in (red, orange)
        book.Save()
        book.Close(False)
        book = app.Workbooks.Open(str(target))
        assert not bool(book.Worksheets('_STAN').Range('B3').Value)
        sheet = book.Worksheets('DANE')
        assert int(sheet.Cells(first, 1).DisplayFormat.Interior.Color) not in (red, orange)
        checks.append('wyłączenie kolorów zachowane po ponownym otwarciu')
        macro('PrzelaczKoloryTerminow')
        assert int(sheet.Cells(first, 1).DisplayFormat.Interior.Color) == red
        second = int(macro('AddDriverRecord', 'Nowy blok', 'TEST200'))
        new_row = int(macro('DriverFirstRow', second))
        assert int(sheet.Cells(new_row, 1).DisplayFormat.Interior.Color) not in (red, orange)
        sheet.Cells(new_row, 6).Value = (date.today() - timedelta(days=1)).strftime('%d.%m.%Y')
        app.CalculateFull()
        assert int(sheet.Cells(new_row, 1).DisplayFormat.Interior.Color) == red
        assert sheet.Cells(first, 1).FormatConditions.Count == 2
        checks.append('puste daty bez koloru; nowe bloki i szkice z datą objęte dwiema regułami')
        report = {'ok': True, 'checks': checks, 'excel': app.Version}
        (ROOT / 'docs' / 'excel-deadlines.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(report, ensure_ascii=False))
    finally:
        if book is not None:
            app.EnableEvents = False
            book.Close(False)
        app.Quit()


if __name__ == '__main__':
    run()
