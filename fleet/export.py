"""Eksport bez formuł użytkownika; docelowy plik zastępowany dopiero po sukcesie."""

import os
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from .deadlines import OVERDUE_FILL, SOON_FILL


HEADERS = ['Kierowca', 'TIR — rejestracja', 'Przyczepa — rejestracja', 'Informacje dodatkowe', 'Data załadunku', 'Data rozładunku']
KEYS = ['driver_name', 'truck', 'trailer', 'info', 'load_date', 'unload_date']


def write_sheet(ws, records, archive=False):
    headers = HEADERS + (['Zapisano (UTC)', 'ID kierowcy', 'ID historii'] if archive else ['Miejsce', 'Status']) + ['Trasa (opcjonalnie)', 'Numer zlecenia (opcjonalnie)']
    ws.append(headers)
    if len(records) > 1_048_575:
        raise ValueError('Za dużo wierszy dla jednego arkusza Excela. Wybierz jednego kierowcę.')
    for record in records:
        values = [record[k] for k in KEYS]
        for value in values:
            if isinstance(value, str) and len(value) > 32767:
                raise ValueError('Excel mieści najwyżej 32767 znaków w komórce. Dane w bazie pozostają pełne; skróć pole lub wykonaj kopię bazy.')
        for index in (4, 5):
            values[index] = date.fromisoformat(values[index]) if values[index] else None
        values += ([record['recorded_at'], record['driver_id'], record['id']] if archive else [record['position'], record['status']])
        values += [record.get('route', ''), record.get('order_number', '')]
        ws.append(values)
        for cell, value in zip(ws[ws.max_row], values):
            if isinstance(value, str):
                cell.data_type = 's'
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        for col in ('E', 'F'):
            ws[f'{col}{ws.max_row}'].number_format = 'dd.mm.yyyy'
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='173A4F')
        cell.font = Font(color='FFFFFF', bold=True)
    for col, width in zip('ABCDEFGHI', [29, 21, 24, 46, 20, 20, 28 if archive else 12, 18, 18]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'C2'
    # Tabela ma własny AutoFilter. Drugi filtr arkusza na tym samym zakresie
    # powoduje odrzucenie tabel przez Excel (openpyxl tego nie wykrywa).
    ws.auto_filter.ref = None if records else ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.column_dimensions[ws.cell(1, len(headers) - 1).column_letter].width = 36
    ws.column_dimensions[ws.cell(1, len(headers)).column_letter].width = 30
    if records:
        for formula, color in (
            ('AND(ISNUMBER($F2),$F2<TODAY())', OVERDUE_FILL),
            ('AND(ISNUMBER($F2),$F2>=TODAY(),$F2<=TODAY()+1)', SOON_FILL),
        ):
            ws.conditional_formatting.add(f'A2:{ws.cell(ws.max_row, len(headers)).coordinate}',
                FormulaRule(formula=[formula], fill=PatternFill('solid', fgColor=color),
                            font=Font(color='FFFFFF' if color == OVERDUE_FILL else '211300'), stopIfTrue=True))
        table = Table(displayName='Tabela' + ws.title, ref=ws.dimensions)
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(table)
    ws.print_title_rows = '1:1'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def write_optional_register(ws, headers, rows):
    ws.append(headers)
    for row in rows:
        ws.append(row)
        for cell, value in zip(ws[ws.max_row], row):
            if isinstance(value, str):
                cell.data_type = 's'
            if isinstance(value, date):
                cell.number_format = 'dd.mm.yyyy'
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='173A4F')
        cell.font = Font(color='FFFFFF', bold=True)
        ws.column_dimensions[cell.column_letter].width = 26
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False


def export_xlsx(store, path, scope='both', driver_id=None):
    if scope not in ('both', 'current', 'history'):
        raise ValueError('Nieprawidłowy zakres eksportu.')
    path = Path(path).resolve()
    if path == store.path or path.suffix.lower() != '.xlsx':
        raise ValueError('Wybierz plik z rozszerzeniem .xlsx, inny niż baza danych.')
    wb = Workbook()
    wb.remove(wb.active)
    if scope in ('both', 'current'):
        drivers = [store.driver(driver_id)] if driver_id is not None else store.drivers()
        records = [slot for driver in drivers for slot in store.slots(driver['id']) if slot['status'] != 'Puste']
        write_sheet(wb.create_sheet('DANE'), records)
    if scope in ('both', 'history'):
        write_sheet(wb.create_sheet('ARCHIWUM'), store.history(driver_id), archive=True)
    if scope in ('both', 'current'):
        trucks = store.vehicles('truck', True)
        if driver_id is not None:
            trucks = [item for item in trucks if item['id'] == drivers[0]['truck_id']]
        if any(driver['phone'] for driver in drivers):
            write_optional_register(wb.create_sheet('KIEROWCY'), ['Kierowca', 'TIR', 'Telefon (opcjonalnie)'],
                                    [(driver['name'], driver['truck'], driver['phone']) for driver in drivers])
        if any(any(item[key] is not None and item[key] != '' for key in
                   ('odometer_km', 'inspection_date', 'insurance_date', 'service_date', 'notes')) for item in trucks):
            write_optional_register(wb.create_sheet('TIRY'), ['TIR', 'Przebieg km', 'Przegląd', 'Ubezpieczenie', 'Serwis', 'Uwagi'],
                [(item['plate'], item['odometer_km'], *(date.fromisoformat(item[key]) if item[key] else None
                  for key in ('inspection_date', 'insurance_date', 'service_date')), item['notes']) for item in trucks])
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(dir=path.parent, suffix='.xlsx')
    os.close(handle)
    try:
        wb.save(temporary)
        os.replace(temporary, path)
    finally:
        wb.close()
        if os.path.exists(temporary):
            os.unlink(temporary)
