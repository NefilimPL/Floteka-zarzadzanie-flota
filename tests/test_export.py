from datetime import datetime

from openpyxl import load_workbook
import pytest

from fleet.export import export_xlsx
from fleet.store import Store


def test_export_current_and_archive_with_dates_and_safe_text(tmp_path):
    store = Store(tmp_path / 'test.sqlite3')
    driver = store.add_driver('Łukasz', store.add_vehicle('truck', 'WX1'))
    store.save_slot(driver, 1, 'PO1', '=HYPERLINK("test")', '2026-11-17', '2026-11-20')
    target = tmp_path / 'export.xlsx'
    export_xlsx(store, target)
    wb = load_workbook(target)
    assert wb.sheetnames == ['DANE', 'ARCHIWUM']
    assert wb['DANE'].max_row == 2
    assert wb['ARCHIWUM'].max_row == 2
    assert wb['DANE']['D2'].data_type == 's'
    assert wb['DANE']['D2'].value.startswith('=HYPERLINK')
    assert isinstance(wb['DANE']['E2'].value, datetime)
    assert wb['DANE']['A2'].value == 'Łukasz'
    wb.close()
    store.close()


def test_export_selection(tmp_path):
    store = Store(tmp_path / 'test.sqlite3')
    truck = store.add_vehicle('truck', 'WX1')
    first = store.add_driver('A', truck)
    store.add_driver('B', truck)
    path = tmp_path / 'one.xlsx'
    export_xlsx(store, path, scope='current', driver_id=first)
    wb = load_workbook(path)
    assert wb.sheetnames == ['DANE']
    assert wb['DANE'].max_row == 1
    wb.close()
    store.close()


def test_export_contains_dynamic_deadline_rules(tmp_path):
    store = Store(tmp_path / 'colors.sqlite3')
    driver = store.add_driver('A', store.add_vehicle('truck', 'WX1'))
    store.save_slot(driver, 1, 'PO1', '', '2026-11-17', '2026-11-18')
    target = tmp_path / 'colors.xlsx'
    export_xlsx(store, target)
    wb = load_workbook(target)
    try:
        for sheet in wb:
            rules = [r for key in sheet.conditional_formatting for r in sheet.conditional_formatting[key]]
            assert len(rules) == 2
            assert rules[0].formula == ['AND(ISNUMBER($F2),$F2<TODAY())']
            assert rules[1].formula == ['AND(ISNUMBER($F2),$F2>=TODAY(),$F2<=TODAY()+1)']
            assert rules[0].dxf.fill.fgColor.rgb == '00E53935'
            assert rules[1].dxf.fill.fgColor.rgb == '00FFA726'
    finally:
        wb.close()
        store.close()


def test_export_refuses_truncation_and_keeps_previous_file(tmp_path):
    store = Store(tmp_path / 'long.sqlite3')
    driver = store.add_driver('A', store.add_vehicle('truck', 'WX1'))
    store.save_slot(driver, 1, 'PO1', 'ą' * 40000, '2026-11-17', '2026-11-18')
    target = tmp_path / 'previous.xlsx'
    target.write_bytes(b'previous contents')
    with pytest.raises(ValueError, match='32767'):
        export_xlsx(store, target)
    assert target.read_bytes() == b'previous contents'
    assert len(store.history()[0]['info']) == 40000
    store.close()


def test_optional_data_export_and_empty_rows_removed(tmp_path):
    store = Store(tmp_path / 'optional-export.sqlite3')
    truck = store.add_vehicle('truck', 'WX1')
    driver = store.add_driver('Anna', truck, phone='+48 123')
    store.save_slot(driver, 2, 'PO1', '', '2026-11-17', '2026-11-18', route='Łódź → Berlin', order_number='=ZL-1')
    store.save_truck_details(truck, odometer_km='123', inspection_date='17.12.2026', notes='Zimowe opony')
    target = tmp_path / 'optional.xlsx'
    export_xlsx(store, target)
    wb = load_workbook(target)
    try:
        assert wb.sheetnames == ['DANE', 'ARCHIWUM', 'KIEROWCY', 'TIRY']
        assert wb['DANE'].max_row == 2
        assert wb['DANE']['I2'].value == 'Łódź → Berlin'
        assert wb['DANE']['J2'].value == '=ZL-1'
        assert wb['DANE']['J2'].data_type == 's'
        assert wb['ARCHIWUM']['J2'].value == 'Łódź → Berlin'
        assert wb['ARCHIWUM']['K2'].value == '=ZL-1'
        assert wb['KIEROWCY']['C2'].value == '+48 123'
        assert wb['TIRY']['B2'].value == 123
        assert wb['TIRY']['C2'].value == datetime(2026, 12, 17)
        assert wb['TIRY']['F2'].value == 'Zimowe opony'
    finally:
        wb.close()
        store.close()


@pytest.mark.parametrize('scope', ['both', 'current', 'history'])
def test_excel_tables_do_not_overlap_with_worksheet_filters(tmp_path, scope):
    store = Store(tmp_path / 'filters.sqlite3')
    try:
        truck = store.add_vehicle('truck', 'WX1')
        driver = store.add_driver('Anna', truck, phone='+48 123')
        store.save_truck_details(truck, odometer_km='123')
        store.save_slot(driver, 1, 'PO1', 'Próba', '2026-11-17', '2026-11-18')
        target = tmp_path / 'filters.xlsx'
        export_xlsx(store, target, scope=scope)
        wb = load_workbook(target)
        try:
            for sheet in wb:
                if sheet.title in ('DANE', 'ARCHIWUM'):
                    assert len(sheet.tables) == 1
                    table = next(iter(sheet.tables.values()))
                    assert table.autoFilter.ref == table.ref
                    assert sheet.auto_filter.ref is None
                else:
                    assert not sheet.tables
                    assert sheet.auto_filter.ref == sheet.dimensions
        finally:
            wb.close()
    finally:
        store.close()


def test_empty_export_remains_filterable_without_tables(tmp_path):
    store = Store(tmp_path / 'empty.sqlite3')
    try:
        target = tmp_path / 'empty.xlsx'
        export_xlsx(store, target)
        wb = load_workbook(target)
        try:
            for sheet in wb:
                assert not sheet.tables
                assert sheet.max_row == 1
                assert sheet.auto_filter.ref == sheet.dimensions
        finally:
            wb.close()
    finally:
        store.close()
