"""Uruchamianie aplikacji; --db wskazuje własną bazę, --self-test testuje EXE."""

import argparse
import json
import sys
import tempfile
import traceback
from pathlib import Path

from fleet.store import Store
from fleet.branding import APP_NAME, VERSION
from fleet.paths import database_path, default_data_folder


def self_test(report, export_path=None):
    if export_path is not None and Path(export_path).resolve() == Path(report).resolve():
        raise ValueError('Raport i eksport testowy muszą być różnymi plikami.')
    import tkinter as tk
    from fleet.export import export_xlsx
    from fleet.gui import FleetApp
    with tempfile.TemporaryDirectory(prefix='flota-smoke-') as directory:
        store = Store(Path(directory) / 'smoke.sqlite3')
        driver = store.add_driver('Test Żółć', store.add_vehicle('truck', 'TEST1'))
        root = tk.Tk()
        root.withdraw()
        app = FleetApp(root, store)
        app.select_driver(driver)
        for key, value in dict(trailer='TEST2', info='Próba EXE', load_date='17.11.2026', unload_date='18.11.2026').items():
            app.fields[key].set(value)
        assert app.save_editor()
        assert len(store.history()) == 1
        app.clear_selected_slot()
        assert len(store.history()) == 1
        assert store.slots(driver)[0]['status'] == 'Puste'
        assert not app.slot_tree.get_children()
        for position in range(1, 5):
            store.save_slot(driver, position, info=f'Szkic {position}')
        for expected in range(5, 9):
            app.add_slot()
            assert app.position == expected
            app.fields['info'].set('Szkic ósmego wpisu' if expected == 8 else f'Szkic {expected}')
            assert app.save_editor()
        assert store.slots(driver)[7]['info'] == 'Szkic ósmego wpisu'
        app.fields['route'].set('Łódź → Berlin')
        app.fields['order_number'].set('ZL-EXE-8')
        assert app.save_editor()
        app.dark_mode.set(True)
        app.toggle_theme()
        assert store.get_setting('theme') == 'dark'
        assert app.fields['order_number'].get() == 'ZL-EXE-8'
        store.save_truck_details(store.driver(driver)['truck_id'], odometer_km='123456', inspection_date='18.11.2026')
        app.refresh_trucks()
        assert len(app.truck_tree.get_children()) == 1
        from datetime import date, timedelta
        store.save_slot(driver, 7, unload_date=(date.today() - timedelta(days=1)).isoformat())
        app.refresh_slots()
        assert 'overdue' in app.slot_tree.item('7', 'tags')
        app.deadline_colors.set(False)
        app.refresh_deadline_colors()
        assert 'overdue' not in app.slot_tree.item('7', 'tags')
        # Mierzymy także rzeczywisty układ spakowanego GUI bez pokazywania okna.
        root.attributes('-alpha', 0.0)
        root.geometry('1120x700')
        app.set_editor_visible(False)
        root.deiconify()
        root.update()
        assert app.slot_tree.winfo_width() >= root.winfo_width() * 0.9
        assert app.slot_tree.winfo_height() >= root.winfo_height() * 0.55
        app.set_editor_visible(True)
        root.update()
        assert app.trailer_combo.winfo_ismapped()
        root.withdraw()
        export_file = Path(export_path).resolve() if export_path is not None else Path(directory) / 'export.xlsx'
        export_xlsx(store, export_file)
        from openpyxl import load_workbook
        exported = load_workbook(export_file)
        try:
            for sheet in exported:
                if sheet.tables:
                    assert sheet.auto_filter.ref is None
        finally:
            exported.close()
        root.update()
        root.destroy()
        store.close()
    Path(report).write_text(json.dumps({'ok': True, 'app': APP_NAME, 'version': VERSION, 'frozen': bool(getattr(sys, 'frozen', False)),
                                      'checks': ['Tk GUI', 'SQLite v3', 'zapis formularza', 'historia', 'ukrywanie pustych miejsc', 'osiem miejsc',
                                                 'pola opcjonalne', 'ciemny motyw', 'kontrola tirów', 'kolory terminów i przełącznik', 'duża tabela i zwijany edytor', 'eksport XLSX']}, ensure_ascii=False, indent=2), encoding='utf-8')


def main():
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument('--version', action='version', version=f'{APP_NAME} {VERSION}')
    parser.add_argument('--db', type=Path, help='Ścieżka do własnej bazy SQLite')
    parser.add_argument('--self-test', metavar='REPORT', help='Uruchom izolowany test i zapisz raport JSON')
    parser.add_argument('--self-test-export', type=Path, metavar='XLSX', help='Zachowaj eksport autotestu we wskazanym pliku (wymaga --self-test)')
    args = parser.parse_args()
    if args.self_test_export is not None:
        if not args.self_test:
            parser.error('--self-test-export wymaga --self-test')
        if Path(args.self_test).resolve() == args.self_test_export.resolve():
            parser.error('Raport i eksport testowy muszą być różnymi plikami.')
    if args.self_test:
        self_test(args.self_test, args.self_test_export)
        return
    import tkinter as tk
    from fleet.gui import FleetApp
    folder = default_data_folder()
    store = None
    root = None
    try:
        store = Store(database_path(args.db))
        root = tk.Tk()
        FleetApp(root, store)
        root.mainloop()
    except Exception:
        folder.mkdir(parents=True, exist_ok=True)
        log = folder / 'blad.txt'
        log.write_text(traceback.format_exc(), encoding='utf-8')
        from tkinter import messagebox
        messagebox.showerror(f'{APP_NAME} — błąd', f'Nie udało się uruchomić programu. Szczegóły: {log}')
        raise
    finally:
        if store is not None:
            store.close()


if __name__ == '__main__':
    main()
