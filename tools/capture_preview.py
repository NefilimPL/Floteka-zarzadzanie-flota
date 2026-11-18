"""Podglądy wyłącznie okna Floteki na odrębnych, fikcyjnych danych."""

import sys
import json
import tempfile
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import tkinter as tk
from PIL import ImageGrab
from fleet.gui import FleetApp
from fleet.store import Store


def capture():
    with tempfile.TemporaryDirectory(prefix='floteka-preview-') as folder:
        store = Store(Path(folder) / 'preview.sqlite3')
        truck = store.add_vehicle('truck', 'WX10001')
        driver = store.add_driver('Kierowca pokazowy', truck, phone='+48 123 456 789')
        today = date.today()
        for position in range(1, 25):
            days = (-1, 0, 1, 3)[(position - 1) % 4]
            if position > 4:
                store.add_slot(driver)
            store.save_slot(driver, position, 'PO' + str(100 + position), 'Przykładowy transport',
                            (today - timedelta(days=3)).isoformat(), (today + timedelta(days=days)).isoformat(),
                            route='Poznań → Berlin', order_number='ZL-' + str(position))
        store.save_truck_details(truck, odometer_km='245000', inspection_date=(today + timedelta(days=14)).isoformat(),
                                 service_date=(today + timedelta(days=30)).isoformat(), notes='Sprawdzić opony')
        root = tk.Tk()
        try:
            app = FleetApp(root, store)
            app.select_driver(driver)
            app.select_slot(3)
            root.geometry('1366x700+20+20')
            root.attributes('-topmost', True)
            metrics = {}
            def measure(name):
                app.slot_tree.yview_moveto(0)
                root.update()
                boxes = [app.slot_tree.bbox(row) for row in app.slot_tree.get_children()]
                metrics[name] = {'width': app.slot_tree.winfo_width(), 'height': app.slot_tree.winfo_height(),
                                 'full_visible_rows': sum(bool(box) and box[1] + box[3] <= app.slot_tree.winfo_height() for box in boxes)}
            for dark, name in ((False, 'podglad-jasny.png'), (True, 'podglad-ciemny.png')):
                app.dark_mode.set(dark)
                app.toggle_theme()
                app.slot_tree.yview_moveto(0)
                root.update()
                ImageGrab.grab(window=root.winfo_id()).save(ROOT / 'docs' / name)
            measure('collapsed')
            app.optional_visible.set(True)
            app.toggle_optional_fields()
            root.update()
            ImageGrab.grab(window=root.winfo_id()).save(ROOT / 'docs' / 'podglad-edycja.png')
            measure('expanded_with_optional')
            (ROOT / 'docs' / 'gui-layout.json').write_text(json.dumps(metrics, indent=2), encoding='utf-8')
            print(json.dumps(metrics), flush=True)
            app.notebook.select(3)
            root.update()
            ImageGrab.grab(window=root.winfo_id()).save(ROOT / 'docs' / 'podglad-kontrola-tirow.png')
        finally:
            root.destroy()
            store.close()


if __name__ == '__main__':
    capture()
