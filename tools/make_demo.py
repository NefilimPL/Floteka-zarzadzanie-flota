from pathlib import Path
from datetime import date, timedelta
import argparse
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from fleet.store import Store


def build(output=None):
    folder = Path(output).resolve() if output else ROOT / 'dist' / 'Floteka'
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / 'Floteka-DEMO.sqlite3'
    if target.exists():
        print('Baza DEMO już istnieje — zachowano jej dane.')
        return
    store = Store(target)
    try:
        truck = store.add_vehicle('truck', 'WX10001')
        first = store.add_driver('Kierowca pokazowy A', truck, phone='+48 000 000 000')
        store.add_driver('Kierowca pokazowy B', store.add_vehicle('truck', 'PO20002'))
        store.save_slot(first, 1, 'WX30003', 'PRZYKŁAD — firma A', '2026-11-17', '2026-11-19',
                        route='Poznań → Berlin', order_number='DEMO-001')
        store.save_slot(first, 2, 'WX40004', 'PRZYKŁAD — firma B', '2026-11-20', '2026-11-23')
        store.save_slot(first, 3, 'WX50005', 'PRZYKŁAD — szkic, uzupełnij daty')
        store.save_truck_details(truck, odometer_km='245000', inspection_date=(date.today() + timedelta(days=14)).isoformat(),
                                 service_date=(date.today() + timedelta(days=30)).isoformat(), notes='PRZYKŁAD — kontrola opon')
    finally:
        store.close()
    print('Utworzono oddzielną bazę DEMO.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path)
    build(parser.parse_args().output)
