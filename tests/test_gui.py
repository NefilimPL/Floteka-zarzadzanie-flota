import tkinter as tk
from datetime import date, timedelta

import pytest

from fleet.gui import FleetApp
from fleet.store import Store


@pytest.fixture
def app(tmp_path):
    root = tk.Tk()
    root.withdraw()
    store = Store(tmp_path / 'gui.sqlite3')
    driver = store.add_driver('Anna Żak', store.add_vehicle('truck', 'WX100'))
    gui = FleetApp(root, store)
    gui.select_driver(driver)
    root.update()
    yield gui, store, driver
    root.destroy()
    store.close()


def test_save_editor_updates_real_database_and_history_tree(app):
    gui, store, driver = app
    assert gui.root.title().startswith('Floteka')
    gui.select_slot(2)
    for key, value in dict(trailer='PO200', info='Firma', load_date='17.11.2026', unload_date='19.11.2026').items():
        gui.fields[key].set(value)
    assert gui.save_editor()
    assert store.slots(driver)[1]['status'] == 'Kompletny'
    assert len(store.history()) == 1
    gui.refresh_history()
    assert len(gui.history_tree.get_children()) == 1
    assert not gui.dirty


def test_saved_draft_stays_in_editor_after_refresh(app):
    gui, store, driver = app
    gui.fields['trailer'].set('PO1')
    assert gui.save_editor()
    gui.select_slot(3)
    gui.select_slot(1)
    assert gui.fields['trailer'].get() == 'PO1'
    assert not store.history()


def test_clear_complete_place_leaves_history(app):
    gui, store, driver = app
    store.save_slot(driver, 1, 'PO1', '', '2026-11-17', '2026-11-18')
    gui.select_driver(driver)
    gui.clear_selected_slot()
    assert store.slots(driver)[0]['status'] == 'Puste'
    assert len(store.history()) == 1


def test_empty_save_cancel_preserves_draft(app, monkeypatch):
    gui, store, driver = app
    store.save_slot(driver, 1, trailer='PO1')
    gui.select_driver(driver)
    gui.fields['trailer'].set('')
    prompts = []
    monkeypatch.setattr('fleet.gui.messagebox.askyesno', lambda *a, **k: prompts.append(a) or False)
    assert not gui.save_editor()
    assert prompts
    assert store.slots(driver)[0]['trailer'] == 'PO1'
    assert gui.dirty


def test_each_visible_slot_has_clear_action(app):
    gui, store, driver = app
    assert gui.slot_tree['columns'][-1] == '×'
    assert all(gui.slot_tree.item(row, 'values')[-1] == '×' for row in gui.slot_tree.get_children())


def test_add_slot_edits_fifth_record(app):
    gui, store, driver = app
    for position in range(1, 5):
        store.save_slot(driver, position, info=f'Szkic {position}')
    gui.add_slot()
    assert gui.position == 5
    gui.fields['info'].set('piąty szkic')
    assert gui.save_editor()
    assert store.slots(driver)[4]['info'] == 'piąty szkic'
    assert len(gui.slot_tree.get_children()) == 5


def test_empty_rows_hidden_and_cleared_row_reused(app):
    gui, store, driver = app
    assert not gui.slot_tree.get_children()
    gui.add_slot()
    assert gui.position == 1
    gui.fields['info'].set('szkic')
    assert gui.save_editor()
    assert gui.slot_tree.get_children() == ('1',)
    store.save_slot(driver, 1, trailer='PO1', load_date='2026-11-17', unload_date='2026-11-18')
    gui.refresh_slots()
    gui.clear_selected_slot()
    assert not gui.slot_tree.get_children()
    assert all(not field.get() for field in gui.fields.values())
    for _ in range(10):
        gui.add_slot()
    assert gui.position == 1
    assert len(store.slots(driver)) == 4
    assert len(store.history()) == 1


def test_dark_mode_preserves_unsaved_optional_fields_and_persists(app):
    gui, store, driver = app
    gui.fields['route'].set('Poznań → Berlin')
    gui.fields['order_number'].set('ZL-12')
    gui.dark_mode.set(True)
    gui.toggle_theme()
    assert gui.fields['route'].get() == 'Poznań → Berlin'
    assert gui.dirty
    assert store.get_setting('theme') == 'dark'
    assert gui.root.cget('background') == '#151b26'
    assert str(gui.slot_tree.tag_configure('overdue', 'background')) == '#E53935'
    assert str(gui.slot_tree.tag_configure('soon', 'background')) == '#FFA726'
    other = Store(store.path)
    try:
        assert other.get_setting('theme') == 'dark'
    finally:
        other.close()


def test_optional_entry_fields_restore_into_editor(app):
    gui, store, driver = app
    gui.fields['trailer'].set('PO1')
    gui.fields['load_date'].set('17.11.2026')
    gui.fields['unload_date'].set('18.11.2026')
    gui.fields['route'].set('Łódź → Gdańsk')
    gui.fields['order_number'].set('ZL-77')
    assert gui.save_editor()
    gui.select_slot(2)
    gui.select_slot(1)
    assert gui.fields['route'].get() == 'Łódź → Gdańsk'
    assert gui.fields['order_number'].get() == 'ZL-77'
    assert store.history()[0]['order_number'] == 'ZL-77'


def test_cancel_switch_from_hidden_new_entry_preserves_draft(app, monkeypatch):
    gui, store, driver = app
    store.save_slot(driver, 2, info='istniejący szkic')
    gui.select_driver(driver)
    gui.add_slot()
    assert gui.position == 1
    gui.fields['info'].set('niezapisany nowy szkic')
    monkeypatch.setattr('fleet.gui.messagebox.askyesnocancel', lambda *a, **k: None)
    gui.slot_tree.selection_set('2')
    gui._slot_selected()
    assert gui.position == 1
    assert not gui.slot_tree.selection()
    assert gui.fields['info'].get() == 'niezapisany nowy szkic'
    assert gui.dirty


def test_concurrent_change_keeps_unsaved_editor_on_conflict(app, monkeypatch):
    gui, store, driver = app
    store.save_slot(driver, 1, info='wersja początkowa')
    gui.select_driver(driver)
    gui.fields['info'].set('moja zmiana')
    store.save_slot(driver, 1, info='zmiana w drugim oknie')
    monkeypatch.setattr('fleet.gui.messagebox.showerror', lambda *a, **k: None)
    assert not gui.save_editor()
    assert gui.fields['info'].get() == 'moja zmiana'
    assert gui.dirty
    assert store.slots(driver)[0]['info'] == 'zmiana w drugim oknie'


def test_deadline_colors_can_be_disabled_without_losing_editor(app):
    gui, store, driver = app
    today = date.today()
    for position, days in enumerate((-1, 0, 1, 2), 1):
        store.save_slot(driver, position, unload_date=(today + timedelta(days=days)).isoformat())
    gui.refresh_slots()
    assert 'overdue' in gui.slot_tree.item('1', 'tags')
    assert 'soon' in gui.slot_tree.item('2', 'tags')
    assert 'soon' in gui.slot_tree.item('3', 'tags')
    assert not {'overdue', 'soon'} & set(gui.slot_tree.item('4', 'tags'))
    gui.fields['info'].set('niezapisane dane')
    gui.deadline_colors.set(False)
    gui.refresh_deadline_colors()
    assert all(not {'overdue', 'soon'} & set(gui.slot_tree.item(row, 'tags'))
               for row in gui.slot_tree.get_children())
    assert gui.fields['info'].get() == 'niezapisane dane'
    assert gui.dirty


@pytest.mark.parametrize('scaling', [1.33, 2.0])
def test_add_driver_button_visible_at_minimum_window_size(app, scaling):
    gui, store, driver = app
    old_scale = gui.root.tk.call('tk', 'scaling')
    try:
        gui.root.tk.call('tk', 'scaling', scaling)
        gui.root.geometry('1120x780')
        gui.root.deiconify()
        gui.root.update()
        buttons = []
        def find(widget):
            for child in widget.winfo_children():
                if child.winfo_class() == 'TButton' and child.cget('text') == '+ Dodaj kierowcę':
                    buttons.append(child)
                find(child)
        find(gui.root)
        assert len(buttons) == 1
        button = buttons[0]
        assert button.winfo_ismapped()
        assert button.winfo_height() >= button.winfo_reqheight()
        assert button.winfo_y() + button.winfo_height() <= button.master.winfo_height()
    finally:
        gui.root.tk.call('tk', 'scaling', old_scale)


def seed_visible_rows(gui, store, driver, count=24):
    for position in range(1, count + 1):
        if position > len(store.slots(driver)):
            store.add_slot(driver)
        store.save_slot(driver, position, info=f'Wpis {position}')
    gui.select_driver(driver)
    gui.root.deiconify()
    gui.root.update()
    gui.slot_tree.see('1')
    gui.root.update()


def full_visible_rows(tree):
    return sum(bool(box) and box[1] + box[3] <= tree.winfo_height()
               for box in (tree.bbox(row) for row in tree.get_children()))


@pytest.mark.parametrize('geometry', ['1120x700', '1366x700', '1600x748'])
def test_entry_table_uses_full_width_and_shows_ten_rows(app, geometry):
    gui, store, driver = app
    gui.root.geometry(geometry)
    seed_visible_rows(gui, store, driver)
    assert gui.slot_tree.winfo_width() >= gui.root.winfo_width() * 0.9
    assert full_visible_rows(gui.slot_tree) >= 10


def test_expanded_editor_still_shows_six_rows(app):
    gui, store, driver = app
    gui.root.geometry('1366x700')
    seed_visible_rows(gui, store, driver)
    gui.set_editor_visible(True)
    gui.optional_visible.set(True)
    gui.toggle_optional_fields()
    gui.root.update()
    assert full_visible_rows(gui.slot_tree) >= 6
    assert gui.trailer_combo.winfo_ismapped()


def test_collapsing_editor_keeps_unsaved_fields_and_reclaims_space(app):
    gui, store, driver = app
    gui.root.geometry('1120x700')
    seed_visible_rows(gui, store, driver)
    gui.set_editor_visible(True)
    gui.fields['info'].set('niezapisany opis')
    gui.fields['route'].set('Łódź → Poznań')
    gui.root.update()
    expanded_height = gui.slot_tree.winfo_height()
    gui.set_editor_visible(False)
    gui.root.update()
    assert not gui.editor.winfo_ismapped()
    assert gui.slot_tree.winfo_height() > expanded_height + 100
    gui.set_editor_visible(True)
    gui.root.update()
    assert gui.editor.winfo_ismapped()
    assert gui.fields['info'].get() == 'niezapisany opis'
    assert gui.fields['route'].get() == 'Łódź → Poznań'
    assert gui.dirty


def test_driver_picker_search_and_cancel_preserve_unsaved_entry(app, monkeypatch):
    gui, store, driver = app
    other = store.add_driver('Piotr Nowak', store.add_vehicle('truck', 'PO900'))
    gui.refresh_drivers()
    gui.driver_search.set('PO900')
    choices = list(gui.driver_combo['values'])
    assert len(choices) == 1
    assert 'Piotr Nowak' in choices[0]
    gui.fields['info'].set('szkic przed wyborem')
    monkeypatch.setattr('fleet.gui.messagebox.askyesnocancel', lambda *a, **k: None)
    gui.driver_choice.set(choices[0])
    gui._driver_selected()
    assert gui.driver_id == driver
    assert gui.fields['info'].get() == 'szkic przed wyborem'
    assert 'Anna Żak' in gui.driver_choice.get()
    assert gui.dirty


def test_add_entry_opens_collapsed_editor(app):
    gui, store, driver = app
    gui.root.deiconify()
    gui.root.update()
    assert not gui.editor.winfo_ismapped()
    gui.add_slot()
    gui.root.update()
    assert gui.editor.winfo_ismapped()


@pytest.mark.parametrize('width', [1040, 1120])
def test_large_dpi_keeps_all_visible_buttons_inside_window(tmp_path, width):
    root = tk.Tk()
    root.withdraw()
    old_scale = root.tk.call('tk', 'scaling')
    root.tk.call('tk', 'scaling', 8 / 3)  # 200%: ustawione przed budową kontrolek.
    store = Store(tmp_path / 'large-dpi.sqlite3')
    try:
        driver = store.add_driver('Anna Żak', store.add_vehicle('truck', 'WX1'))
        gui = FleetApp(root, store)
        gui.select_driver(driver)
        gui.set_editor_visible(True)
        gui.optional_visible.set(True)
        gui.toggle_optional_fields()
        root.geometry(f'{width}x700')
        root.attributes('-alpha', 0.0)
        root.deiconify()
        root.update()
        def check(widget):
            for child in widget.winfo_children():
                if child.winfo_class() == 'TButton' and child.winfo_ismapped():
                    right = child.winfo_rootx() - root.winfo_rootx() + child.winfo_width()
                    assert right <= root.winfo_width(), child.cget('text')
                    assert child.winfo_width() >= child.winfo_reqwidth(), child.cget('text')
                    assert child.winfo_height() >= child.winfo_reqheight(), child.cget('text')
                check(child)
        check(root)
    finally:
        root.tk.call('tk', 'scaling', old_scale)
        root.destroy()
        store.close()
