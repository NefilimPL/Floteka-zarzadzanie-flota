"""Polski interfejs lokalnej aplikacji — bez serwera i przeglądarki."""

from __future__ import annotations

import os
import sqlite3
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk

from .export import export_xlsx
from .store import DraftConfirmationRequired
from .branding import APP_NAME, VERSION
from .deadlines import deadline_tag, OVERDUE_FILL, SOON_FILL
from .theme import apply_theme


def display_date(value):
    return date.fromisoformat(value).strftime('%d.%m.%Y') if value else '—'


class FleetApp:
    def __init__(self, root, store):
        self.root, self.store = root, store
        self.driver_id = None
        self.position = 1
        self.dirty = False
        self.loading = False
        self.driver_map = {}
        self.driver_choices = {}
        self.driver_labels = {}
        self.history_map = {}
        self.fields = {k: tk.StringVar() for k in ('trailer', 'info', 'load_date', 'unload_date', 'route', 'order_number')}
        self.status = tk.StringVar(value='Gotowe. Wybierz kierowcę i miejsce albo dodaj dane w Rejestrach.')
        self.driver_search = tk.StringVar()
        self.driver_choice = tk.StringVar()
        self.history_search = tk.StringVar()
        self.history_filter = tk.StringVar(value='Wszyscy kierowcy')
        self.deadline_colors = tk.BooleanVar(value=True)
        self.dark_mode = tk.BooleanVar(value=store.get_setting('theme', 'light') == 'dark')
        self.optional_visible = tk.BooleanVar(value=False)
        self.editor_visible = False
        self.truck_summary = tk.StringVar()
        root.title(f'{APP_NAME} {VERSION} • kierowcy, tiry i przyczepy')
        root.geometry(f'{min(1280, max(1040, root.winfo_screenwidth() - 60))}x{min(800, max(600, root.winfo_screenheight() - 100))}')
        root.minsize(1040, 600)
        self._styles()
        self._build()
        for variable in self.fields.values():
            variable.trace_add('write', self._changed)
        self.refresh_all()
        root.protocol('WM_DELETE_WINDOW', self.close)
        root.bind('<Control-s>', lambda event: self.save_editor())
        self._deadline_timer = root.after(60_000, self._refresh_deadline_clock)
        root.bind('<Destroy>', self._cancel_deadline_timer, add='+')

    def _styles(self):
        self.palette = apply_theme(self.root, self.dark_mode.get())
        if hasattr(self, 'slot_tree'):
            # Starsze tagi statusu mają w Tk pierwszeństwo przed tagiem terminu.
            self.slot_tree.tag_configure('Szkic', foreground='')
            self.slot_tree.tag_configure('Kompletny', foreground='')

    def toggle_theme(self):
        try:
            self.store.set_setting('theme', 'dark' if self.dark_mode.get() else 'light')
        except sqlite3.Error as exc:
            self._error(exc)
            return
        self._styles()

    def toggle_optional_fields(self):
        if self.optional_visible.get():
            self.set_editor_visible(True)
            self.optional_editor.grid()
        else:
            self.optional_editor.grid_remove()

    def set_editor_visible(self, visible):
        self.editor_visible = bool(visible)
        if self.editor_visible:
            self.editor.grid()
        else:
            self.editor.grid_remove()
        self.editor_toggle.configure(text='Zwiń edytor' if self.editor_visible else 'Edytuj wpis')
        self._reflow_current_actions()

    def edit_selected_slot(self, event=None):
        if self.driver_id is None:
            self.status.set('Najpierw wybierz lub dodaj kierowcę.')
            return
        if event is not None:
            row = self.slot_tree.identify_row(event.y)
            if not row:
                return
            if int(row) != self.position:
                if not self._guard():
                    return
                self.select_slot(int(row))
        self.set_editor_visible(True)
        self.trailer_combo.focus_set()

    def toggle_editor(self):
        if self.editor_visible:
            self.set_editor_visible(False)
        else:
            self.edit_selected_slot()

    def _reflow_current_actions(self, event=None):
        required = sum(widget.winfo_reqwidth() for widget in self.current_action_widgets) + 48
        compact = self.current_actions.winfo_width() < required
        if getattr(self, '_actions_compact', None) == compact:
            return
        self._actions_compact = compact
        oldest, clear = self.current_action_widgets[-2:]
        oldest.grid_configure(row=1 if compact else 0, column=1 if compact else 5, pady=(4, 0) if compact else 0)
        clear.grid_configure(row=1 if compact else 0, column=2 if compact else 6, pady=(4, 0) if compact else 0)

    def _reflow_editor_controls(self, event=None):
        required = sum(widget.winfo_reqwidth() for widget in self.editor_control_widgets) + 32
        compact = self.editor_controls.winfo_width() < required
        if getattr(self, '_editor_controls_compact', None) == compact:
            return
        self._editor_controls_compact = compact
        _, _, hint, optional = self.editor_control_widgets
        hint.grid_configure(row=1 if compact else 0, column=1 if compact else 2,
                            columnspan=3 if compact else 1, pady=(4, 0) if compact else 0)
        optional.grid_configure(row=1 if compact else 0, column=0 if compact else 3,
                                sticky='w', pady=(4, 0) if compact else 0)

    def _build(self):
        top = ttk.Frame(self.root, padding=(14, 8))
        top.pack(fill='x')
        ttk.Label(top, text=APP_NAME, style='Title.TLabel').pack(side='left')
        ttk.Label(top, text='Zarządzanie flotą', style='Subtitle.TLabel').pack(side='left', padx=14)
        ttk.Button(top, text='Eksport do Excela', style='Compact.Accent.TButton', command=self.export_dialog).pack(side='right')
        ttk.Button(top, text='Kopia bazy', style='Compact.TButton', command=self.backup).pack(side='right', padx=8)
        ttk.Checkbutton(top, text='Tryb ciemny', variable=self.dark_mode, command=self.toggle_theme).pack(side='right', padx=12)
        bottom = ttk.Frame(self.root, padding=(14, 4, 14, 8))
        bottom.pack(side='bottom', fill='x')
        ttk.Button(bottom, text='Folder danych', style='Compact.TButton', command=lambda: os.startfile(self.store.path.parent)).pack(side='right', padx=(12, 0))
        ttk.Label(bottom, textvariable=self.status, wraplength=780).pack(side='left', fill='x', expand=True)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=(0, 2))
        current, archive, registers, trucks = (ttk.Frame(self.notebook, padding=8) for _ in range(4))
        self.notebook.add(current, text='Bieżące wpisy')
        self.notebook.add(archive, text='Historia')
        self.notebook.add(registers, text='Rejestry')
        self.notebook.add(trucks, text='Kontrola tirów')
        self._build_current(current)
        self._build_history(archive)
        self._build_registers(registers)
        self._build_trucks(trucks)
        for tree in (self.slot_tree, self.history_tree, self.truck_tree):
            tree.tag_configure('overdue', background='#' + OVERDUE_FILL, foreground='#ffffff')
            tree.tag_configure('soon', background='#' + SOON_FILL, foreground='#211300')
        self._styles()

    def _tree(self, parent, columns, widths, height=8):
        frame = ttk.Frame(parent)
        frame.pack(fill='both', expand=True)
        tree = ttk.Treeview(frame, columns=columns, show='headings', height=height, selectmode='browse')
        for column, width in zip(columns, widths):
            tree.heading(column, text=column)
            tree.column(column, width=width, minwidth=65, stretch=True)
        yscroll = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        xscroll = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _build_current(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        driver_bar = ttk.Frame(parent)
        driver_bar.grid(row=0, column=0, sticky='ew', pady=(0, 6))
        driver_bar.columnconfigure(1, weight=1)
        ttk.Label(driver_bar, text='Kierowca / TIR').grid(row=0, column=0, padx=(0, 8))
        self.driver_combo = ttk.Combobox(driver_bar, textvariable=self.driver_choice, state='readonly', width=34)
        self.driver_combo.grid(row=0, column=1, sticky='ew')
        self.driver_combo.bind('<<ComboboxSelected>>', self._driver_selected)
        ttk.Label(driver_bar, text='Szukaj').grid(row=0, column=2, padx=(14, 6))
        ttk.Entry(driver_bar, textvariable=self.driver_search, width=20).grid(row=0, column=3, sticky='ew')
        ttk.Button(driver_bar, text='+ Dodaj kierowcę', style='Compact.TButton', command=self.add_driver_dialog).grid(row=0, column=4, padx=(12, 0))
        self.driver_search.trace_add('write', lambda *_: self.refresh_drivers())
        actions = ttk.Frame(parent)
        self.current_actions = actions
        actions.grid(row=1, column=0, sticky='ew', pady=(0, 6))
        actions.columnconfigure(3, weight=1)
        self.current_title = ttk.Label(actions, text='Wpisy', style='Heading.TLabel')
        self.current_title.grid(row=0, column=0, sticky='w', padx=(0, 12))
        add = ttk.Button(actions, text='+ Dodaj wpis', style='Compact.Accent.TButton', command=self.add_slot)
        add.grid(row=0, column=1, padx=(0, 6))
        self.editor_toggle = ttk.Button(actions, text='Edytuj wpis', style='Compact.TButton', command=self.toggle_editor)
        self.editor_toggle.grid(row=0, column=2)
        colors = ttk.Checkbutton(actions, text='Kolory terminów', variable=self.deadline_colors,
                                 command=self.refresh_deadline_colors)
        colors.grid(row=0, column=4, padx=10)
        oldest = ttk.Button(actions, text='Zwolnij najstarsze', style='Compact.TButton', command=self.clear_oldest)
        oldest.grid(row=0, column=5, padx=(0, 6))
        clear = ttk.Button(actions, text='× Wyczyść wybrane', style='Compact.TButton', command=self.clear_selected_slot)
        clear.grid(row=0, column=6)
        self.current_action_widgets = (self.current_title, add, self.editor_toggle, colors, oldest, clear)
        actions.bind('<Configure>', self._reflow_current_actions)
        list_area = ttk.Frame(parent)
        list_area.grid(row=2, column=0, sticky='nsew')
        self.slot_tree = self._tree(list_area, ('Miejsce', 'Przyczepa', 'Informacje', 'Załadunek', 'Rozładunek', 'Status', '×'),
                                    (65, 125, 320, 115, 115, 180, 35), 14)
        self.slot_tree.column('×', width=35, minwidth=35, stretch=False, anchor='center')
        self.slot_tree.column('Status', width=180, minwidth=150)
        self.slot_tree.bind('<<TreeviewSelect>>', self._slot_selected)
        self.slot_tree.bind('<Button-1>', self._slot_click)
        self.slot_tree.bind('<Double-1>', self.edit_selected_slot)
        self.slot_tree.tag_configure('Szkic', foreground='#9D6713')
        self.slot_tree.tag_configure('Kompletny', foreground='#0C7370')
        self.editor = ttk.LabelFrame(parent, text='Nowy wpis', padding=(10, 6))
        self.editor.grid(row=3, column=0, sticky='ew', pady=(6, 0))
        for column in (1, 3, 5):
            self.editor.columnconfigure(column, weight=2 if column == 1 else 1)
        for column, (key, label) in enumerate((('trailer', 'Przyczepa'), ('load_date', 'Załadunek'), ('unload_date', 'Rozładunek'))):
            ttk.Label(self.editor, text=label).grid(row=0, column=column * 2, sticky='w', padx=(8 if column else 0, 6), pady=3)
            if key == 'trailer':
                self.trailer_combo = ttk.Combobox(self.editor, textvariable=self.fields[key], width=20)
                control = self.trailer_combo
            else:
                control = ttk.Entry(self.editor, textvariable=self.fields[key], width=14)
            control.grid(row=0, column=column * 2 + 1, sticky='ew', pady=3)
        ttk.Label(self.editor, text='Informacje').grid(row=1, column=0, sticky='w', padx=(0, 6), pady=3)
        ttk.Entry(self.editor, textvariable=self.fields['info']).grid(row=1, column=1, columnspan=5, sticky='ew', pady=3)
        controls = ttk.Frame(self.editor)
        self.editor_controls = controls
        controls.grid(row=3, column=0, columnspan=6, sticky='ew', pady=(4, 0))
        controls.columnconfigure(2, weight=1)
        save = ttk.Button(controls, text='Zapisz szkic / wpis   Ctrl+S', style='Compact.Accent.TButton', command=self.save_editor)
        save.grid(row=0, column=0)
        undo = ttk.Button(controls, text='Cofnij niezapisane zmiany', style='Compact.TButton', command=self.load_editor)
        undo.grid(row=0, column=1, padx=(8, 0))
        hint = ttk.Label(controls, text='Daty: DD.MM.RRRR', style='Subtitle.TLabel')
        hint.grid(row=0, column=2, sticky='w', padx=12)
        optional = ttk.Checkbutton(controls, text='Dane opcjonalne wpisu', variable=self.optional_visible,
                                  command=self.toggle_optional_fields)
        optional.grid(row=0, column=3)
        self.editor_control_widgets = (save, undo, hint, optional)
        controls.bind('<Configure>', self._reflow_editor_controls)
        self.optional_editor = ttk.Frame(self.editor)
        self.optional_editor.grid(row=2, column=0, columnspan=6, sticky='ew', pady=3)
        self.optional_editor.columnconfigure(1, weight=1)
        self.optional_editor.columnconfigure(3, weight=1)
        for column, (key, label) in enumerate((('route', 'Trasa'), ('order_number', 'Numer zlecenia'))):
            ttk.Label(self.optional_editor, text=label).grid(row=0, column=column * 2, sticky='w', padx=(12 if column else 0, 6))
            ttk.Entry(self.optional_editor, textvariable=self.fields[key], width=20).grid(row=0, column=column * 2 + 1, sticky='ew')
        self.optional_editor.grid_remove()
        self.set_editor_visible(False)

    def _build_history(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill='x', pady=(0, 10))
        ttk.Label(top, text='Pełna historia', style='Heading.TLabel').pack(side='left')
        self.history_combo = ttk.Combobox(top, textvariable=self.history_filter, state='readonly', width=29)
        self.history_combo.pack(side='right', padx=8)
        self.history_combo.bind('<<ComboboxSelected>>', lambda event: self.refresh_history())
        search = ttk.Entry(top, textvariable=self.history_search, width=25)
        search.pack(side='right')
        search.bind('<Return>', lambda event: self.refresh_history())
        ttk.Button(top, text='Szukaj', command=self.refresh_history).pack(side='right', padx=8)
        self.history_tree = self._tree(parent, ('ID', 'Kierowca', 'TIR', 'Przyczepa', 'Informacje', 'Załadunek', 'Rozładunek', 'Zapisano UTC', 'Trasa', 'Numer zlecenia'),
                                       (55, 180, 100, 100, 230, 110, 110, 175, 230, 150), 14)
        bottom = ttk.Frame(parent)
        bottom.pack(fill='x', pady=(10, 0))
        ttk.Label(bottom, text='Kopie są niezmienne. Poprawki tworzą nowe wersje, bez duplikatów.',
                  style='Subtitle.TLabel').pack(side='left')
        ttk.Button(bottom, text='Przywróć do pustego miejsca', command=self.restore_dialog).pack(side='right')

    def _build_registers(self, parent):
        ttk.Label(parent, text='Dodawaj kierowców i pojazdy. Wyłączenie z listy zachowuje historię.',
                  style='Subtitle.TLabel').pack(anchor='w', pady=(0, 12))
        panels = ttk.Frame(parent)
        panels.pack(fill='both', expand=True)
        for col in range(3):
            panels.columnconfigure(col, weight=1)
        panels.rowconfigure(0, weight=1)
        self.register_trees = {}
        for col, (kind, title, columns, widths) in enumerate([
            ('driver', 'Kierowcy', ('ID', 'Kierowca', 'TIR', 'Aktywny', 'Telefon'), (40, 155, 90, 65, 140)),
            ('truck', 'TIR-y', ('ID', 'Rejestracja', 'Aktywny'), (45, 135, 70)),
            ('trailer', 'Przyczepy', ('ID', 'Rejestracja', 'Aktywny'), (45, 135, 70)),
        ]):
            frame = ttk.LabelFrame(panels, text=title, padding=8)
            frame.grid(row=0, column=col, sticky='nsew', padx=(0, 10) if col < 2 else 0)
            list_area = ttk.Frame(frame)
            command = self.add_driver_dialog if kind == 'driver' else lambda k=kind: self.add_vehicle_dialog(k)
            controls = ttk.Frame(frame)
            controls.pack(side='bottom', fill='x')
            ttk.Button(controls, text='+ Dodaj', command=command).pack(fill='x', pady=(8, 0))
            if kind == 'driver':
                ttk.Button(controls, text='Zmień nazwę / przypisz tira', command=self.edit_driver_dialog).pack(fill='x', pady=(5, 0))
            elif kind == 'truck':
                ttk.Button(controls, text='Dane opcjonalne tira', command=lambda: self.truck_details_dialog(from_register=True)).pack(fill='x', pady=(5, 0))
            ttk.Button(controls, text='Włącz / wyłącz z listy', command=lambda k=kind: self.toggle_active(k)).pack(fill='x', pady=(5, 0))
            list_area.pack(fill='both', expand=True)
            self.register_trees[kind] = self._tree(list_area, columns, widths, 12)

    def _build_trucks(self, parent):
        ttk.Label(parent, text='Terminy i dane tirów — wszystkie pola są opcjonalne', style='Heading.TLabel').pack(anchor='w')
        ttk.Label(parent, text='Czerwony: termin minął. Pomarańczowy: przegląd, ubezpieczenie lub serwis w ciągu 30 dni.',
                  style='Subtitle.TLabel').pack(anchor='w', pady=(5, 8))
        ttk.Label(parent, textvariable=self.truck_summary).pack(anchor='w', pady=(0, 8))
        controls = ttk.Frame(parent)
        controls.pack(side='bottom', fill='x', pady=(8, 0))
        ttk.Button(controls, text='Edytuj dane opcjonalne tira', style='Accent.TButton', command=self.truck_details_dialog).pack(side='left')
        self.truck_tree = self._tree(parent, ('TIR', 'Przebieg km', 'Przegląd', 'Ubezpieczenie', 'Serwis', 'Uwagi'),
                                    (130, 110, 130, 130, 130, 340), 12)
        self.truck_tree.bind('<Double-1>', lambda event: self.truck_details_dialog())

    def refresh_trucks(self):
        selected = self.truck_tree.selection()
        self.truck_tree.delete(*self.truck_tree.get_children())
        overdue = soon = 0
        today = date.today()
        trucks = self.store.vehicles('truck')
        for truck in trucks:
            days = [(date.fromisoformat(truck[key]) - today).days for key in
                    ('inspection_date', 'insurance_date', 'service_date') if truck[key]]
            tag = 'overdue' if any(value < 0 for value in days) else 'soon' if any(value <= 30 for value in days) else ''
            overdue += tag == 'overdue'
            soon += tag == 'soon'
            self.truck_tree.insert('', 'end', iid=str(truck['id']), tags=(tag,) if tag else (), values=(
                truck['plate'], truck['odometer_km'] if truck['odometer_km'] is not None else '—',
                *(display_date(truck[key]) for key in ('inspection_date', 'insurance_date', 'service_date')), truck['notes']))
        if selected and self.truck_tree.exists(selected[0]):
            self.truck_tree.selection_set(selected[0])
        self.truck_summary.set(f'Aktywne tiry: {len(trucks)}   •   Po terminie: {overdue}   •   Termin do 30 dni: {soon}')

    def truck_details_dialog(self, from_register=False):
        selected = (self.register_trees['truck'] if from_register else self.truck_tree).selection()
        if not selected or not self._guard():
            return
        truck = next(item for item in self.store.vehicles('truck', True) if item['id'] == int(selected[0]))
        definitions = [(key, label + ' (opcjonalnie)',
                        display_date(truck[key]) if key.endswith('_date') and truck[key] else str(truck[key]) if truck[key] is not None else '', None) for key, label in (
            ('odometer_km', 'Przebieg w km'), ('inspection_date', 'Termin przeglądu'),
            ('insurance_date', 'Ubezpieczenie ważne do'), ('service_date', 'Termin serwisu'), ('notes', 'Uwagi'))]
        self.form_dialog(f"Dane tira {truck['plate']} — daty DD.MM.RRRR", definitions,
                         lambda values: self.store.save_truck_details(truck['id'], **values))

    def _changed(self, *_):
        if not self.loading:
            self.dirty = True
            self.status.set('Niezapisane zmiany — naciśnij „Zapisz szkic / wpis”.')

    def _error(self, exc):
        messagebox.showerror('Nie udało się wykonać operacji', str(exc), parent=self.root)

    def _guard(self):
        if not self.dirty:
            return True
        answer = messagebox.askyesnocancel('Niezapisane zmiany', 'Zapisać zmiany w edytowanym miejscu?\nTak — zapisz, Nie — pomiń zmiany, Anuluj — zostań.', parent=self.root)
        if answer is None:
            return False
        if answer:
            return self.save_editor()
        self.dirty = False
        return True

    def refresh_all(self):
        self.refresh_drivers()
        self.refresh_registers()
        self.refresh_history()
        self.refresh_trucks()
        self.trailer_combo['values'] = [v['plate'] for v in self.store.vehicles('trailer')]
        if self.driver_id is not None:
            self.refresh_slots()

    def refresh_drivers(self):
        needle = self.driver_search.get().casefold().strip()
        self.driver_map = {d['id']: d for d in self.store.drivers()}
        self.driver_choices = {}
        self.driver_labels = {}
        choices = []
        for driver in self.driver_map.values():
            label = f"{driver['name']}   •   {driver['truck']}"
            if label in self.driver_choices:
                label += f" [#{driver['id']}]"
            self.driver_choices[label] = driver['id']
            self.driver_labels[driver['id']] = label
            if not needle or needle in f"{driver['name']} {driver['truck']}".casefold():
                choices.append(label)
        self.driver_combo['values'] = choices
        self.driver_choice.set(self.driver_labels.get(self.driver_id, ''))

    def select_driver(self, driver_id):
        self.driver_id = driver_id
        self.position = next((slot['position'] for slot in self.store.slots(driver_id) if slot['status'] != 'Puste'), 1)
        self.refresh_slots()
        self.refresh_drivers()

    def _driver_selected(self, event=None):
        selected = self.driver_choices.get(self.driver_choice.get())
        if self.loading or selected is None or selected == self.driver_id:
            return
        if self._guard():
            self.select_driver(selected)
        else:
            self.driver_choice.set(self.driver_labels.get(self.driver_id, ''))

    def refresh_slots(self):
        self.loading = True
        self.slot_tree.delete(*self.slot_tree.get_children())
        if self.driver_id is None:
            self.current_title.configure(text='Wpisy: 0')
        else:
            occupied = [slot for slot in self.store.slots(self.driver_id) if slot['status'] != 'Puste']
            self.current_title.configure(text=f'Wpisy: {len(occupied)}')
            for slot in occupied:
                self.slot_tree.insert('', 'end', iid=str(slot['position']), tags=(slot['status'],),
                                      values=(slot['position'], slot['trailer'] or '—', slot['info'],
                                              display_date(slot['load_date']), display_date(slot['unload_date']), slot['status'], '×'))
            if self.slot_tree.exists(str(self.position)):
                self.slot_tree.selection_set(str(self.position))
                self.slot_tree.see(str(self.position))
        self.loading = False
        self._reflow_current_actions()
        self.refresh_deadline_colors()
        self.load_editor()

    def add_slot(self):
        if self.driver_id is None or not self._guard():
            return
        try:
            empty = next((slot['position'] for slot in self.store.slots(self.driver_id) if slot['status'] == 'Puste'), None)
            self.position = empty if empty is not None else self.store.add_slot(self.driver_id)
        except (ValueError, sqlite3.Error) as exc:
            self._error(exc)
            return
        self.refresh_slots()
        self.set_editor_visible(True)
        self.trailer_combo.focus_set()
        self.status.set('Uzupełnij nowy wpis i zapisz. Puste miejsca nie zajmują listy.')

    def refresh_deadline_colors(self):
        # Zmieniamy tylko tagi, bez przeładowania niezapisanego formularza.
        slots = self.store.slots(self.driver_id) if self.driver_id is not None else []
        for row in slots:
            tag = deadline_tag(row['unload_date']) if self.deadline_colors.get() else ''
            if self.slot_tree.exists(str(row['position'])):
                self.slot_tree.item(str(row['position']), tags=(tag, row['status']) if tag else (row['status'],))
                values = list(self.slot_tree.item(str(row['position']), 'values'))
                values[5] = row['status'] + (' • PO TERMINIE' if tag == 'overdue' else ' • DZIŚ / JUTRO' if tag == 'soon' else '')
                self.slot_tree.item(str(row['position']), values=values)
        for row in self.store.history():
            if self.history_tree.exists(str(row['id'])):
                tag = deadline_tag(row['unload_date']) if self.deadline_colors.get() else ''
                self.history_tree.item(str(row['id']), tags=(tag,) if tag else ())

    def _refresh_deadline_clock(self):
        self.refresh_deadline_colors()
        self.refresh_trucks()
        self._deadline_timer = self.root.after(60_000, self._refresh_deadline_clock)

    def _cancel_deadline_timer(self, event):
        if event.widget is self.root:
            self.root.after_cancel(self._deadline_timer)

    def select_slot(self, position):
        self.position = position
        if self.slot_tree.exists(str(position)):
            self.slot_tree.selection_set(str(position))
            self.slot_tree.see(str(position))
        self.load_editor()

    def _slot_selected(self, event=None):
        selected = self.slot_tree.selection()
        if self.loading or not selected or int(selected[0]) == self.position:
            return
        if self._guard():
            self.select_slot(int(selected[0]))
        elif self.slot_tree.exists(str(self.position)):
            self.slot_tree.selection_set(str(self.position))
        else:
            self.slot_tree.selection_remove(*self.slot_tree.selection())

    def _slot_click(self, event):
        row = self.slot_tree.identify_row(event.y)
        if row and self.slot_tree.identify_column(event.x) == '#7':
            if self._guard():
                self.select_slot(int(row))
                self.clear_selected_slot()
            return 'break'

    def load_editor(self):
        self.loading = True
        data = self.store.slots(self.driver_id)[self.position - 1] if self.driver_id is not None else {}
        self.editor_snapshot = dict(data)
        for key, variable in self.fields.items():
            value = data.get(key, '')
            if key.endswith('date') and value:
                value = display_date(value)
            variable.set(value)
        self.editor.configure(text=f'Wpis {self.position} — edycja' if data.get('status') != 'Puste' and data else 'Nowy wpis')
        self.loading = False
        self.dirty = False

    def save_editor(self):
        if self.driver_id is None:
            self.status.set('Najpierw wybierz lub dodaj kierowcę.')
            return False
        try:
            values = {k: v.get() for k, v in self.fields.items()}
            try:
                position = self.store.save_slot(self.driver_id, self.position, **values, expected=self.editor_snapshot)
            except DraftConfirmationRequired as exc:
                if not messagebox.askyesno('Wyczyścić niekompletny szkic?', str(exc), parent=self.root):
                    return False
                position = self.store.save_slot(self.driver_id, self.position, **values, allow_draft=True, expected=self.editor_snapshot)
        except (ValueError, sqlite3.Error, OSError) as exc:
            self._error(exc)
            return False
        self.dirty = False
        self.position = position
        self.refresh_all()
        slot = self.store.slots(self.driver_id)[self.position - 1]
        self.status.set('Zapisano. Kompletne dane są już w historii.' if slot['status'] == 'Kompletny'
                        else 'Zapisano szkic. Uzupełnij przyczepę i obie daty, aby zapisać historię.' if slot['status'] == 'Szkic' else 'Zapisano puste miejsce.')
        return True

    def _clear(self, oldest=False):
        if self.driver_id is None or not self._guard():
            return
        operation = self.store.clear_oldest if oldest else self.store.clear_slot
        args = (self.driver_id,) if oldest else (self.driver_id, self.position)
        try:
            try:
                operation(*args)
            except DraftConfirmationRequired as exc:
                if not messagebox.askyesno('Wyczyścić niekompletny szkic?', str(exc), parent=self.root):
                    return
                operation(*args, allow_draft=True)
        except (ValueError, sqlite3.Error) as exc:
            self._error(exc)
            return
        self.refresh_all()
        self.status.set('Zwolniono miejsce. Historia pozostała nienaruszona.')

    def clear_selected_slot(self):
        self._clear()

    def clear_oldest(self):
        self._clear(oldest=True)

    def refresh_history(self):
        all_drivers = self.store.drivers(include_inactive=True)
        self.history_map = {f"{d['name']} [#{d['id']}]": d['id'] for d in all_drivers}
        values = ['Wszyscy kierowcy', *self.history_map]
        self.history_combo['values'] = values
        if self.history_filter.get() not in values:
            self.history_filter.set(values[0])
        driver_id = self.history_map.get(self.history_filter.get())
        self.history_tree.delete(*self.history_tree.get_children())
        for row in self.store.history(driver_id, self.history_search.get()):
            self.history_tree.insert('', 'end', iid=str(row['id']), values=(row['id'], row['driver_name'], row['truck'], row['trailer'], row['info'],
                                     display_date(row['load_date']), display_date(row['unload_date']), row['recorded_at'].replace('T', ' ').replace('+00:00', ''), row['route'], row['order_number']))
        self.refresh_deadline_colors()

    def refresh_registers(self):
        for tree in self.register_trees.values():
            tree.delete(*tree.get_children())
        for driver in self.store.drivers(include_inactive=True):
            self.register_trees['driver'].insert('', 'end', iid=str(driver['id']),
                values=(driver['id'], driver['name'], driver['truck'], 'Tak' if driver['active'] else 'Nie', driver['phone']))
        for kind in ('truck', 'trailer'):
            for item in self.store.vehicles(kind, include_inactive=True):
                self.register_trees[kind].insert('', 'end', iid=str(item['id']), values=(item['id'], item['plate'], 'Tak' if item['active'] else 'Nie'))

    def form_dialog(self, title, definitions, callback):
        """definitions: (key, label, initial, choices_or_None)."""
        window = tk.Toplevel(self.root)
        window.title(title)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=22)
        frame.pack(fill='both', expand=True)
        variables = {}
        for index, (key, label, initial, choices) in enumerate(definitions):
            ttk.Label(frame, text=label).grid(row=index, column=0, sticky='w', pady=8, padx=(0, 14))
            variable = tk.StringVar(value=initial)
            variables[key] = variable
            control = ttk.Entry(frame, textvariable=variable, width=40) if choices is None else ttk.Combobox(frame, textvariable=variable, values=choices, state='readonly', width=38)
            control.grid(row=index, column=1, sticky='ew', pady=8)
            if index == 0:
                control.focus_set()
        def submit():
            try:
                callback({k: v.get() for k, v in variables.items()})
            except (ValueError, sqlite3.Error, OSError) as exc:
                messagebox.showerror(title, str(exc), parent=window)
                return
            window.destroy()
            self.refresh_all()
        ttk.Button(frame, text='Zapisz', style='Accent.TButton', command=submit).grid(row=len(definitions), column=1, sticky='e', pady=(16, 0))
        window.bind('<Return>', lambda event: submit())
        window.bind('<Escape>', lambda event: window.destroy())
        window.grab_set()
        return window

    def add_vehicle_dialog(self, kind):
        if not self._guard():
            return
        self.form_dialog('Dodaj tira' if kind == 'truck' else 'Dodaj przyczepę', [('plate', 'Numer rejestracyjny', '', None)],
                         lambda values: self.store.add_vehicle(kind, values['plate']))

    def add_driver_dialog(self):
        if not self._guard():
            return
        trucks = {v['plate']: v['id'] for v in self.store.vehicles('truck')}
        if not trucks:
            messagebox.showinfo('Najpierw dodaj tira', 'Dodaj przynajmniej jednego tira w zakładce Rejestry.', parent=self.root)
            self.notebook.select(2)
            return
        def add(values):
            driver_id = self.store.add_driver(values['name'], trucks[values['truck']], values['phone'])
            self.select_driver(driver_id)
            self.status.set('Dodano kierowcę. Kolejne miejsca dodasz przyciskiem „+ Dodaj wpis”.')
        self.form_dialog('Nowy kierowca', [('name', 'Imię i nazwisko', '', None), ('truck', 'TIR', next(iter(trucks)), list(trucks)),
                                         ('phone', 'Telefon (opcjonalnie)', '', None)], add)

    def edit_driver_dialog(self):
        selection = self.register_trees['driver'].selection()
        if not selection or not self._guard():
            return
        driver_id = int(selection[0])
        driver = self.store.driver(driver_id)
        trucks = {v['plate']: v['id'] for v in self.store.vehicles('truck')}
        def edit(values):
            name = values['name'].strip()
            if not name or values['truck'] not in trucks:
                raise ValueError('Podaj nazwę i wybierz tira.')
            # Jedna transakcja obejmuje nazwę, tira i nowe kopie historii.
            with self.store.conn:
                if len(values['phone'].strip()) > 100:
                    raise ValueError('Telefon: maksymalnie 100 znaków.')
                self.store.conn.execute('UPDATE drivers SET name=?,truck_id=?,phone=? WHERE id=?', (name, trucks[values['truck']], values['phone'].strip(), driver_id))
                for slot in self.store.slots(driver_id):
                    self.store._archive(slot)
        self.form_dialog('Zmień kierowcę / przypisanie tira', [('name', 'Imię i nazwisko', driver['name'], None),
                         ('truck', 'TIR', driver['truck'], list(trucks)), ('phone', 'Telefon (opcjonalnie)', driver['phone'], None)], edit)

    def toggle_active(self, kind):
        selection = self.register_trees[kind].selection()
        if not selection or not self._guard():
            return
        item_id = int(selection[0])
        try:
            if kind == 'driver':
                item = self.store.driver(item_id)
                if item['active'] and not messagebox.askyesno('Wyłączyć kierowcę?', 'Kierowca zniknie z widoku bieżącego. Wpisy i historia pozostaną. Możesz go ponownie włączyć.', parent=self.root):
                    return
                self.store.set_driver_active(item_id, not item['active'])
                if self.driver_id == item_id and item['active']:
                    self.driver_id = None
                    self.refresh_slots()
            else:
                item = next(v for v in self.store.vehicles(kind, True) if v['id'] == item_id)
                self.store.set_vehicle_active(item_id, not item['active'])
            self.refresh_all()
        except (ValueError, sqlite3.Error) as exc:
            self._error(exc)

    def restore_dialog(self):
        selected = self.history_tree.selection()
        if not selected or not self._guard():
            return
        item = next(h for h in self.store.history() if h['id'] == int(selected[0]))
        driver = self.store.driver(item['driver_id'])
        if not driver['active']:
            messagebox.showinfo('Kierowca jest wyłączony', 'Włącz kierowcę w Rejestrach, aby przywrócić wpis.', parent=self.root)
            return
        empty = [str(s['position']) for s in self.store.slots(driver['id']) if s['status'] == 'Puste']
        empty.append('Nowe miejsce')
        if driver['truck'] != item['truck'] and not messagebox.askyesno('Przywrócenie z obecnym tirem',
            f"Historyczny TIR: {item['truck']}\nObecny TIR: {driver['truck']}\nPrzywrócić dane z obecnym tirem? Stara kopia pozostanie.", parent=self.root):
            return
        def restore(values):
            position = None if values['position'] == 'Nowe miejsce' else int(values['position'])
            position = self.store.restore(item['id'], position)
            self.select_driver(driver['id'])
            self.select_slot(position)
            self.notebook.select(0)
            self.status.set('Przywrócono wpis do pustego miejsca.')
        self.form_dialog('Przywróć wpis', [('position', f"Miejsce — {driver['name']}", empty[0], empty)], restore)

    def export_dialog(self):
        if not self._guard():
            return
        drivers = {'Wszyscy kierowcy': None, **{f"{d['name']} [#{d['id']}]": d['id'] for d in self.store.drivers(True)}}
        scopes = {'Bieżące + pełna historia': 'both', 'Tylko bieżące': 'current', 'Tylko historia': 'history'}
        def export(values):
            target = filedialog.asksaveasfilename(parent=self.root, title='Zapisz eksport', defaultextension='.xlsx',
                filetypes=[('Excel', '*.xlsx')], initialfile='Floteka-eksport.xlsx')
            if not target:
                return
            export_xlsx(self.store, target, scopes[values['scope']], drivers[values['driver']])
            self.status.set(f'Zapisano eksport: {target}')
        self.form_dialog('Eksport do Excela', [('scope', 'Zakres danych', next(iter(scopes)), list(scopes)),
                         ('driver', 'Kierowca', 'Wszyscy kierowcy', list(drivers))], export)

    def backup(self):
        if not self._guard():
            return
        target = filedialog.asksaveasfilename(parent=self.root, title='Kopia zapasowa bazy', defaultextension='.sqlite3',
            filetypes=[('Baza SQLite', '*.sqlite3')], initialfile=f'Floteka-kopia-{date.today().isoformat()}.sqlite3')
        if target:
            try:
                self.store.backup(target)
                self.status.set(f'Kopia zapasowa: {target}')
            except (OSError, sqlite3.Error, ValueError) as exc:
                self._error(exc)

    def close(self):
        if self._guard():
            self.root.destroy()
