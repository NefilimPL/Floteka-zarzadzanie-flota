"""Jasny i ciemny motyw ttk z czytelnymi polami formularza."""

from tkinter import ttk


def apply_theme(root, dark=False):
    palette = ({'bg': '#151b26', 'surface': '#202b3b', 'text': '#eef4fc', 'muted': '#afc1d5',
                'border': '#46566e', 'button': '#30425a', 'selected': '#275d96'} if dark else
               {'bg': '#eff3f5', 'surface': '#ffffff', 'text': '#183a4d', 'muted': '#536d7e',
                'border': '#9aabb8', 'button': '#e0e9ef', 'selected': '#275d96'})
    style = ttk.Style(root)
    style.theme_use('clam')
    root.configure(background=palette['bg'])
    style.configure('.', font=('Segoe UI', 10), background=palette['bg'], foreground=palette['text'])
    for name in ('TFrame', 'TLabel', 'TLabelframe', 'TNotebook'):
        style.configure(name, background=palette['bg'], foreground=palette['text'], bordercolor=palette['border'])
    style.configure('TLabelframe.Label', background=palette['bg'], foreground=palette['text'])
    style.configure('TButton', padding=(12, 7), background=palette['button'], foreground=palette['text'],
                    bordercolor=palette['border'], lightcolor=palette['button'], darkcolor=palette['button'])
    style.map('TButton', background=[('active', palette['selected'])], foreground=[('active', '#ffffff')])
    style.configure('Accent.TButton', background='#087f87', foreground='#ffffff', font=('Segoe UI', 10, 'bold'))
    style.configure('Compact.TButton', padding=(10, 4))
    style.configure('Compact.Accent.TButton', padding=(10, 4))
    style.map('Accent.TButton', background=[('active', '#08656c')], foreground=[('active', '#ffffff')])
    style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'))
    style.configure('Subtitle.TLabel', foreground=palette['muted'])
    style.configure('Heading.TLabel', font=('Segoe UI', 13, 'bold'))
    for name in ('TEntry', 'TCombobox'):
        style.configure(name, fieldbackground=palette['surface'], background=palette['button'],
                        foreground=palette['text'], insertcolor=palette['text'], bordercolor=palette['border'],
                        selectbackground=palette['selected'], selectforeground='#ffffff')
        style.map(name, fieldbackground=[('readonly', palette['surface'])], foreground=[('readonly', palette['text'])])
    style.configure('TCheckbutton', background=palette['bg'], foreground=palette['text'])
    style.map('TCheckbutton', background=[('active', palette['bg'])], foreground=[('active', palette['text'])])
    style.configure('Treeview', rowheight=31, background=palette['surface'], fieldbackground=palette['surface'],
                    foreground=palette['text'], bordercolor=palette['border'])
    style.map('Treeview', background=[('selected', palette['selected'])], foreground=[('selected', '#ffffff')])
    style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), padding=(6, 5),
                    background=palette['button'], foreground=palette['text'], bordercolor=palette['border'])
    style.map('Treeview.Heading', background=[('active', palette['selected'])], foreground=[('active', '#ffffff')])
    style.configure('TNotebook.Tab', padding=(14, 6), background=palette['button'], foreground=palette['text'])
    style.map('TNotebook.Tab', background=[('selected', palette['surface'])], foreground=[('selected', palette['text'])])
    style.configure('TScrollbar', background=palette['button'], troughcolor=palette['bg'],
                    arrowcolor=palette['text'], bordercolor=palette['border'])
    for key, value in (('Background', palette['surface']), ('Foreground', palette['text']),
                       ('selectBackground', palette['selected']), ('selectForeground', '#ffffff')):
        root.option_add('*TCombobox*Listbox*' + key, value)
    return palette
