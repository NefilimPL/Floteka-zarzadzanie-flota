"""Diagnostyka okien wyłącznie wskazanego procesu budowy Excela."""
import sys
import win32gui
import win32process
import win32con

pid = int(sys.argv[1])

def visit(hwnd, _):
    if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
        return
    title, cls = win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd)
    if '--all' in sys.argv:
        print('WINDOW', hwnd, cls, title, 'visible', win32gui.IsWindowVisible(hwnd), flush=True)
    if cls in ('#32770', 'NUIDialog'):
        print('DIALOG', title, flush=True)
        win32gui.EnumChildWindows(hwnd, lambda child, __: print(win32gui.GetDlgCtrlID(child), win32gui.GetClassName(child), win32gui.GetWindowText(child), flush=True), None)
        if '--dismiss' in sys.argv and title == 'Floteka':
            def ok_button(child, __):
                if win32gui.GetWindowText(child) == 'OK':
                    win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
            win32gui.EnumChildWindows(hwnd, ok_button, None)
        if '--end' in sys.argv and title == 'Microsoft Visual Basic':
            def end_button(child, __):
                if win32gui.GetWindowText(child) == '&End':
                    win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
            win32gui.EnumChildWindows(hwnd, end_button, None)

win32gui.EnumWindows(visit, None)
