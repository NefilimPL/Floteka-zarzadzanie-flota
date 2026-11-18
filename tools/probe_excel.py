import win32com.client

app = None
book = None
try:
    app = win32com.client.DispatchEx('Excel.Application')
    app.Visible = False
    book = app.Workbooks.Add()
    print('Excel:', app.Version, flush=True)
    try:
        print('VB components:', book.VBProject.VBComponents.Count, flush=True)
    except Exception as exc:
        print('VBProject blocked:', str(exc), flush=True)
finally:
    if book is not None:
        book.Close(False)
    if app is not None:
        app.Quit()
