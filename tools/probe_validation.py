import pythoncom
import win32com.client

app = win32com.client.DispatchEx('Excel.Application')
app.Visible = False
app.DisplayAlerts = False
book = app.Workbooks.Add()
try:
    first = book.Worksheets(1)
    first.Name = 'TIRY'
    first.Range('A6').Value = 'WX1'
    other = book.Worksheets.Add()
    other.Name = 'KIEROWCY'
    print('International:', app.International, flush=True)
    name = book.Names.Add('ListaTirow', "='TIRY'!$A$6:$A$6")
    print('Name:', name.Name, 'RefersTo:', name.RefersTo, flush=True)
    print('Protected:', other.ProtectContents, 'Active:', app.ActiveSheet.Name, flush=True)
    for expression in ('=ListaTirow', 'A,B', 'A;B', '=$A$1:$A$2'):
        try:
            other.Range('C6').Validation.Delete()
            other.Range('C6').Validation.Add(3, 1, 1, expression, pythoncom.Missing)
            print(expression, 'OK positional', flush=True)
        except Exception as exc:
            print(expression, 'FAILED', str(exc), flush=True)
finally:
    book.Close(False)
    app.Quit()
