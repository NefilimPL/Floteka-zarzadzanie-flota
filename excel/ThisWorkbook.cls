Option Explicit

Private Sub Workbook_Open()
    Flota.Initialize
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Flota.HandleChange Sh, Target
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not Flota.FlushAll() Then Cancel = True
End Sub
