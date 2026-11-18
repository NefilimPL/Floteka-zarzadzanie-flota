Option Explicit

Public Const FIRST_ROW As Long = 6
Private Const PASS As String = "flota"
Private Const REGISTRY_LAST_ROW As Long = 5005
Private keys As Object
Private vehicleKeys As Object
Private registryFormulas As Object
Private busy As Boolean
Public TestMode As Boolean

Private Function Ws(ByVal sheetName As String) As Worksheet
    Set Ws = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function LastRow(ByVal sheetName As String, Optional ByVal col As Long = 1) As Long
    Dim sh As Worksheet, upper As Long, row As Long, values As Variant
    Set sh = Ws(sheetName)
    ' Nawigacja End/Find potrafi pomijać odfiltrowane komórki. Odczyt Value2
    ' obejmuje fizyczne komórki, również ukryte i odfiltrowane.
    upper = sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    If upper > 1 Then
        values = sh.Range(sh.Cells(1, col), sh.Cells(upper, col)).Value2
        For row = upper To 1 Step -1
            If IsError(values(row, 1)) Then
                LastRow = row: Exit For
            ElseIf Len(CStr(values(row, 1))) > 0 Then
                LastRow = row: Exit For
            End If
        Next row
    ElseIf Len(CStr(sh.Cells(1, col).Value2)) > 0 Then
        LastRow = 1
    End If
    If LastRow < FIRST_ROW - 1 Then LastRow = FIRST_ROW - 1
End Function

Private Function RegPlate(ByVal value As Variant) As String
    RegPlate = UCase$(Replace(Replace(Trim$(CStr(value)), " ", ""), vbTab, ""))
End Function

Private Function TextValue(ByVal value As Variant) As String
    If IsError(value) Then Err.Raise vbObjectError + 1, , "Komórka zawiera błąd."
    TextValue = Trim$(CStr(value))
End Function

Public Sub Initialize(Optional ByVal captureRegistry As Boolean = True)
    Dim sh As Worksheet, last As Long
    On Error GoTo Failed
    Application.EnableEvents = False
    Ws("_STAN").Unprotect PASS
    If IsEmpty(Ws("_STAN").Range("B3").Value2) Then Ws("_STAN").Range("B3").Value2 = True
    Ws("_STAN").Range("A3").Value2 = "Kolory terminów"
    ThisWorkbook.Names.Add Name:="KoloryTerminow", RefersTo:="='_STAN'!$B$3"
    If captureRegistry Then
        Set registryFormulas = CreateObject("Scripting.Dictionary")
        CaptureRegistryDetails Ws("KIEROWCY"), 5, 5, 7
        CaptureRegistryDetails Ws("TIRY"), 2, 6, 11
    End If
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = "DANE" Or sh.Name = "ARCHIWUM" Then
            sh.Unprotect PASS
            last = LastRow(sh.Name)
            UpdateDeadlineRules sh, last
            If sh.Name = "ARCHIWUM" Then
                If Not sh.AutoFilterMode Then
                    RefreshArchiveFilter sh, last
                ElseIf sh.AutoFilter.Range.Columns.Count <> 12 Then
                    RefreshArchiveFilter sh, last
                End If
            End If
        End If
        sh.Protect Password:=PASS, UserInterfaceOnly:=True, AllowFiltering:=True
        sh.EnableSelection = xlNoRestrictions
    Next sh
    ThisWorkbook.Protect Password:=PASS, Structure:=True
    Set keys = Nothing
    Set vehicleKeys = Nothing
    LoadKeys
    UpdateLists
    UpdateDeadlineButtons
Done:
    Application.EnableEvents = True
    Exit Sub
Failed:
    Application.EnableEvents = True
    NotifyError "Nie udało się przygotować arkuszy: " & Err.Description
End Sub

Private Sub UpdateDeadlineRules(ByVal sh As Worksheet, ByVal last As Long)
    Dim area As Range, rule As FormatCondition
    Set area = sh.Range("A" & FIRST_ROW & ":L" & Application.Max(FIRST_ROW, last))
    ' Odtwarzamy dwie reguły dla całego obszaru. ModifyAppliesToRange na
    ' zachodzących zakresach potrafi destabilizować lokalny Excel 16.
    area.FormatConditions.Delete
    Set rule = area.FormatConditions.Add(Type:=xlExpression, Formula1:=LocalDeadlineFormula("=AND(KoloryTerminow,ISNUMBER($F6),$F6<TODAY())"))
    rule.Interior.Color = RGB(229, 57, 53)
    rule.Font.Color = RGB(255, 255, 255)
    rule.StopIfTrue = True
    Set rule = area.FormatConditions.Add(Type:=xlExpression, Formula1:=LocalDeadlineFormula("=AND(KoloryTerminow,ISNUMBER($F6),$F6>=TODAY(),$F6<=TODAY()+1)"))
    rule.Interior.Color = RGB(255, 167, 38)
    rule.Font.Color = RGB(33, 19, 0)
    rule.StopIfTrue = True
End Sub

Private Function LocalDeadlineFormula(ByVal expression As String) As String
    ' FormatConditions.Add oczekuje lokalnej składni, np. ORAZ i średników.
    With Ws("_STAN").Range("C3")
        .Formula = expression
        LocalDeadlineFormula = .FormulaLocal
        .ClearContents
    End With
End Function

Private Sub UpdateDeadlineButtons()
    Dim sheetName As Variant
    For Each sheetName In Array("START", "DANE")
        On Error Resume Next
        Ws(CStr(sheetName)).Shapes("toolbar_PrzelaczKoloryTerminow").TextFrame.Characters.Text = _
            "Kolory terminów: " & IIf(CBool(Ws("_STAN").Range("B3").Value2), "WŁ.", "WYŁ.")
        On Error GoTo 0
    Next sheetName
End Sub

Public Sub PrzelaczKoloryTerminow()
    Ws("_STAN").Range("B3").Value2 = Not CBool(Ws("_STAN").Range("B3").Value2)
    UpdateDeadlineButtons
    Application.Calculate
End Sub

Private Sub LoadKeys()
    Dim row As Long, key As String, last As Long, values As Variant
    If Not keys Is Nothing Then Exit Sub
    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbBinaryCompare
    last = LastRow("ARCHIWUM")
    If last < FIRST_ROW Then Exit Sub
    ' Odtwarzamy indeks z rzeczywistych danych, także gdy stara wersja
    ' pozostawiła pusty lub nieaktualny techniczny klucz w kolumnie J.
    values = Ws("ARCHIWUM").Range("A" & FIRST_ROW & ":L" & last).Value2
    For row = 1 To UBound(values, 1)
        If Len(CStr(values(row, 1))) > 0 And Len(CStr(values(row, 2))) > 0 And Len(CStr(values(row, 3))) > 0 _
            And Len(CStr(values(row, 5))) > 0 And Len(CStr(values(row, 6))) > 0 And Val(values(row, 8)) > 0 Then
            key = MakeHistoryKey(values(row, 8), values, row)
            keys(key) = True
        End If
    Next row
End Sub

Public Function HistoryCount() As Long
    HistoryCount = LastRow("ARCHIWUM") - FIRST_ROW + 1
End Function

Public Function DriverFirstRow(ByVal driverId As Long) As Long
    Dim row As Long
    For row = FIRST_ROW To LastRow("DANE") Step 4
        If Val(Ws("DANE").Cells(row, 9).Value2) = driverId Then
            DriverFirstRow = row
            Exit Function
        End If
    Next row
End Function

Private Function RegisterRow(ByVal driverId As Long) As Long
    Dim row As Long
    For row = FIRST_ROW To LastRow("KIEROWCY")
        If Val(Ws("KIEROWCY").Cells(row, 1).Value2) = driverId Then RegisterRow = row: Exit Function
    Next row
End Function

Private Sub AddPlate(ByVal sheetName As String, ByVal number As String)
    Dim row As Long, last As Long, values As Variant, known As Object
    number = RegPlate(number)
    If number = "" Then Exit Sub
    If vehicleKeys Is Nothing Then Set vehicleKeys = CreateObject("Scripting.Dictionary")
    If Not vehicleKeys.Exists(sheetName) Then
        Set known = CreateObject("Scripting.Dictionary")
        last = LastRow(sheetName)
        If last >= FIRST_ROW Then
            values = Ws(sheetName).Range("A" & FIRST_ROW & ":A" & last).Value2
            If last = FIRST_ROW Then
                known(RegPlate(values)) = True
            Else
                For row = 1 To UBound(values, 1): known(RegPlate(values(row, 1))) = True: Next row
            End If
        End If
        vehicleKeys.Add sheetName, known
    End If
    Set known = vehicleKeys(sheetName)
    If known.Exists(number) Then Exit Sub
    row = LastRow(sheetName) + 1
    If sheetName = "TIRY" Then
        ' Przy ręcznym wpisywaniu pola opcjonalne mogą pojawić się przed rejestracją.
        row = Application.Max(row, LastRow(sheetName, 2) + 1, LastRow(sheetName, 3) + 1, _
            LastRow(sheetName, 4) + 1, LastRow(sheetName, 5) + 1, LastRow(sheetName, 6) + 1)
        If row > REGISTRY_LAST_ROW Then
            Ws(sheetName).Cells(row, 2).Resize(1, 5).Locked = False
            Ws(sheetName).Cells(row, 2).NumberFormat = "0"
            Ws(sheetName).Cells(row, 3).Resize(1, 3).NumberFormat = "dd.mm.yyyy"
            Ws(sheetName).Cells(row, 6).NumberFormat = "@"
        End If
    End If
    Ws(sheetName).Cells(row, 1).NumberFormat = "@"
    Ws(sheetName).Cells(row, 1).Value2 = number
    known(number) = True
    UpdateLists
End Sub

Private Sub UpdateLists()
    ThisWorkbook.Names("ListaTirow").RefersTo = "='TIRY'!$A$6:$A$" & Application.Max(FIRST_ROW, LastRow("TIRY"))
    ThisWorkbook.Names("ListaPrzyczep").RefersTo = "='PRZYCZEPY'!$A$6:$A$" & Application.Max(FIRST_ROW, LastRow("PRZYCZEPY"))
End Sub

Public Function AddDriverRecord(ByVal driverName As String, ByVal truck As String) As Long
    Dim row As Long, driverId As Long, oldEvents As Boolean, description As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    driverName = Trim$(driverName): truck = RegPlate(truck)
    If driverName = "" Or truck = "" Then Err.Raise vbObjectError + 2, , "Podaj kierowcę i tira."
    If Len(driverName) > 150 Or Len(truck) > 40 Then Err.Raise vbObjectError + 3, , "Zbyt długa nazwa lub rejestracja."
    driverId = CLng(Val(Ws("_STAN").Range("B1").Value2)) + 1
    row = Application.Max(LastRow("KIEROWCY"), LastRow("KIEROWCY", 2), LastRow("KIEROWCY", 3), LastRow("KIEROWCY", 5)) + 1
    Ws("KIEROWCY").Cells(row, 1).Value2 = driverId
    Ws("KIEROWCY").Cells(row, 2).NumberFormat = "@"
    Ws("KIEROWCY").Cells(row, 2).Value2 = driverName
    Ws("KIEROWCY").Cells(row, 3).NumberFormat = "@"
    Ws("KIEROWCY").Cells(row, 3).Value2 = truck
    Ws("KIEROWCY").Cells(row, 4).Value2 = "TAK"
    Ws("KIEROWCY").Cells(row, 5).NumberFormat = "@"
    Ws("KIEROWCY").Cells(row, 5).Locked = False
    Ws("_STAN").Range("B1").Value2 = driverId
    AddPlate "TIRY", truck
    SetupDriverBlock driverId, driverName, truck
    AddDriverRecord = driverId
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    description = Err.Description
    On Error Resume Next
    If row >= FIRST_ROW Then Ws("KIEROWCY").Cells(row, 1).Resize(1, 4).ClearContents
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 32, , description
End Function

Private Sub SetupDriverBlock(ByVal driverId As Long, ByVal driverName As String, ByVal truck As String)
    Dim startRow As Long, row As Long, block As Range, btn As Shape, background As Long
    Dim stage As String, description As String
    On Error GoTo Failed
    stage = "formatowanie bloku"
    Ws("DANE").Unprotect PASS
    startRow = LastRow("DANE") + 1
    If startRow > 1048572 Then Err.Raise vbObjectError + 4, , "Brak miejsca na nowy blok kierowcy."
    Set block = Ws("DANE").Cells(startRow, 1).Resize(4, 12)
    background = RGB(237, 246, 247)
    If ((startRow - FIRST_ROW) \ 4) Mod 2 = 1 Then background = RGB(248, 250, 252)
    block.Interior.Color = background
    block.RowHeight = 38
    block.Font.Name = "Calibri": block.Font.Size = 11
    block.VerticalAlignment = xlCenter
    block.Borders(xlInsideHorizontal).Color = RGB(218, 229, 235)
    block.Borders(xlEdgeBottom).Color = RGB(153, 181, 193)
    For row = startRow To startRow + 3
        With Ws("DANE")
            stage = "wpisanie nazwy i tira"
            .Cells(row, 1).NumberFormat = "@": .Cells(row, 1).Value2 = driverName
            .Cells(row, 2).NumberFormat = "@": .Cells(row, 2).Value2 = truck
            .Cells(row, 3).NumberFormat = "@": .Cells(row, 4).NumberFormat = "@"
            .Cells(row, 5).Resize(1, 2).NumberFormat = "dd.mm.yyyy"
            .Cells(row, 3).Resize(1, 4).Locked = False
            .Cells(row, 4).WrapText = True
            With .Cells(row, 11).Resize(1, 2)
                .Locked = False
                .NumberFormat = "@"
                .WrapText = True
            End With
            .Cells(row, 9).Value2 = driverId
            .Cells(row, 10).Value2 = row - startRow + 1
            .Cells(row, 7).Value2 = "Puste"
            stage = "lista wyboru przyczepy"
            With .Cells(row, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:="=ListaPrzyczep"
                .ShowError = False: .IgnoreBlank = True
                .InputTitle = "Numer przyczepy"
                .InputMessage = "Wybierz z listy lub wpisz nowy numer."
            End With
            stage = "przycisk czyszczenia miejsca"
            Set btn = .Shapes.AddShape(5, .Cells(row, 8).Left + 8, .Cells(row, 8).Top + 7, 50, 24)
            btn.Name = "clear_" & row
            btn.TextFrame.Characters.Text = "x"
            btn.TextFrame.Characters.Font.Color = RGB(126, 53, 53)
            btn.TextFrame.HorizontalAlignment = xlHAlignCenter
            btn.TextFrame.VerticalAlignment = xlVAlignCenter
            btn.Fill.ForeColor.RGB = RGB(249, 231, 231)
            btn.Line.Visible = False
            btn.OnAction = "WyczyscMiejsce"
            btn.Placement = xlMoveAndSize
            btn.AlternativeText = "Wyczyść miejsce " & (row - startRow + 1) & "; historia pozostaje"
        End With
        Ws("_STAN").Cells(row, 1).Value2 = driverId
        Ws("_STAN").Cells(row, 2).Value2 = row - startRow + 1
        Ws("_STAN").Cells(row, 8).Value2 = driverName
        Ws("_STAN").Cells(row, 9).Value2 = truck
    Next row
    stage = "obszar wydruku"
    UpdateDeadlineRules Ws("DANE"), startRow + 3
    Ws("DANE").PageSetup.PrintArea = "$A$5:$L$" & startRow + 3
    Ws("DANE").Protect Password:=PASS, UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub
Failed:
    description = Err.Description
    ' Nowy blok jest dopisywany na końcu. Cofamy wyłącznie ten blok,
    ' aby kolejny kierowca nadal zaczynał się na granicy czterech miejsc.
    On Error Resume Next
    Ws("DANE").Unprotect PASS
    If startRow >= FIRST_ROW Then
        For row = startRow To startRow + 3
            Ws("DANE").Shapes("clear_" & row).Delete
        Next row
        Ws("DANE").Cells(startRow, 1).Resize(4, 12).Clear
        Ws("_STAN").Cells(startRow, 1).Resize(4, 11).ClearContents
    End If
    Ws("DANE").Protect Password:=PASS, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    Err.Raise vbObjectError + 30, , stage & ": " & description
End Sub

Private Function ParsedDate(ByVal value As Variant) As Variant
    Dim pieces() As String, text As String, result As Date, day As Long, month As Long, year As Long
    Dim serial As Double, firstSerial As Long, lastSerial As Long
    text = Trim$(CStr(value))
    If text = "" Then ParsedDate = Empty: Exit Function
    On Error GoTo Invalid
    firstSerial = 1: lastSerial = 2958465
    If ThisWorkbook.Date1904 Then firstSerial = 0: lastSerial = 2957003
    If IsNumeric(value) Then
        serial = CDbl(value)
        If serial < firstSerial Or serial > lastSerial Or serial <> Fix(serial) Then GoTo Invalid
        If Not ThisWorkbook.Date1904 And serial = 60 Then GoTo Invalid
        ParsedDate = serial: Exit Function
    End If
    text = Replace(text, "/", ".")
    If InStr(text, "-") > 0 Then
        pieces = Split(text, "-")
        If UBound(pieces) <> 2 Then GoTo Invalid
        If Len(pieces(0)) <> 4 Then GoTo Invalid
        If text Like "*[!0-9-]*" Then GoTo Invalid
        year = CLng(pieces(0)): month = CLng(pieces(1)): day = CLng(pieces(2))
    Else
        pieces = Split(text, ".")
        If UBound(pieces) <> 2 Then GoTo Invalid
        If Len(pieces(2)) <> 4 Then GoTo Invalid
        If text Like "*[!0-9.]*" Then GoTo Invalid
        day = CLng(pieces(0)): month = CLng(pieces(1)): year = CLng(pieces(2))
    End If
    If year < 1900 Or year > 9999 Then GoTo Invalid
    result = DateSerial(year, month, day)
    If VBA.Day(result) <> day Or VBA.Month(result) <> month Or VBA.Year(result) <> year Then GoTo Invalid
    serial = CDbl(result)
    If ThisWorkbook.Date1904 Then
        serial = serial - 1462
    ElseIf result < DateSerial(1900, 3, 1) Then
        serial = serial - 1
    End If
    If serial < firstSerial Then GoTo Invalid
    ParsedDate = serial: Exit Function
Invalid:
    Err.Raise vbObjectError + 5, , "Nieprawidłowa data. Użyj DD.MM.RRRR, np. 17.11.2026."
End Function

Private Function InputValue(ByVal cell As Range) As Variant
    If cell.HasFormula Then Err.Raise vbObjectError + 7, , "Wpisz dane, nie formułę."
    If IsError(cell.Value2) Then Err.Raise vbObjectError + 8, , "Komórka zawiera błąd."
    InputValue = cell.Value2
End Function

Private Function OptionalText(ByVal value As Variant, ByVal maxLength As Long, ByVal label As String) As Variant
    Dim text As String
    text = TextValue(value)
    If Len(text) > maxLength Then Err.Raise vbObjectError + 34, , label & ": maksymalnie " & maxLength & " znaków."
    If Len(text) > 0 Then OptionalText = text
End Function

Private Function ParsedMileage(ByVal value As Variant) As Variant
    Dim number As Double
    If Len(TextValue(value)) = 0 Then Exit Function
    On Error GoTo Invalid
    If Not IsNumeric(value) Then GoTo Invalid
    number = CDbl(value)
    If number < 0 Or number <> Fix(number) Then GoTo Invalid
    ParsedMileage = number
    Exit Function
Invalid:
    Err.Raise vbObjectError + 35, , "Podaj przebieg w kilometrach jako nieujemną liczbę całkowitą."
End Function

Private Function RowState(ByVal row As Long) As String
    Dim sh As Worksheet, count As Long, col As Long
    Set sh = Ws("DANE")
    For col = 3 To 6
        If Len(CStr(sh.Cells(row, col).Value2)) > 0 Then count = count + 1
    Next col
    For col = 11 To 12
        If Len(CStr(sh.Cells(row, col).Value2)) > 0 Then count = count + 1
    Next col
    If count = 0 Then RowState = "Puste": Exit Function
    If sh.Cells(row, 3).Value2 <> "" And sh.Cells(row, 5).Value2 <> "" And sh.Cells(row, 6).Value2 <> "" Then
        RowState = "Kompletny"
    Else
        RowState = "Szkic"
    End If
End Function

Private Sub ArchiveRow(ByVal row As Long)
    Dim col As Long, key As String, destination As Long, sh As Worksheet, arc As Worksheet, values As Variant
    Dim record(1 To 1, 1 To 12) As Variant, details As Range, description As String
    If RowState(row) <> "Kompletny" Then Exit Sub
    LoadKeys
    Set sh = Ws("DANE"): Set arc = Ws("ARCHIWUM")
    values = sh.Cells(row, 1).Resize(1, 12).Value2
    key = MakeHistoryKey(sh.Cells(row, 9).Value2, values, 1)
    If keys.Exists(key) Then Exit Sub
    On Error GoTo Failed
    arc.Unprotect PASS
    destination = LastRow("ARCHIWUM") + 1
    If destination > arc.Rows.Count Then Err.Raise vbObjectError + 6, , "Archiwum jest pełne. Nie można zwolnić miejsca."
    ' Pojedynczy zapis całego rekordu ogranicza ryzyko częściowego wiersza.
    For col = 1 To 6: record(1, col) = values(1, col): Next col
    record(1, 7) = CDbl(Now): record(1, 8) = sh.Cells(row, 9).Value2
    record(1, 9) = destination - FIRST_ROW + 1: record(1, 10) = key
    record(1, 11) = values(1, 11): record(1, 12) = values(1, 12)
    arc.Cells(destination, 1).Resize(1, 4).NumberFormat = "@"
    arc.Cells(destination, 5).Resize(1, 2).NumberFormat = "dd.mm.yyyy"
    arc.Cells(destination, 7).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    arc.Cells(destination, 10).Resize(1, 3).NumberFormat = "@"
    arc.Cells(destination, 1).Resize(1, 12).Value2 = record
    ' Pod aktywnym filtrem Excel potrafi pominąć sąsiednią ukrytą kolumnę
    ' przy zapisie tablicy. Zapisujemy indeks jawnie i sprawdzamy wynik.
    arc.Cells(destination, 10).Value2 = key
    Set details = arc.Cells(destination, 11).Resize(1, 2)
    details.Value2 = sh.Cells(row, 11).Resize(1, 2).Value2
    If CStr(arc.Cells(destination, 10).Value2) <> key Or CStr(arc.Cells(destination, 11).Value2) <> CStr(record(1, 11)) _
        Or CStr(arc.Cells(destination, 12).Value2) <> CStr(record(1, 12)) Then
        Err.Raise vbObjectError + 33, , "Excel nie zapisał pełnego wpisu historii. Miejsce nie zostało wyczyszczone."
    End If
    keys(key) = True
    On Error Resume Next
    arc.Cells(destination, 1).Resize(1, 12).Interior.Color = RGB(240, 247, 249)
    arc.Cells(destination, 4).WrapText = True
    details.WrapText = True
    arc.Rows(destination).RowHeight = 32
    UpdateDeadlineRules arc, destination
    RefreshArchiveFilter arc, destination
    arc.Protect Password:=PASS, UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub
Failed:
    description = Err.Description
    On Error Resume Next
    If destination >= FIRST_ROW And destination <= arc.Rows.Count Then arc.Cells(destination, 1).Resize(1, 12).ClearContents
    arc.Protect Password:=PASS, UserInterfaceOnly:=True, AllowFiltering:=True
    On Error GoTo 0
    Err.Raise vbObjectError + 31, , "Nie zapisano historii: " & description
End Sub

Private Function MakeHistoryKey(ByVal driverId As Variant, ByRef values As Variant, ByVal row As Long) As String
    Dim col As Long, part As String, result As String
    result = CStr(driverId) & "|"
    For col = 1 To 6
        part = CStr(values(row, col))
        result = result & Len(part) & ":" & part
    Next col
    ' Pusta para opcjonalna zachowuje dokładnie dawny klucz.
    If Len(CStr(values(row, 11))) + Len(CStr(values(row, 12))) > 0 Then
        For col = 11 To 12
            part = CStr(values(row, col))
            result = result & Len(part) & ":" & part
        Next col
    End If
    MakeHistoryKey = result
End Function

Private Sub RefreshArchiveFilter(ByVal arc As Worksheet, ByVal last As Long)
    Dim enabled(1 To 12) As Boolean, second(1 To 12) As Boolean
    Dim firstCriteria(1 To 12) As Variant, secondCriteria(1 To 12) As Variant
    Dim operators(1 To 12) As Long, col As Long, filterCount As Long
    If arc.AutoFilterMode Then
        filterCount = Application.Min(12, arc.AutoFilter.Filters.Count)
        For col = 1 To filterCount
            With arc.AutoFilter.Filters(col)
                enabled(col) = .On
                If enabled(col) Then
                    operators(col) = .Operator
                    On Error Resume Next
                    firstCriteria(col) = .Criteria1
                    Err.Clear
                    secondCriteria(col) = .Criteria2
                    second(col) = (Err.Number = 0)
                    Err.Clear
                    On Error GoTo 0
                End If
            End With
        Next col
        arc.AutoFilterMode = False
    End If
    With arc.Range("A5:L" & last)
        .AutoFilter
        For col = 1 To 12
            If enabled(col) Then
                If second(col) Then
                    If IsEmpty(firstCriteria(col)) Then
                        .AutoFilter Field:=col, Operator:=operators(col), Criteria2:=secondCriteria(col)
                    Else
                        .AutoFilter Field:=col, Criteria1:=firstCriteria(col), Operator:=operators(col), Criteria2:=secondCriteria(col)
                    End If
                ElseIf operators(col) = 0 Then
                    .AutoFilter Field:=col, Criteria1:=firstCriteria(col)
                Else
                    .AutoFilter Field:=col, Criteria1:=firstCriteria(col), Operator:=operators(col)
                End If
            End If
        Next col
    End With
End Sub

Private Sub CommitRow(ByVal row As Long, Optional ByVal confirmDraft As Boolean = True)
    Dim sh As Worksheet, state As Worksheet, status As String, occupied As Boolean
    Dim trailer As String, notes As String, loadValue As Variant, unloadValue As Variant, col As Variant
    Dim details(1 To 1, 1 To 2) As Variant
    Set sh = Ws("DANE"): Set state = Ws("_STAN")
    trailer = RegPlate(InputValue(sh.Cells(row, 3)))
    notes = TextValue(InputValue(sh.Cells(row, 4)))
    details(1, 1) = OptionalText(InputValue(sh.Cells(row, 11)), 4000, "Trasa")
    details(1, 2) = OptionalText(InputValue(sh.Cells(row, 12)), 4000, "Numer zlecenia")
    If Len(notes) > 4000 Then Err.Raise vbObjectError + 9, , "Informacje dodatkowe: maksymalnie 4000 znaków."
    If Len(trailer) > 40 Then Err.Raise vbObjectError + 10, , "Zbyt długa rejestracja przyczepy."
    loadValue = ParsedDate(InputValue(sh.Cells(row, 5)))
    unloadValue = ParsedDate(InputValue(sh.Cells(row, 6)))
    If Not IsEmpty(loadValue) And Not IsEmpty(unloadValue) Then
        If CDbl(unloadValue) < CDbl(loadValue) Then Err.Raise vbObjectError + 11, , "Rozładunek nie może być wcześniejszy niż załadunek."
    End If
    sh.Cells(row, 3).Value2 = trailer
    sh.Cells(row, 4).Value2 = notes
    sh.Cells(row, 5).Value2 = loadValue: sh.Cells(row, 6).Value2 = unloadValue
    sh.Cells(row, 11).Resize(1, 2).NumberFormat = "@"
    sh.Cells(row, 11).Resize(1, 2).Value2 = details
    status = RowState(row)
    If status = "Puste" And confirmDraft Then
        For Each col In Array(4, 5, 6, 7, 10, 11)
            If Len(CStr(state.Cells(row, col).Value2)) > 0 Then occupied = True: Exit For
        Next col
        If occupied Then
            If state.Cells(row, 4).Value2 = "" Or state.Cells(row, 6).Value2 = "" Or state.Cells(row, 7).Value2 = "" Then
                If Not TestMode Then
                    If MsgBox("Szkic nie ma kompletnej kopii w historii. Wyczyścić go?", vbYesNo + vbExclamation, "Wyczyścić szkic?") <> vbYes Then
                        RestoreSavedRow row
                        status = RowState(row)
                    End If
                End If
            End If
        End If
    End If
    ArchiveRow row
    If sh.Cells(row, 3).Value2 <> "" Then AddPlate "PRZYCZEPY", CStr(sh.Cells(row, 3).Value2)
    If status = "Puste" Then
        state.Cells(row, 3).ClearContents
    ElseIf Val(state.Cells(row, 3).Value2) = 0 Then
        state.Range("B2").Value2 = Val(state.Range("B2").Value2) + 1
        state.Cells(row, 3).Value2 = state.Range("B2").Value2
    End If
    state.Cells(row, 4).Resize(1, 4).Value2 = sh.Cells(row, 3).Resize(1, 4).Value2
    state.Cells(row, 10).Resize(1, 2).NumberFormat = "@"
    state.Cells(row, 10).Resize(1, 2).Value2 = sh.Cells(row, 11).Resize(1, 2).Value2
    sh.Cells(row, 7).Value2 = status
    sh.Cells(row, 7).Font.Color = RGB(91, 113, 127)
    If status = "Szkic" Then sh.Cells(row, 7).Font.Color = RGB(155, 100, 15)
    If status = "Kompletny" Then sh.Cells(row, 7).Font.Color = RGB(12, 115, 112)
End Sub

Private Sub RestoreSavedRow(ByVal row As Long)
    Ws("DANE").Cells(row, 3).Resize(1, 4).Value2 = Ws("_STAN").Cells(row, 4).Resize(1, 4).Value2
    With Ws("DANE").Cells(row, 11).Resize(1, 2)
        .NumberFormat = "@"
        .Value2 = Ws("_STAN").Cells(row, 10).Resize(1, 2).Value2
    End With
End Sub

Private Function ChangedRows(ByVal affected As Range) As Object
    Dim area As Range, item As Range, rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    For Each area In affected.Areas
        For Each item In area.Rows: rows(CStr(item.Row)) = True: Next item
    Next area
    Set ChangedRows = rows
End Function

Private Sub CaptureRegistryDetails(ByVal sh As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long, ByVal offset As Long)
    Dim state As Worksheet, source As Range, snapshot As Range, cell As Range, col As Long, last As Long, upper As Long
    Set state = Ws("_STAN")
    For col = firstCol To lastCol
        last = Application.Max(last, LastRow(sh.Name, col), LastRow(state.Name, col + offset))
    Next col
    upper = sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    If upper >= FIRST_ROW Then
        Set source = sh.Range(sh.Cells(FIRST_ROW, firstCol), sh.Cells(upper, lastCol))
        If ContainsFormulas(source) Then
            ' SpecialCells jest niedostępne na chronionym rejestrze.
            For Each cell In source.Cells
                If cell.HasFormula Then
                    registryFormulas(state.Cells(cell.Row, cell.Column + offset).Address) = cell.Formula
                    last = Application.Max(last, cell.Row)
                End If
            Next cell
        End If
    End If
    If last < FIRST_ROW Then Exit Sub
    Set source = sh.Range(sh.Cells(FIRST_ROW, firstCol), sh.Cells(last, lastCol))
    Set snapshot = state.Range(state.Cells(FIRST_ROW, firstCol + offset), state.Cells(last, lastCol + offset))
    ' Plik mógł zostać zapisany bez makr. Punktem odniesienia jest stan
    ' widoczny przy otwarciu, także jeśli zawiera nieprawidłową wartość.
    snapshot.NumberFormat = "@"
    snapshot.Value2 = source.Value2
End Sub

Private Function RegistryNumberFormat(ByVal cell As Range) As String
    RegistryNumberFormat = "@"
    If cell.Worksheet.Name <> "TIRY" Then Exit Function
    Select Case cell.Column
        Case 2: RegistryNumberFormat = "0"
        Case 3 To 5: RegistryNumberFormat = "dd.mm.yyyy"
    End Select
End Function

Private Sub CommitRegistryDetails(ByVal edited As Range, ByVal offset As Long)
    Dim cell As Range, snapshot As Range, value As Variant, description As String
    Dim details(2 To 6) As Variant
    On Error GoTo Invalid
    For Each cell In edited.Cells
        value = InputValue(cell)
        If edited.Worksheet.Name = "KIEROWCY" Then
            details(cell.Column) = OptionalText(value, 100, "Telefon")
        Else
            Select Case cell.Column
                Case 2: details(cell.Column) = ParsedMileage(value)
                Case 3 To 5: details(cell.Column) = ParsedDate(value)
                Case 6: details(cell.Column) = OptionalText(value, 4000, "Uwagi")
            End Select
        End If
    Next cell
    For Each cell In edited.Cells
        cell.NumberFormat = RegistryNumberFormat(cell)
        cell.Value2 = details(cell.Column)
    Next cell
    For Each cell In edited.Cells
        Set snapshot = Ws("_STAN").Cells(cell.Row, cell.Column + offset)
        snapshot.NumberFormat = "@"
        snapshot.Value2 = details(cell.Column)
        If Not registryFormulas Is Nothing Then
            If registryFormulas.Exists(snapshot.Address) Then registryFormulas.Remove snapshot.Address
        End If
    Next cell
    Exit Sub
Invalid:
    description = Err.Description
    For Each cell In edited.Cells
        Set snapshot = Ws("_STAN").Cells(cell.Row, cell.Column + offset)
        RestoreRegistryCell cell, snapshot
    Next cell
    Err.Raise vbObjectError + 36, , description
End Sub

Private Sub RestoreRegistryCell(ByVal cell As Range, ByVal snapshot As Range)
    If Not registryFormulas Is Nothing Then
        If registryFormulas.Exists(snapshot.Address) Then
            cell.NumberFormat = "0"
            cell.Formula = registryFormulas(snapshot.Address)
            cell.NumberFormat = RegistryNumberFormat(cell)
            Exit Sub
        End If
    End If
    If VarType(snapshot.Value2) = vbString Then cell.NumberFormat = "@"
    cell.Value2 = snapshot.Value2
    cell.NumberFormat = RegistryNumberFormat(cell)
End Sub

Public Sub HandleChange(ByVal sh As Object, ByVal target As Range)
    Dim affected As Range, details As Range, item As Range, rows As Object, row As Variant, message As String, last As Long
    If busy Then Exit Sub
    If sh.Name <> "DANE" And sh.Name <> "KIEROWCY" And sh.Name <> "TIRY" And sh.Name <> "PRZYCZEPY" Then Exit Sub
    On Error GoTo Failed
    busy = True: Application.EnableEvents = False
    If sh.Name = "DANE" Then
        last = Application.Max(FIRST_ROW, LastRow("DANE"))
        Set affected = Intersect(target, sh.Range("C" & FIRST_ROW & ":F" & last & ",K" & FIRST_ROW & ":L" & last))
        If Not affected Is Nothing Then
            Set rows = ChangedRows(affected)
            For Each row In rows.Keys
                If Val(sh.Cells(CLng(row), 9).Value2) > 0 Then
                    On Error Resume Next
                    CommitRow CLng(row)
                    If Err.Number <> 0 Then
                        message = Err.Description
                        Err.Clear
                        RestoreSavedRow CLng(row)
                    End If
                    On Error GoTo Failed
                End If
            Next row
        End If
    ElseIf sh.Name = "KIEROWCY" Then
        last = Application.Max(REGISTRY_LAST_ROW, LastRow(sh.Name))
        Set affected = Intersect(target, sh.Range("B" & FIRST_ROW & ":E" & last))
        If Not affected Is Nothing Then
            Set rows = ChangedRows(affected)
            For Each row In rows.Keys
                On Error Resume Next
                If Not Intersect(target, sh.Cells(CLng(row), 2).Resize(1, 3)) Is Nothing Then SyncDriver CLng(row)
                If Err.Number <> 0 Then
                    message = Err.Description
                    Err.Clear
                End If
                Set details = Intersect(target, sh.Cells(CLng(row), 5))
                If Not details Is Nothing Then CommitRegistryDetails details, 7
                If Err.Number <> 0 Then
                    message = Err.Description
                    Err.Clear
                End If
                On Error GoTo Failed
            Next row
        End If
    Else
        last = Application.Max(REGISTRY_LAST_ROW, LastRow(sh.Name))
        If sh.Name = "TIRY" Then
            Set affected = Intersect(target, sh.Range("B" & FIRST_ROW & ":F" & last))
            If Not affected Is Nothing Then
                Set rows = ChangedRows(affected)
                For Each row In rows.Keys
                    On Error Resume Next
                    Set details = Intersect(affected, sh.Cells(CLng(row), 2).Resize(1, 5))
                    CommitRegistryDetails details, 11
                    If Err.Number <> 0 Then
                        message = Err.Description
                        Err.Clear
                    End If
                    On Error GoTo Failed
                Next row
            End If
        End If
        Set affected = Intersect(target, sh.Range("A" & FIRST_ROW & ":A" & last))
        If Not affected Is Nothing Then
            Set vehicleKeys = Nothing
            For Each item In affected.Cells
                If item.HasFormula Or IsError(item.Value2) Then Err.Raise vbObjectError + 12, , "W rejestrze wpisz tekst, nie formułę."
                item.NumberFormat = "@": item.Value2 = RegPlate(item.Value2)
            Next item
            UpdateLists
        End If
    End If
Done:
    busy = False: Application.EnableEvents = True
    On Error GoTo 0
    If message <> "" Then NotifyError message & vbCrLf & "Przywrócono poprzednie dane nieprawidłowego wiersza."
    Exit Sub
Failed:
    message = Err.Description
    Resume Done
End Sub

Private Sub SyncDriver(ByVal regRow As Long)
    Dim driverId As Long, first As Long, row As Long, driverName As String, truck As String, active As Boolean
    Dim sh As Worksheet, oldName As String, oldTruck As String, created As Boolean
    Set sh = Ws("KIEROWCY")
    driverId = Val(sh.Cells(regRow, 1).Value2)
    first = DriverFirstRow(driverId)
    On Error GoTo Invalid
    If sh.Cells(regRow, 2).HasFormula Or sh.Cells(regRow, 3).HasFormula Then Err.Raise vbObjectError + 13, , "Podaj tekst, nie formułę."
    driverName = TextValue(sh.Cells(regRow, 2).Value2): truck = RegPlate(sh.Cells(regRow, 3).Value2)
    If driverId = 0 And (driverName = "" Or truck = "") Then Exit Sub
    If driverName = "" Or truck = "" Then Err.Raise vbObjectError + 14, , "Kierowca i tir są wymagani. Aby ukryć kierowcę, ustaw Aktywny = NIE."
    If Len(driverName) > 150 Or Len(truck) > 40 Then Err.Raise vbObjectError + 15, , "Zbyt długa nazwa lub rejestracja."
    If driverId = 0 Then
        created = True
        driverId = Val(Ws("_STAN").Range("B1").Value2) + 1
        Ws("_STAN").Range("B1").Value2 = driverId
        sh.Cells(regRow, 1).Value2 = driverId
        sh.Cells(regRow, 4).Value2 = "TAK"
        SetupDriverBlock driverId, driverName, truck
        first = DriverFirstRow(driverId)
    End If
    AddPlate "TIRY", truck
    active = UCase$(TextValue(sh.Cells(regRow, 4).Value2)) <> "NIE"
    sh.Cells(regRow, 4).Value2 = IIf(active, "TAK", "NIE")
    For row = first To first + 3
        Ws("DANE").Cells(row, 1).Value2 = driverName
        Ws("DANE").Cells(row, 2).Value2 = truck
        ArchiveRow row
    Next row
    Ws("DANE").Rows(first & ":" & first + 3).Hidden = Not active
    For row = first To first + 3
        Ws("_STAN").Cells(row, 8).Value2 = driverName
        Ws("_STAN").Cells(row, 9).Value2 = truck
    Next row
    sh.Cells(regRow, 3).Value2 = truck
    Exit Sub
Invalid:
    Dim description As String: description = Err.Description
    If created Then
        sh.Cells(regRow, 1).ClearContents
        sh.Cells(regRow, 4).ClearContents
        first = 0
    End If
    If first >= FIRST_ROW Then
        oldName = CStr(Ws("_STAN").Cells(first, 8).Value2): oldTruck = CStr(Ws("_STAN").Cells(first, 9).Value2)
        sh.Cells(regRow, 2).Value2 = oldName: sh.Cells(regRow, 3).Value2 = oldTruck
        For row = first To first + 3
            Ws("DANE").Cells(row, 1).Value2 = oldName: Ws("DANE").Cells(row, 2).Value2 = oldTruck
        Next row
    End If
    Err.Raise vbObjectError + 16, , description
End Sub

Public Sub ClearRow(ByVal row As Long, Optional ByVal allowDraft As Boolean = False)
    Dim oldEvents As Boolean, description As String
    If row < FIRST_ROW Or row > LastRow("DANE") Then Err.Raise vbObjectError + 17, , "Wybierz miejsce kierowcy."
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    If RowState(row) = "Szkic" And Not allowDraft Then
        If MsgBox("Szkic nie ma kompletnej kopii w historii. Wyczyścić go?", vbYesNo + vbExclamation, "Wyczyścić szkic?") <> vbYes Then GoTo Done
    End If
    CommitRow row, False
    Ws("DANE").Cells(row, 3).Resize(1, 4).ClearContents
    Ws("DANE").Cells(row, 11).Resize(1, 2).ClearContents
    CommitRow row, False
Done:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    description = Err.Description
    RestoreSavedRow row
    Application.EnableEvents = oldEvents
    Err.Raise vbObjectError + 18, , description
End Sub

Public Sub ClearOldestForDriver(ByVal driverId As Long, Optional ByVal allowDraft As Boolean = False)
    Dim first As Long, row As Long, oldestRow As Long, order As Double, oldest As Double
    first = DriverFirstRow(driverId)
    If first = 0 Then Err.Raise vbObjectError + 19, , "Nie znaleziono kierowcy."
    oldest = 1E+30
    For row = first To first + 3
        order = Val(Ws("_STAN").Cells(row, 3).Value2)
        If order > 0 And order < oldest Then oldest = order: oldestRow = row
    Next row
    If oldestRow > 0 Then ClearRow oldestRow, allowDraft
End Sub

Public Sub RestoreRow(ByVal archiveRow As Long, ByVal destination As Long)
    Dim oldEvents As Boolean, description As String, first As Long, reg As Long, driverId As Long
    If archiveRow < FIRST_ROW Or archiveRow > LastRow("ARCHIWUM") Then Err.Raise vbObjectError + 20, , "Wybierz wpis z archiwum."
    driverId = Val(Ws("ARCHIWUM").Cells(archiveRow, 8).Value2)
    first = DriverFirstRow(driverId): reg = RegisterRow(driverId)
    If first = 0 Or destination < first Or destination > first + 3 Then Err.Raise vbObjectError + 21, , "Wybierz miejsce tego samego kierowcy."
    If UCase$(CStr(Ws("KIEROWCY").Cells(reg, 4).Value2)) = "NIE" Then Err.Raise vbObjectError + 22, , "Najpierw aktywuj kierowcę w rejestrze."
    If RowState(destination) <> "Puste" Then Err.Raise vbObjectError + 23, , "Najpierw zwolnij wybrane miejsce."
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Ws("DANE").Cells(destination, 3).Resize(1, 4).Value2 = Ws("ARCHIWUM").Cells(archiveRow, 3).Resize(1, 4).Value2
    Ws("DANE").Cells(destination, 11).Resize(1, 2).Value2 = Ws("ARCHIWUM").Cells(archiveRow, 11).Resize(1, 2).Value2
    CommitRow destination
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    description = Err.Description
    RestoreSavedRow destination
    Application.EnableEvents = oldEvents
    Err.Raise vbObjectError + 24, , description
End Sub

Public Function FlushAll() As Boolean
    Dim row As Long, col As Long, last As Long, oldEvents As Boolean, changed As Boolean
    Dim current As Variant, saved As Variant, checkAll As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    LoadKeys
    last = LastRow("DANE")
    If last >= FIRST_ROW Then
        current = Ws("DANE").Range("A" & FIRST_ROW & ":L" & last).Value2
        saved = Ws("_STAN").Range("A" & FIRST_ROW & ":K" & last).Value2
        checkAll = ContainsFormulas(Ws("DANE").Range("C" & FIRST_ROW & ":F" & last)) _
            Or ContainsFormulas(Ws("DANE").Range("K" & FIRST_ROW & ":L" & last))
        For row = 1 To UBound(current, 1)
            changed = checkAll Or CStr(current(row, 1)) <> CStr(saved(row, 8)) Or CStr(current(row, 2)) <> CStr(saved(row, 9))
            For col = 3 To 6
                If CStr(current(row, col)) <> CStr(saved(row, col + 1)) Then changed = True
            Next col
            For col = 11 To 12
                If CStr(current(row, col)) <> CStr(saved(row, col - 1)) Then changed = True
            Next col
            If Not changed Then
                If Len(CStr(current(row, 3))) > 0 And Len(CStr(current(row, 5))) > 0 And Len(CStr(current(row, 6))) > 0 Then
                    changed = Not keys.Exists(MakeHistoryKey(current(row, 9), current, row))
                End If
            End If
            If changed Then CommitRow FIRST_ROW + row - 1
        Next row
    End If
    FlushAll = True
Done:
    Application.EnableEvents = oldEvents
    Exit Function
Failed:
    NotifyError "Nie zapisano pliku: " & Err.Description
    Resume Done
End Function

Private Function ContainsFormulas(ByVal cells As Range) As Boolean
    Dim formulas As Variant
    formulas = cells.HasFormula
    If IsNull(formulas) Then ContainsFormulas = True Else ContainsFormulas = CBool(formulas)
End Function

Private Sub NotifyError(ByVal message As String)
    If TestMode Then Err.Raise vbObjectError + 25, , message
    MsgBox message, vbExclamation, "Floteka"
End Sub

Public Sub DodajKierowce()
    Dim driverName As String, truck As String, driverId As Long
    On Error GoTo Failed
    driverName = InputBox("Imię i nazwisko kierowcy:", "Dodaj kierowcę")
    If Trim$(driverName) = "" Then Exit Sub
    truck = InputBox("Numer rejestracyjny tira (jeden na kierowcę):", "Przypisz tira")
    If Trim$(truck) = "" Then Exit Sub
    driverId = AddDriverRecord(driverName, truck)
    Ws("DANE").Activate: Ws("DANE").Cells(DriverFirstRow(driverId), 3).Select
    Exit Sub
Failed:
    NotifyError Err.Description
End Sub

Public Sub DodajTir()
    AddVehiclePrompt "TIRY", "Dodaj tira"
End Sub

Public Sub DodajPrzyczepe()
    AddVehiclePrompt "PRZYCZEPY", "Dodaj przyczepę"
End Sub

Private Sub AddVehiclePrompt(ByVal sheetName As String, ByVal title As String)
    Dim number As String
    On Error GoTo Failed
    number = InputBox("Numer rejestracyjny:", title)
    If Trim$(number) = "" Then Exit Sub
    Application.EnableEvents = False
    AddPlate sheetName, number
    Application.EnableEvents = True
    Exit Sub
Failed:
    Application.EnableEvents = True
    NotifyError Err.Description
End Sub

Public Sub WyczyscMiejsce()
    Dim row As Long, caller As String
    On Error GoTo Failed
    If VarType(Application.Caller) = vbString Then caller = CStr(Application.Caller)
    If Left$(caller, 6) = "clear_" Then
        row = CLng(Mid$(caller, 7))
    ElseIf ActiveSheet.Name = "DANE" Then
        row = ActiveCell.Row
    End If
    ClearRow row
    Exit Sub
Failed:
    NotifyError Err.Description
End Sub

Public Sub ZwolnijNajstarsze()
    On Error GoTo Failed
    If ActiveSheet.Name <> "DANE" Or ActiveCell.Row < FIRST_ROW Then
        Ws("DANE").Activate
        MsgBox "Zaznacz wiersz kierowcy w DANE, następnie kliknij Zwolnij najstarsze.", vbInformation
        Exit Sub
    End If
    ClearOldestForDriver CLng(Ws("DANE").Cells(ActiveCell.Row, 9).Value2)
    Exit Sub
Failed:
    NotifyError Err.Description
End Sub

Public Sub PrzywrocWpis()
    Dim row As Long, first As Long, destination As Long, driverId As Long, choice As Variant, suggested As Long
    On Error GoTo Failed
    If ActiveSheet.Name <> "ARCHIWUM" Or ActiveCell.Row < FIRST_ROW Then
        Ws("ARCHIWUM").Activate
        MsgBox "Zaznacz wpis w ARCHIWUM i kliknij Przywróć wpis.", vbInformation
        Exit Sub
    End If
    row = ActiveCell.Row
    driverId = CLng(Ws("ARCHIWUM").Cells(row, 8).Value2): first = DriverFirstRow(driverId)
    If first = 0 Then Err.Raise vbObjectError + 26, , "Brak kierowcy w rejestrze."
    For destination = first To first + 3
        If RowState(destination) = "Puste" Then suggested = destination - first + 1: Exit For
    Next destination
    If suggested = 0 Then Err.Raise vbObjectError + 27, , "Najpierw zwolnij jedno miejsce tego kierowcy."
    If CStr(Ws("DANE").Cells(first, 2).Value2) <> CStr(Ws("ARCHIWUM").Cells(row, 2).Value2) Then
        If MsgBox("Wpis zostanie przywrócony z OBECNYM tirem: " & Ws("DANE").Cells(first, 2).Value2 & ". Historyczny tir: " & Ws("ARCHIWUM").Cells(row, 2).Value2 & ". Kontynuować?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    choice = Application.InputBox("Numer pustego miejsca (1–4):", "Przywróć wpis", suggested, Type:=1)
    If VarType(choice) = vbBoolean Then Exit Sub
    If choice < 1 Or choice > 4 Or choice <> Fix(choice) Then Err.Raise vbObjectError + 28, , "Podaj liczbę 1–4."
    RestoreRow row, first + CLng(choice) - 1
    Ws("DANE").Activate: Ws("DANE").Cells(first + CLng(choice) - 1, 3).Select
    Exit Sub
Failed:
    NotifyError Err.Description
End Sub

Public Sub ZapiszWszystko()
    If FlushAll() Then ThisWorkbook.Save
End Sub

Public Sub KopiaPliku()
    Dim target As Variant
    On Error GoTo Failed
    If Not FlushAll() Then Exit Sub
    target = Application.GetSaveAsFilename("Floteka-kopia-" & Format(Now, "yyyy-mm-dd-hhnnss") & ".xlsm", "Excel z makrami (*.xlsm),*.xlsm")
    If VarType(target) = vbBoolean Then Exit Sub
    If LCase$(CStr(target)) = LCase$(ThisWorkbook.FullName) Then Err.Raise vbObjectError + 29, , "Wybierz inny plik niż bieżący."
    ThisWorkbook.SaveCopyAs CStr(target)
    Exit Sub
Failed:
    NotifyError Err.Description
End Sub
