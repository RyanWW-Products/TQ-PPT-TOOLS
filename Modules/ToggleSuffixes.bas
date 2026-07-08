Attribute VB_Name = "ToggleSuffixes"
' Toggle 1 <-> 1st on a datebar's day-number cells. Works on the object-based "Datebar NNN"
' group (the two-bar Days bottom row, or any bar with bare day-number cells).
Sub ToggleDaySuffixes(control As IRibbonControl)
    Dim pptSlide As slide, bar As Shape, cells As Collection, shp As Shape
    Dim i As Long, dayNum As Integer, txt As String, body As String, suffixOn As Boolean

    On Error Resume Next
    Set pptSlide = ActiveWindow.View.slide
    On Error GoTo 0
    If pptSlide Is Nothing Then MsgBox "Select a slide with a datebar first.", vbExclamation: Exit Sub

    Set bar = FindDateBar(pptSlide)
    If bar Is Nothing Then MsgBox "No datebar found on this slide.", vbExclamation: Exit Sub

    Set cells = New Collection
    CollectDayCells bar, cells
    If cells.count = 0 Then
        MsgBox "No day-number cells found (this toggles a day-numbered datebar, e.g. a two-bar Days bar).", vbExclamation
        Exit Sub
    End If

    ' current state from the first day cell
    suffixOn = HasDaySuffix(Trim$(cells(1).TextFrame.TextRange.text))

    For i = 1 To cells.count
        Set shp = cells(i)
        txt = Trim$(shp.TextFrame.TextRange.text)
        body = txt
        If HasDaySuffix(txt) Then body = Left$(txt, Len(txt) - 2)
        If IsNumeric(body) Then
            dayNum = CInt(body)
            If suffixOn Then
                shp.TextFrame.TextRange.text = CStr(dayNum)
            Else
                shp.TextFrame.TextRange.text = CStr(dayNum) & GetDaySuffix(dayNum)
            End If
        End If
    Next i
End Sub

' Collect leaf cells whose text is a 1-2 digit day number (1-31), with or without a suffix.
Private Sub CollectDayCells(ByVal shp As Shape, ByVal acc As Collection)
    Dim it As Shape, txt As String, body As String
    If shp.Type = msoGroup Then
        For Each it In shp.GroupItems
            CollectDayCells it, acc
        Next it
    Else
        txt = ""
        On Error Resume Next
        txt = Trim$(shp.TextFrame.TextRange.text)
        On Error GoTo 0
        If txt <> "" Then
            body = txt
            If HasDaySuffix(txt) Then body = Left$(txt, Len(txt) - 2)
            If Len(body) >= 1 And Len(body) <= 2 And IsNumeric(body) Then
                If CInt(body) >= 1 And CInt(body) <= 31 Then acc.Add shp
            End If
        End If
    End If
End Sub

Private Function HasDaySuffix(ByVal s As String) As Boolean
    Dim suf As String, body As String
    s = Trim$(s)
    If Len(s) < 3 Then Exit Function
    suf = LCase$(Right$(s, 2))
    If suf <> "st" And suf <> "nd" And suf <> "rd" And suf <> "th" Then Exit Function
    body = Left$(s, Len(s) - 2)
    HasDaySuffix = IsNumeric(body)
End Function


Function GetDaySuffix(ByVal day As Integer) As String
    Select Case day
        Case 1, 21, 31: GetDaySuffix = "st"
        Case 2, 22: GetDaySuffix = "nd"
        Case 3, 23: GetDaySuffix = "rd"
        Case Else: GetDaySuffix = "th"
    End Select
End Function

Function ShapeExists(slide As slide, shapeName As String) As Boolean
    Dim shp As Shape
    For Each shp In slide.Shapes
        If shp.Name = shapeName Then
            ShapeExists = True
            Exit Function
        End If
    Next shp
    ShapeExists = False
End Function

