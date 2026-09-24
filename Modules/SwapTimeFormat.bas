Attribute VB_Name = "SwapTimeFormat"
Option Explicit

Public Sub SwapTimeFormats(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    ProcessShapes sld.Shapes
    Exit Sub
Fail:
    MsgBox "Could not swap time formats:" & vbCrLf & Err.Description, vbExclamation, "Swap Time Format"
End Sub

Public Sub ProcessShapes(ByVal shapesCollection As Object)
    Dim shp As Shape, tbl As Table, row As Long, col As Long
    For Each shp In shapesCollection
        If shp.Type = msoGroup Then
            ProcessShapes shp.GroupItems
        ElseIf shp.HasTable Then
            Set tbl = shp.Table
            For row = 1 To tbl.Rows.count
                For col = 1 To tbl.Columns.count
                    SwapShapeTimeTokens tbl.cell(row, col).Shape
                Next col
            Next row
        ElseIf shp.HasTextFrame Then
            SwapShapeTimeTokens shp
        End If
    Next shp
End Sub

' Work from the end so replacing one token cannot move the remaining matches.
' Updating only a token's TextRange preserves all other rich-text formatting.
' Bates numbers can resemble times: remove the managed footer before matching.
Private Sub SwapShapeTimeTokens(ByVal shp As Shape)
    If Not shp.TextFrame.HasText Then Exit Sub
    Dim showBates As Boolean, matches As Object, token As Object, regex As Object
    Dim i As Long, replacement As String
    showBates = TimelineEntryBatesVisible(shp)
    If showBates Then ApplyTimelineEntryBates shp, False
    Set regex = TimeTokenRegex()
    Set matches = regex.Execute(shp.TextFrame.TextRange.text)
    For i = matches.count - 1 To 0 Step -1
        Set token = matches(i)
        replacement = SwapTimeToken(token)
        If replacement <> token.Value Then _
            shp.TextFrame.TextRange.Characters(token.FirstIndex + 1, token.Length).text = replacement
    Next i
    If showBates Then ApplyTimelineEntryBates shp, True
End Sub

' Retained for callers that need plain text. Each token chooses its own toggle
' direction, so dates, words, range delimiters, and seconds survive unchanged.
Public Function SwapFormat(ByVal text As String) As String
    Dim matches As Object, token As Object, regex As Object, i As Long, replacement As String
    Set regex = TimeTokenRegex()
    Set matches = regex.Execute(text)
    For i = matches.count - 1 To 0 Step -1
        Set token = matches(i)
        replacement = SwapTimeToken(token)
        text = Left$(text, token.FirstIndex) & replacement & Mid$(text, token.FirstIndex + token.Length + 1)
    Next i
    SwapFormat = text
End Function

Private Function TimeTokenRegex() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    ' H:MM[:SS] with optional meridiem, or an hour with required meridiem.
    ' Negative lookahead prevents matching a prefix of an invalid long token.
    regex.Pattern = "\b([0-9]{1,2}):([0-9]{2})(:([0-9]{2}))?([ \t]*([AP])(?:[ \t]*M|\.[ \t]*M\.?))?(?![A-Za-z0-9_:])" & _
                    "|\b([0-9]{1,2})[ \t]*([AP])(?:[ \t]*M|\.[ \t]*M\.?)(?![A-Za-z0-9_:])"
    Set TimeTokenRegex = regex
End Function

Private Function SwapTimeToken(ByVal token As Object) As String
    Dim hh As Long, mm As Long, ss As Long, meridiem As String, hasSeconds As Boolean
    SwapTimeToken = token.Value
    If Len(token.SubMatches(0)) > 0 Then
        hh = CLng(token.SubMatches(0))
        mm = CLng(token.SubMatches(1))
        hasSeconds = (Len(token.SubMatches(3)) > 0)
        If hasSeconds Then ss = CLng(token.SubMatches(3))
        meridiem = UCase$(token.SubMatches(5))
    Else
        hh = CLng(token.SubMatches(6))
        meridiem = UCase$(token.SubMatches(7))
    End If
    If mm > 59 Or ss > 59 Then Exit Function
    If Len(meridiem) > 0 Then
        If hh < 1 Or hh > 12 Then Exit Function
        hh = hh Mod 12
        If meridiem = "P" Then hh = hh + 12
        SwapTimeToken = TwoTimeDigits(hh) & ":" & TwoTimeDigits(mm)
        If hasSeconds Then SwapTimeToken = SwapTimeToken & ":" & TwoTimeDigits(ss)
    Else
        If hh > 23 Then Exit Function
        If hh >= 12 Then meridiem = "PM" Else meridiem = "AM"
        hh = hh Mod 12
        If hh = 0 Then hh = 12
        SwapTimeToken = TwoTimeDigits(hh) & ":" & TwoTimeDigits(mm)
        If hasSeconds Then SwapTimeToken = SwapTimeToken & ":" & TwoTimeDigits(ss)
        SwapTimeToken = SwapTimeToken & " " & meridiem
    End If
End Function

Private Function TwoTimeDigits(ByVal value As Long) As String
    TwoTimeDigits = Right$("0" & CStr(value), 2)
End Function

