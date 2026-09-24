Attribute VB_Name = "TimelineBates"
Option Explicit

' Bates is a managed final paragraph in the existing Entry Box. Only that range
' is inserted/deleted, so toggling preserves every rich-text run in the body.
' PowerPoint uppercases tag values: UTF-16 hex keeps the source's exact case.
Private Const TAG_DATA As String = "TLBatesData"
Private Const TAG_RENDERED As String = "TLBatesRendered"
Private Const TAG_VISIBLE As String = "TLBatesVisible"

Public Sub ToggleTimelineBates(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    On Error Resume Next
    Set sld = ActiveWindow.View.slide
    If sld Is Nothing Then Set sld = ActiveWindow.Selection.SlideRange(1)
    On Error GoTo Fail
    If sld Is Nothing Then
        MsgBox "Select a slide with imported timeline entries first.", vbExclamation, "Toggle Bates Numbers"
        Exit Sub
    End If

    Dim boxes As New Collection, box As Shape, showBates As Boolean, n As Long
    CollectBatesBoxes sld.Shapes, boxes, ""
    showBates = Not TimelineBatesAreVisible(sld)
    For Each box In boxes
        If Len(BatesTextOf(box)) > 0 Then
            ApplyTimelineEntryBates box, showBates
            n = n + 1
        End If
    Next box
    sld.Tags.Add TAG_VISIBLE, IIf(showBates, "1", "0")
    If n = 0 Then
        MsgBox "No Bates numbers were stored with entries on this slide. Import a spreadsheet with a Bates column to add them.", _
               vbInformation, "Toggle Bates Numbers"
    End If
    Exit Sub
Fail:
    MsgBox "Could not toggle Bates numbers:" & vbCrLf & Err.Description, vbExclamation, "Toggle Bates Numbers"
End Sub

' The creator calls this after grouping. The box keeps its own metadata so it
' still works after ungrouping, copying, or moving entries between slides.
Public Sub SetTimelineEntryBates(ByVal entry As Shape, ByVal batesText As String, _
                                 Optional ByVal showBates As Boolean = False)
    Dim box As Shape
    Set box = BatesEntryBox(entry)
    If box Is Nothing Then Exit Sub
    ApplyTimelineEntryBates box, False
    batesText = Trim$(NormalizeBatesBreaks(batesText))
    entry.Tags.Add "TLBates", batesText
    entry.Tags.Add TAG_DATA, EncodeBates(batesText)
    box.Tags.Add "TLBates", batesText
    box.Tags.Add TAG_DATA, EncodeBates(batesText)
    ApplyTimelineEntryBates box, showBates
End Sub

Public Function TimelineBatesAreVisible(ByVal sld As slide) As Boolean
    On Error Resume Next
    TimelineBatesAreVisible = (sld.Tags(TAG_VISIBLE) = "1")
    On Error GoTo 0
End Function

Public Function TimelineEntryBatesVisible(ByVal entry As Shape) As Boolean
    Dim box As Shape
    Set box = BatesEntryBox(entry)
    If box Is Nothing Then Exit Function
    TimelineEntryBatesVisible = (BatesTag(box, TAG_VISIBLE) = "1")
End Function

' Public helpers let Parse Content and Auto Center work on the body only.
' Save TimelineEntryBatesVisible, hide, format, then apply the saved state.
Public Sub ApplyTimelineEntryBates(ByVal entry As Shape, ByVal showBates As Boolean)
    Dim box As Shape, bates As String, suffix As String, suffixStart As Long
    Set box = BatesEntryBox(entry)
    If box Is Nothing Then Exit Sub
    If BatesTag(box, TAG_DATA) = "" And BatesTag(box, "TLBates") = "" Then
        bates = BatesTextOf(entry)
        If Len(bates) > 0 Then
            box.Tags.Add TAG_DATA, EncodeBates(bates)
            box.Tags.Add "TLBates", bates
        End If
    End If

    ' Never remove arbitrary body text: the exact tagged suffix must still be
    ' present. If another tool replaced the description, only its stale marker
    ' is cleared and the current Bates value can be safely appended anew.
    suffixStart = RenderedBatesStart(box, suffix)
    If suffixStart > 0 Then box.TextFrame.TextRange.Characters(suffixStart, Len(suffix)).Delete
    box.Tags.Add TAG_VISIBLE, "0"
    box.Tags.Add TAG_RENDERED, ""
    If Not showBates Then Exit Sub
    bates = BatesTextOf(box)
    If Len(bates) = 0 Then Exit Sub

    Dim bodyLen As Long, fontSize As Single, footer As TextRange
    bodyLen = Len(box.TextFrame.TextRange.text)
    fontSize = MainBatesFontSize(box)
    suffix = vbCr & bates
    box.TextFrame.TextRange.InsertAfter suffix
    Set footer = box.TextFrame.TextRange.Characters(bodyLen + 2, Len(bates))
    With footer
        .Font.Size = fontSize * 0.6
        .Font.Bold = msoFalse
        .Font.Italic = msoTrue
        .Font.Underline = msoFalse
        .Font.Color.RGB = RGB(128, 128, 128)
        .IndentLevel = 1
        With .ParagraphFormat
            .Alignment = ppAlignRight
            .Bullet.Visible = msoFalse
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    End With
    box.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    box.Tags.Add TAG_RENDERED, EncodeBates(suffix)
    box.Tags.Add TAG_VISIBLE, "1"
End Sub

Public Sub RefreshTimelineBates(ByVal entry As Shape)
    If TimelineEntryBatesVisible(entry) Then ApplyTimelineEntryBates entry, True
End Sub

' Read-only body extraction: Bates never enters AI rewording or title previews.
Public Function TimelineEntryBodyText(ByVal entry As Shape) As String
    Dim box As Shape, suffix As String, suffixStart As Long
    Set box = BatesEntryBox(entry)
    If box Is Nothing Then Exit Function
    TimelineEntryBodyText = box.TextFrame.TextRange.text
    suffixStart = RenderedBatesStart(box, suffix)
    If suffixStart > 0 Then TimelineEntryBodyText = Left$(TimelineEntryBodyText, suffixStart - 1)
End Function

Private Function RenderedBatesStart(ByVal box As Shape, ByRef suffix As String) As Long
    If BatesTag(box, TAG_VISIBLE) <> "1" Then Exit Function
    suffix = DecodeBates(BatesTag(box, TAG_RENDERED))
    If Len(suffix) = 0 Then Exit Function
    Dim currentText As String
    currentText = box.TextFrame.TextRange.text
    If Len(currentText) < Len(suffix) Then Exit Function
    If StrComp(Right$(currentText, Len(suffix)), suffix, vbBinaryCompare) = 0 Then _
        RenderedBatesStart = Len(currentText) - Len(suffix) + 1
End Function

' Prefer the main text's predominant point size; this tolerates a larger title,
' mixed bold/italic runs, and text resized after the citation was first shown.
Private Function MainBatesFontSize(ByVal box As Shape) As Single
    Dim tr As TextRange, run As TextRange, sizes() As Single, weights() As Long
    Dim count As Long, i As Long, j As Long, winner As Long, found As Boolean
    Set tr = box.TextFrame.TextRange
    ReDim sizes(1 To tr.Runs.count + 1)
    ReDim weights(1 To tr.Runs.count + 1)
    For i = 1 To tr.Runs.count
        Set run = tr.Runs(i)
        If run.Font.Size > 0 Then
            found = False
            For j = 1 To count
                If Abs(sizes(j) - run.Font.Size) < 0.01 Then
                    weights(j) = weights(j) + Len(run.text)
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                count = count + 1
                sizes(count) = run.Font.Size
                weights(count) = Len(run.text)
            End If
        End If
    Next i
    For i = 1 To count
        If winner = 0 Then
            winner = i
        ElseIf weights(i) > weights(winner) Then
            winner = i
        End If
    Next i
    If winner > 0 Then MainBatesFontSize = sizes(winner) Else MainBatesFontSize = 12
End Function

Private Sub CollectBatesBoxes(ByVal shapesColl As Object, ByVal boxes As Collection, ByVal inheritedBates As String)
    Dim shp As Shape, value As String
    For Each shp In shapesColl
        value = BatesTextOf(shp)
        If Len(value) = 0 Then value = inheritedBates
        If IsBatesEntryBox(shp) Then
            If BatesTag(shp, TAG_DATA) = "" And Len(value) > 0 Then shp.Tags.Add TAG_DATA, EncodeBates(value)
            boxes.Add shp
        ElseIf shp.Type = msoGroup Then
            CollectBatesBoxes shp.GroupItems, boxes, value
        End If
    Next shp
End Sub

Private Function BatesEntryBox(ByVal entry As Shape) As Shape
    If entry Is Nothing Then Exit Function
    If IsBatesEntryBox(entry) Then
        Set BatesEntryBox = entry
    ElseIf entry.Type = msoGroup Then
        Dim child As Shape, found As Shape
        For Each child In entry.GroupItems
            Set found = BatesEntryBox(child)
            If Not found Is Nothing Then Set BatesEntryBox = found: Exit Function
        Next child
    End If
End Function

Private Function IsBatesEntryBox(ByVal shp As Shape) As Boolean
    If StrComp(BatesTag(shp, "GroupStyle"), "Entry Box", vbTextCompare) <> 0 Then Exit Function
    IsBatesEntryBox = (shp.HasTextFrame = msoTrue)
End Function

Private Function BatesTextOf(ByVal shp As Shape) As String
    Dim encoded As String
    encoded = BatesTag(shp, TAG_DATA)
    If Len(encoded) > 0 Then BatesTextOf = DecodeBates(encoded) Else BatesTextOf = BatesTag(shp, "TLBates")
End Function

Private Function BatesTag(ByVal shp As Shape, ByVal key As String) As String
    On Error Resume Next
    BatesTag = shp.Tags(key)
    On Error GoTo 0
End Function

Private Function NormalizeBatesBreaks(ByVal value As String) As String
    value = Replace$(value, vbCrLf, vbCr)
    NormalizeBatesBreaks = Replace$(value, vbLf, vbCr)
End Function

Private Function EncodeBates(ByVal value As String) As String
    Dim i As Long, code As Long, result As String
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1))
        If code < 0 Then code = code + 65536
        result = result & Right$("0000" & Hex$(code), 4)
    Next i
    EncodeBates = result
End Function

Private Function DecodeBates(ByVal value As String) As String
    On Error GoTo Invalid
    If Len(value) Mod 4 <> 0 Then Exit Function
    Dim i As Long, code As Long, result As String
    For i = 1 To Len(value) Step 4
        code = CLng("&H" & Mid$(value, i, 4))
        If code > 32767 Then code = code - 65536
        result = result & ChrW$(code)
    Next i
    DecodeBates = result
Invalid:
End Function
