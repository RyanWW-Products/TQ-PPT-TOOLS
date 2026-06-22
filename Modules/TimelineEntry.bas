Attribute VB_Name = "TimelineEntry"
Option Explicit

' Make Entry: a navy date box + white entry box (GroupStyle-tagged, drop-shadowed)
' plus a LeadingLine, all in ONE group. The core CreateTimelineEntry is reused by
' the Timeline Creator so imported entries are identical to hand-made ones.

Sub EntryGenerator(control As IRibbonControl)
    Dim oSlide As slide, dummy As Single
    If ActiveWindow.Selection.SlideRange.count = 0 Then
        MsgBox "Please select a slide first!", vbExclamation
        Exit Sub
    End If
    Set oSlide = ActiveWindow.Selection.SlideRange(1)
    ' Manual entry: no real date, so name it with a sequential counter (TimelineEntry 0001, 0002, ...)
    Dim manualSuffix As String
    manualSuffix = Format$(NextManualEntryNumber(oSlide), "0000")
    ' lineX/lineTopY < 0 -> Make-Entry default (centered, 50pt stub above)
    CreateTimelineEntry oSlide, "12/12/2099", "Lorem Ipsum", 100, 100, 2 * 72, _
                        -1, -1, False, 1, dummy, CDate("12/12/2099"), "Days", manualSuffix
End Sub

' Next sequential number for a manual entry on this slide ("TimelineEntry 0001", "0002", ...).
' Date-named imported entries are ignored (their suffix contains spaces, so isn't numeric).
Private Function NextManualEntryNumber(ByVal oSlide As slide) As Long
    Dim shp As Shape, mx As Long, suff As String
    mx = 0
    For Each shp In oSlide.Shapes
        If Left$(shp.Name, 14) = "TimelineEntry " Then
            suff = Mid$(shp.Name, 15)
            If IsNumeric(suff) Then
                If CLng(suff) > mx Then mx = CLng(suff)
            End If
        End If
    Next shp
    NextManualEntryNumber = mx + 1
End Function

' Build one entry; return the final group. lineX/lineTopY >= 0 place the leader
' line exactly there (the Timeline Creator passes a date-accurate X + bar bottom);
' < 0 uses the Make-Entry default. boxesHeight returns the date+entry box height
' (the caller uses it to stack the next entry).
Public Function CreateTimelineEntry(ByVal oSlide As slide, ByVal dateText As String, ByVal descText As String, _
        ByVal boxLeft As Single, ByVal boxTop As Single, ByVal boxW As Single, _
        ByVal lineX As Single, ByVal lineTopY As Single, ByVal fixedWidth As Boolean, _
        ByVal fontScale As Single, ByRef boxesHeight As Single, _
        ByVal entryDate As Date, ByVal unitType As String, _
        Optional ByVal nameOverride As String = "") As Shape
    Static lineCounter As Long
    Dim oDateBox As Shape, oEntryBox As Shape, oBoxGroup As Shape, oLine As Shape, entryGroup As Shape
    Dim suffix As String
    lineCounter = lineCounter + 1
    If nameOverride <> "" Then          ' manual entry: counter suffix (no real date to encode)
        suffix = nameOverride
    Else
        suffix = DateNameSuffix(entryDate, unitType)
    End If

    ' Date box (navy fill, white bold text)
    Set oDateBox = oSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, boxLeft, boxTop, boxW, 50)
    With oDateBox
        .fill.ForeColor.RGB = RGB(4, 61, 102)
        .line.Weight = 1.5
        .line.ForeColor.RGB = RGB(0, 0, 0)
        With .TextFrame
            .MarginTop = 3: .MarginBottom = 3
            If fixedWidth Then .WordWrap = msoTrue
            With .TextRange
                .text = dateText
                .Font.Name = "Arial"
                .Font.Size = 12 * fontScale
                .Font.Bold = msoTrue
                .Font.Color.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = ppAlignCenter
            End With
            .AutoSize = ppAutoSizeShapeToFitText
            .VerticalAnchor = msoAnchorBottom
        End With
    End With
    oDateBox.Tags.Add "GroupStyle", "Date Box"
    ' best-effort drop shadow on the date (header) text
    On Error Resume Next
    With oDateBox.TextFrame2.TextRange.Font.Shadow
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 3
        .Transparency = 0.5
        .Size = 100
        .OffsetX = 1.5
        .OffsetY = 1.5
    End With
    On Error GoTo 0

    ' Entry box (white fill, BLACK text)
    Set oEntryBox = oSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, oDateBox.Left, _
                                             oDateBox.Top + oDateBox.Height, boxW, 50)
    With oEntryBox
        .fill.ForeColor.RGB = RGB(255, 255, 255)
        .line.Weight = 1.5
        .line.ForeColor.RGB = RGB(0, 0, 0)
        With .TextFrame
            .MarginTop = 3: .MarginBottom = 3
            If fixedWidth Then .WordWrap = msoTrue
            With .TextRange
                .text = descText
                .Font.Name = "Arial"
                .Font.Size = 12 * fontScale
                .Font.Color.RGB = RGB(0, 0, 0)
                .ParagraphFormat.Alignment = ppAlignLeft
            End With
            .AutoSize = ppAutoSizeShapeToFitText
        End With
    End With
    oEntryBox.Tags.Add "GroupStyle", "Entry Box"

    boxesHeight = oDateBox.Height + oEntryBox.Height

    Set oBoxGroup = oSlide.Shapes.range(Array(oDateBox.Name, oEntryBox.Name)).Group
    With oBoxGroup.Shadow
        .Type = msoShadow21
        .IncrementOffsetX 3
        .IncrementOffsetY 3
    End With

    Dim lx As Single, lty As Single
    If lineX < 0 Then lx = oBoxGroup.Left + oBoxGroup.Width / 2 Else lx = lineX
    If lineTopY < 0 Then lty = oBoxGroup.Top - 50 Else lty = lineTopY
    Set oLine = oSlide.Shapes.AddLine(lx, lty, lx, oDateBox.Top + oDateBox.Height)
    oLine.Name = "LeadingLine" & Format$(lineCounter, "00")
    With oLine.line
        .Weight = 1.5
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    oLine.ZOrder msoSendToBack

    ' group [boxes + line] using the still-unique auto names, then rename every part from
    ' FRESH references walked out of the new group. Re-grouping invalidates the original child
    ' references (the boxes are now two levels deep, so oDateBox.Name would raise 424
    ' "Object required"), so we never touch oDateBox/oEntryBox/oBoxGroup/oLine again here.
    ' Renaming after grouping (not before) also keeps two same-date entries from colliding on
    ' the names used to group.
    Set entryGroup = oSlide.Shapes.range(Array(oBoxGroup.Name, oLine.Name)).Group
    RenameEntryParts entryGroup, suffix
    Set CreateTimelineEntry = entryGroup
End Function

' Rename the parts of a freshly-built entry group to the date-encoded scheme, using live
' references taken from the group itself (the pre-grouping ones can be stale). Date vs entry
' box is told apart by the GroupStyle tag; the leader line by its name prefix.
Private Sub RenameEntryParts(ByVal entryGroup As Shape, ByVal suffix As String)
    Dim child As Shape, gc As Shape
    entryGroup.name = "TimelineEntry " & suffix
    For Each child In entryGroup.GroupItems
        If child.Type = msoGroup Then
            child.name = "BoxGroup " & suffix
            For Each gc In child.GroupItems
                Select Case ShapeTag(gc, "GroupStyle")
                    Case "Date Box": gc.name = "DateBox " & suffix
                    Case "Entry Box": gc.name = "Entry Box " & suffix
                End Select
            Next gc
        ElseIf child.name Like "LeadingLine*" Then
            child.name = "LeadingLine " & suffix
        End If
    Next child
End Sub

Private Function ShapeTag(ByVal shp As Shape, ByVal key As String) As String
    On Error Resume Next
    ShapeTag = shp.Tags(key)
    On Error GoTo 0
End Function

' "YYYY MM DD HH MM" for a date, zeroing parts finer than the timeline unit
' (Years -> "2020 00 00 00 00"; Months -> "2020 06 00 00 00"; Days -> "2020 06 15 00 00";
'  Hours -> "2020 06 15 09 00"). Minutes stay 00 (Hours is the finest unit).
Private Function DateNameSuffix(ByVal d As Date, ByVal t As String) As String
    Dim yy As String, mo As String, dd As String, hh As String, mn As String
    yy = Format$(Year(d), "0000")
    mo = "00": dd = "00": hh = "00": mn = "00"
    Select Case t
        Case "Hours"
            mo = Format$(Month(d), "00"): dd = Format$(Day(d), "00"): hh = Format$(Hour(d), "00")
        Case "Days"
            mo = Format$(Month(d), "00"): dd = Format$(Day(d), "00")
        Case "Months"
            mo = Format$(Month(d), "00")
    End Select
    DateNameSuffix = yy & " " & mo & " " & dd & " " & hh & " " & mn
End Function
