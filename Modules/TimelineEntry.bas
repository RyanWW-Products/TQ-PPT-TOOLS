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
    ' lineX/lineTopY < 0 -> Make-Entry default (centered, 50pt stub above)
    CreateTimelineEntry oSlide, "12/12/2099", "Lorem Ipsum", 100, 100, 2 * 72, _
                        -1, -1, False, 1, dummy
End Sub

' Build one entry; return the final group. lineX/lineTopY >= 0 place the leader
' line exactly there (the Timeline Creator passes a date-accurate X + bar bottom);
' < 0 uses the Make-Entry default. boxesHeight returns the date+entry box height
' (the caller uses it to stack the next entry).
Public Function CreateTimelineEntry(ByVal oSlide As slide, ByVal dateText As String, ByVal descText As String, _
        ByVal boxLeft As Single, ByVal boxTop As Single, ByVal boxW As Single, _
        ByVal lineX As Single, ByVal lineTopY As Single, ByVal fixedWidth As Boolean, _
        ByVal fontScale As Single, ByRef boxesHeight As Single) As Shape
    Static lineCounter As Long
    Dim oDateBox As Shape, oEntryBox As Shape, oBoxGroup As Shape, oLine As Shape
    lineCounter = lineCounter + 1

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

    Set CreateTimelineEntry = oSlide.Shapes.range(Array(oBoxGroup.Name, oLine.Name)).Group
End Function
