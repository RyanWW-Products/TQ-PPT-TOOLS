Attribute VB_Name = "TimelineEntry"
Sub EntryGenerator(control As IRibbonControl)
    Static lineCounter As Integer
    Dim oSlide As slide
    Dim oDateBox As Shape
    Dim oEntryBox As Shape
    Dim oGroup As Shape
    Dim oLine As Shape
    Dim lineName As String

    ' Ensure that a slide is selected.
    If ActiveWindow.Selection.SlideRange.count = 0 Then
        MsgBox "Please select a slide first!", vbExclamation
        Exit Sub
    End If

    ' Increment line counter for unique naming.
    lineCounter = lineCounter + 1

    ' Set the current slide.
    Set oSlide = ActiveWindow.Selection.SlideRange(1)

    ' Create the Date box.
    Set oDateBox = oSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                            Left:=100, Top:=100, Width:=2 * 72, Height:=50)
    With oDateBox
        .Fill.ForeColor.RGB = RGB(4, 61, 102)
        .line.Weight = 1.5
        .line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.text = "12/12/2099"
        .TextFrame.TextRange.Font.Name = "Arial"
        .TextFrame.TextRange.Font.Size = 12
        .TextFrame.TextRange.Font.Bold = msoTrue
        .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
        .TextFrame.MarginTop = 3
        .TextFrame.MarginBottom = 3
        .TextFrame.AutoSize = ppAutoSizeShapeToFitText
        .TextFrame.VerticalAnchor = msoAnchorBottom
    End With
    ' Auto-assign "Date Box" style.
    oDateBox.Tags.Add "GroupStyle", "Date Box"

    ' Create the Entry box.
    Set oEntryBox = oSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                             Left:=oDateBox.Left, Top:=oDateBox.Top + oDateBox.Height, _
                                             Width:=2 * 72, Height:=50)
    With oEntryBox
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .line.Weight = 1.5
        .line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.text = "Lorem Ipsum"
        .TextFrame.TextRange.Font.Name = "Arial"
        .TextFrame.TextRange.Font.Size = 12
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
        .TextFrame.MarginTop = 3
        .TextFrame.MarginBottom = 3
        .TextFrame.AutoSize = ppAutoSizeShapeToFitText
    End With
    ' Auto-assign "Entry Box" style.
    oEntryBox.Tags.Add "GroupStyle", "Entry Box"

    ' Group the two boxes.
    Set oGroup = oSlide.Shapes.range(Array(oDateBox.Name, oEntryBox.Name)).Group

    ' Add drop shadow to the group.
    With oGroup.Shadow
        .Type = msoShadow21
        .IncrementOffsetX 3
        .IncrementOffsetY 3
    End With

    ' Create and name the vertical line segment.
    lineName = "LeadingLine" & Format(lineCounter, "00")
    Set oLine = oSlide.Shapes.AddLine(oGroup.Left + oGroup.Width / 2, oGroup.Top - 50, _
                                      oGroup.Left + oGroup.Width / 2, oGroup.Top + oGroup.Height - oDateBox.Height)
    oLine.Name = lineName
    With oLine
        .line.Weight = 1.5
        .line.ForeColor.RGB = RGB(0, 0, 0)
        .ZOrder msoSendToBack
    End With

    ' Group everything.
    oSlide.Shapes.range(Array(oGroup.Name, oLine.Name)).Group
End Sub


