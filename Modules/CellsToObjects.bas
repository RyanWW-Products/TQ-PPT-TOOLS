Attribute VB_Name = "CellsToObjects"
Sub ConvertSelectedTableToTextBoxes(control As IRibbonControl)
    Dim pptSlide As slide
    Dim pptTable As Table
    Dim sel As Object
    Dim i As Integer, j As Integer
    Dim newTextBox As Shape
    Dim cellText As String
    Dim cellLeft As Single, cellTop As Single, cellWidth As Single, cellHeight As Single

    ' Define a standard border color and weight
    Dim standardBorderColor As Long
    Dim standardBorderWeight As Single
    standardBorderColor = RGB(0, 0, 0) ' Black color
    standardBorderWeight = 1 ' 1 point

    Set sel = ActiveWindow.Selection
    If Not sel.Type = ppSelectionShapes Then
        MsgBox "No shape selected. Please select a table."
        Exit Sub
    End If

    If sel.ShapeRange.count <> 1 Or Not sel.ShapeRange(1).HasTable Then
        MsgBox "Selected shape is not a table. Please select a table."
        Exit Sub
    End If

    Set pptSlide = sel.ShapeRange(1).parent
    Set pptTable = sel.ShapeRange(1).Table

    ' Loop through each cell in the table
    For i = 1 To pptTable.Rows.count
        For j = 1 To pptTable.Columns.count
            With pptTable.cell(i, j).Shape
                cellLeft = .Left
                cellTop = .Top
                cellWidth = .Width
                cellHeight = .Height
                cellText = .TextFrame.TextRange.text
            End With

            Set newTextBox = pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, cellLeft, cellTop, cellWidth, cellHeight)
            
            With newTextBox.TextFrame
                .AutoSize = ppAutoSizeNone
                .parent.Height = cellHeight
                .VerticalAnchor = msoAnchorMiddle  ' Vertically center the text

                With .TextRange
                    .text = cellText
                    .Font.Name = pptTable.cell(i, j).Shape.TextFrame.TextRange.Font.Name
                    .Font.Size = pptTable.cell(i, j).Shape.TextFrame.TextRange.Font.Size
                    .Font.Bold = pptTable.cell(i, j).Shape.TextFrame.TextRange.Font.Bold
                    .Font.Italic = pptTable.cell(i, j).Shape.TextFrame.TextRange.Font.Italic
                    .Font.Underline = pptTable.cell(i, j).Shape.TextFrame.TextRange.Font.Underline
                    .Font.Color = pptTable.cell(i, j).Shape.TextFrame.TextRange.Font.Color
                    .ParagraphFormat.Alignment = pptTable.cell(i, j).Shape.TextFrame.TextRange.ParagraphFormat.Alignment
                End With
            End With

            If pptTable.cell(i, j).Shape.Fill.Type = msoFillGradient Then
                With newTextBox.Fill
                    .TwoColorGradient Style:=msoGradientVertical, Variant:=1
                    .GradientAngle = 90
                    .GradientStops(1).Color.RGB = pptTable.cell(i, j).Shape.Fill.GradientStops(1).Color.RGB
                    .GradientStops(2).Color.RGB = pptTable.cell(i, j).Shape.Fill.GradientStops(2).Color.RGB
                End With
            Else
                newTextBox.Fill.ForeColor.RGB = pptTable.cell(i, j).Shape.Fill.ForeColor.RGB
            End If

            With newTextBox.line
                .ForeColor.RGB = standardBorderColor
                .Weight = standardBorderWeight
                .Style = msoLineSingle
            End With
        Next j
    Next i

    pptTable.parent.Delete
End Sub

