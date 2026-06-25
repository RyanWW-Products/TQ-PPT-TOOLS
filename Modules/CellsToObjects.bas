Attribute VB_Name = "CellsToObjects"
' Ribbon: convert the selected table into individual text boxes.
Sub ConvertSelectedTableToTextBoxes(control As IRibbonControl)
    Dim sel As Object
    Set sel = ActiveWindow.Selection
    If Not sel.Type = ppSelectionShapes Then
        MsgBox "No shape selected. Please select a table."
        Exit Sub
    End If
    If sel.ShapeRange.count <> 1 Or Not sel.ShapeRange(1).HasTable Then
        MsgBox "Selected shape is not a table. Please select a table."
        Exit Sub
    End If
    ConvertTableToTextBoxes sel.ShapeRange(1), 0
End Sub

' Convert a table shape into one text box per cell (preserving geometry / text / format),
' delete the table, and return the new text boxes as a Collection. If fixedRowHeight > 0,
' every row is set to that height first so a converted datebar has a consistent vertical size.
Public Function ConvertTableToTextBoxes(ByVal tableShape As Shape, Optional ByVal fixedRowHeight As Single = 0, Optional ByVal cellDates As Variant) As Collection
    Dim pptSlide As slide, pptTable As Table
    Dim i As Integer, j As Integer
    Dim newTextBox As Shape, cellText As String
    Dim cellLeft As Single, cellTop As Single, cellWidth As Single, cellHeight As Single
    Dim made As Collection, kk As Long, hasDates As Boolean
    Set made = New Collection
    hasDates = Not IsMissing(cellDates)

    Set pptSlide = tableShape.parent
    Set pptTable = tableShape.Table

    If fixedRowHeight > 0 Then
        For i = 1 To pptTable.Rows.count
            pptTable.Rows(i).Height = fixedRowHeight
        Next i
    End If

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
                .VerticalAnchor = msoAnchorMiddle
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
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 1
                .Style = msoLineSingle
            End With

            kk = kk + 1
            If hasDates Then
                On Error Resume Next
                newTextBox.Tags.Add "TLCellDate", CStr(CDbl(cellDates(kk)))
                On Error GoTo 0
            End If
            made.Add newTextBox
        Next j
    Next i

    tableShape.Delete
    Set ConvertTableToTextBoxes = made
End Function
