Attribute VB_Name = "ToggleSuffixes"
Sub ToggleDaySuffixes(control As IRibbonControl)
    Dim pptSlide As slide
    Dim pptTable As Table
    Dim i As Integer, dayNum As Integer
    Dim cellText As String, plainText As String, suffix As String
    Dim suffixOn As Boolean

    Set pptSlide = ActivePresentation.Slides(ActivePresentation.Slides.count)

    ' Check if BottomBar exists
    If ShapeExists(pptSlide, "BottomBar") Then
        Set pptTable = pptSlide.Shapes("BottomBar").Table

        ' Determine if the suffix is currently on or off
        cellText = pptTable.cell(1, 1).Shape.TextFrame.TextRange.text
        suffixOn = (Len(cellText) > 2) And Not IsNumeric(Right(cellText, 2))

        For i = 1 To pptTable.Columns.count
            cellText = pptTable.cell(1, i).Shape.TextFrame.TextRange.text
            If IsNumeric(Left(cellText, Len(cellText) - IIf(suffixOn, 2, 0))) Then
                dayNum = CInt(Left(cellText, Len(cellText) - IIf(suffixOn, 2, 0)))
                plainText = CStr(dayNum)
                suffix = GetDaySuffix(dayNum)

                If suffixOn Then
                    ' Remove suffix
                    pptTable.cell(1, i).Shape.TextFrame.TextRange.text = plainText
                Else
                    ' Add suffix
                    pptTable.cell(1, i).Shape.TextFrame.TextRange.text = plainText & suffix
                End If
            End If
        Next i
    Else
        MsgBox "BottomBar not found on the slide."
    End If
End Sub


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

