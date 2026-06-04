Attribute VB_Name = "SnapLines"
Sub SnapLeadingLines(control As IRibbonControl)
    Dim oSlide As slide
    Dim bottomBar As Shape
    Dim shapeItem As Shape
    Dim i As Integer, j As Integer
    Dim middleY As Single, lineBottomY As Single

    ' Ensure a slide is selected
    If ActiveWindow.Selection.SlideRange.count = 0 Then
        MsgBox "Please select a slide first!"
        Exit Sub
    End If

    ' Set the current slide
    Set oSlide = ActiveWindow.Selection.SlideRange(1)

    ' Check for BottomBar
    If Not ShapeExists(oSlide, "BottomBar") Then
        MsgBox "BottomBar not found on the slide."
        Exit Sub
    End If
    Set bottomBar = oSlide.Shapes("BottomBar")

    ' Calculate the middle Y position of the BottomBar
    middleY = bottomBar.Top + (bottomBar.Height / 2)

    ' Loop through all shapes on the slide
    For i = 1 To oSlide.Shapes.count
        If oSlide.Shapes(i).Type = msoGroup Then
            ' Handle grouped shapes
            For j = 1 To oSlide.Shapes(i).GroupItems.count
                Set shapeItem = oSlide.Shapes(i).GroupItems.item(j)
                If Left(shapeItem.Name, 11) = "LeadingLine" Then
                    ' Calculate the bottom Y position of the line
                    lineBottomY = shapeItem.Top + shapeItem.Height
                    ' Move the top of the line to the middle Y
                    shapeItem.Top = middleY
                    ' Adjust the height to keep the bottom point stationary
                    shapeItem.Height = lineBottomY - middleY
                End If
            Next j
        ElseIf Left(oSlide.Shapes(i).Name, 11) = "LeadingLine" Then
            ' Handle ungrouped line shapes
            Set shapeItem = oSlide.Shapes(i)
            ' Calculate the bottom Y position of the line
            lineBottomY = shapeItem.Top + shapeItem.Height
            ' Move the top of the line to the middle Y
            shapeItem.Top = middleY
            ' Adjust the height to keep the bottom point stationary
            shapeItem.Height = lineBottomY - middleY
        End If
    Next i
    
    ' After adjusting lines, arrange entries in Z-order
    ArrangeEntriesInZOrder oSlide

    ' Bring timeline tables to the front
    BringTimelineTablesToFront oSlide
End Sub

Sub ArrangeEntriesInZOrder(oSlide As slide)
    Dim shapeItem As Shape
    For Each shapeItem In oSlide.Shapes
        If Left(shapeItem.Name, 8) = "DateBox" Or Left(shapeItem.Name, 9) = "EntryBox" Or Left(shapeItem.Name, 11) = "LeadingLine" Then
            shapeItem.ZOrder msoSendToBack
        End If
    Next shapeItem
End Sub

Sub BringTimelineTablesToFront(oSlide As slide)
    Dim shapeItem As Shape
    For Each shapeItem In oSlide.Shapes
        If shapeItem.Name = "TopBar" Or shapeItem.Name = "BottomBar" Then
            shapeItem.ZOrder msoBringToFront
        End If
    Next shapeItem
End Sub

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

