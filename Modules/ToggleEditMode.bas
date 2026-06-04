Attribute VB_Name = "ToggleEditMode"
Sub ToggleTableLayer(control As IRibbonControl)
    Dim slide As PowerPoint.slide
    Set slide = ActiveWindow.View.slide


    Dim calendar As PowerPoint.Shape
    Dim foundTable As Boolean
    foundTable = False

    ' Debugging: Count total shapes on the slide
    Debug.Print "Total shapes on slide: " & slide.Shapes.count

    ' Find the first (and assumed only) table on the slide
    Dim i As Integer
    For i = 1 To slide.Shapes.count
        Set calendar = slide.Shapes(i)
        If calendar.HasTable Then
            foundTable = True
            ' Debugging: Output table information
            Debug.Print "Table found: Shape #" & i & "; Name - " & calendar.Name
            Exit For
        End If
    Next i

    If Not foundTable Then
        MsgBox "No table found on the slide."
        Exit Sub
    End If

    ' Debugging: Check current z-order position
    Debug.Print "Current Z-Order Position of Table: " & calendar.ZOrderPosition

    ' Toggle the table layer
    If calendar.ZOrderPosition = 1 Then
        ' If table is already at the back, bring it to front
        calendar.ZOrder msoBringToFront
        Debug.Print "Table moved to Front"
    Else
        ' Send table to back
        calendar.ZOrder msoSendToBack
        Debug.Print "Table moved to Back"
    End If
End Sub

