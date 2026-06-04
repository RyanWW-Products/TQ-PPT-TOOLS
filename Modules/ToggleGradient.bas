Attribute VB_Name = "ToggleGradient"
Sub ToggleGradientFill(control As IRibbonControl)
    Dim SelectedShape As Shape
    Dim objectColor As Long, darkerColor As Long

    ' Ensure correct object is selected
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set SelectedShape = ActiveWindow.Selection.ShapeRange(1)
        
        ' Determine the type of fill and toggle accordingly
        If SelectedShape.Fill.Type = msoFillGradient Then
            ' Set to solid fill using the first gradient stop's color
            objectColor = SelectedShape.Fill.GradientStops(1).Color.RGB
            SelectedShape.Fill.Solid
            SelectedShape.Fill.ForeColor.RGB = objectColor
        ElseIf SelectedShape.Fill.Type = msoFillSolid Then
            ' Get the RGB value of the fill color
            objectColor = SelectedShape.Fill.ForeColor.RGB
            
            ' Darken the color
            darkerColor = DarkenColor(objectColor)
            
            ' Change to vertical gradient
            With SelectedShape.Fill
                .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
                .GradientStops(1).Color.RGB = objectColor
                .GradientStops(2).Color.RGB = darkerColor
            End With
        Else
            MsgBox "The selected shape's fill is neither solid nor gradient. No changes made."
        End If
    Else
        MsgBox "No shape is selected. Please select a shape."
    End If
End Sub

Function DarkenColor(origColor As Long) As Long
    Dim r As Integer, g As Integer, b As Integer
    Dim darkFactor As Single

    darkFactor = 0.6 ' Adjust darkness factor here

    ' Extract RGB components from the original color
    r = origColor Mod 256
    g = (origColor \ 256) Mod 256
    b = (origColor \ 65536) Mod 256

    ' Reduce each component to make it darker
    r = r * darkFactor
    g = g * darkFactor
    b = b * darkFactor

    ' Combine the components back into a single RGB value
    DarkenColor = RGB(r, g, b)
End Function

