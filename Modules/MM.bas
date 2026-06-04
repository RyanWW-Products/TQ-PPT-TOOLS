Attribute VB_Name = "MM"
Sub MightyMorpher(control As IRibbonControl)
    Dim shp As Shape
    
    'Loop through all selected shapes on the current slide
    For Each shp In ActiveWindow.Selection.ShapeRange
        
        'Check if the shape has a valid name for Morph transition
        If Not shp.Name Like "!!* " Then
            'If the shape's name doesn't start with "!!! ", rename it to make it valid for Morph transition
            shp.Name = "!!!" & shp.Name
        End If
    Next shp
End Sub
