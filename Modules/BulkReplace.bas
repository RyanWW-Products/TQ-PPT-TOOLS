Attribute VB_Name = "BulkReplace"
' At the very top of your module
Public gOriginalShapes As Collection

Sub StoreOriginalShapes(control As IRibbonControl)
    ' Check that at least one shape is selected.
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes to be replaced first.", vbExclamation
        Exit Sub
    End If
    
    ' Initialize the global collection.
    Set gOriginalShapes = New Collection
    
    Dim i As Long
    Dim shp As Shape
    For i = 1 To ActiveWindow.Selection.ShapeRange.count
        Set shp = ActiveWindow.Selection.ShapeRange(i)
        gOriginalShapes.Add shp
    Next i
    
    MsgBox "Stored " & gOriginalShapes.count & " shape(s) to be replaced." & vbCrLf & _
           "Now, please select EXACTLY ONE replacement shape on the slide, then run the 'ReplaceStoredShapes' macro.", vbInformation
End Sub


Sub ReplaceStoredShapes(control As IRibbonControl)
    ' Ensure that original shapes have been stored.
    If gOriginalShapes Is Nothing Then
        MsgBox "No original shapes stored. Please run the 'StoreOriginalShapes' macro first.", vbExclamation
        Exit Sub
    End If

    ' Check that exactly one shape is selected as the replacement.
    If ActiveWindow.Selection.Type <> ppSelectionShapes Or ActiveWindow.Selection.ShapeRange.count <> 1 Then
        MsgBox "Please select EXACTLY ONE shape as the replacement and then run this macro.", vbExclamation
        Exit Sub
    End If
    
    Dim replacementShape As Shape
    Set replacementShape = ActiveWindow.Selection.ShapeRange(1)
    
    Dim sld As slide
    Set sld = ActiveWindow.View.slide
    
    Dim origShape As Shape, newShape As Shape
    Dim i As Long
    
    For i = 1 To gOriginalShapes.count
        Set origShape = gOriginalShapes(i)
        
        ' Duplicate the replacement shape.
        ' Duplicate returns a ShapeRange; we use the first (and only) shape from it.
        Set newShape = replacementShape.Duplicate()(1)
        
        With newShape
            ' Position and size the new shape to match the original shape.
            .Left = origShape.Left
            .Top = origShape.Top
            .Width = origShape.Width
            .Height = origShape.Height
            .Rotation = origShape.Rotation
            ' You can copy additional properties here if needed.
        End With
        
        ' Delete the original shape.
        origShape.Delete
    Next i
    
    ' Optionally, delete the replacement shape since it served as the "master."
    replacementShape.Delete
    
    ' Clear the stored shapes.
    Set gOriginalShapes = Nothing
    
    MsgBox "Replacement complete!", vbInformation
End Sub

