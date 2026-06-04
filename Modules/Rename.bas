Attribute VB_Name = "Rename"
Sub BulkRenameSelectedObjects(control As IRibbonControl)
    Dim shp As Shape
    Dim baseName As String
    Dim SlideIndex As Integer
    Dim count As Integer
    Dim numFormat As String
    Dim pres As presentation
    Dim sel As Selection
    
    ' Set up references
    Set pres = ActivePresentation
    Set sel = ActiveWindow.Selection
    
    ' Ensure something is selected
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more objects to rename.", vbExclamation, "No Shapes Selected"
        Exit Sub
    End If
    
    ' Get base name from user
    baseName = InputBox("Enter the base name for selected objects:", "Bulk Rename")
    
    ' Ensure the user entered something
    If Trim(baseName) = "" Then
        MsgBox "Invalid name. Operation cancelled.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Start numbering at 1
    count = 1
    numFormat = "000" ' Ensures numbering format like 001, 002, 003
    
    ' Loop through each selected shape and rename it
    For Each shp In sel.ShapeRange
        shp.Name = baseName & " " & Format(count, numFormat)
        count = count + 1
    Next shp
    
    ' Confirmation message
    MsgBox "Renamed " & count - 1 & " objects successfully.", vbInformation, "Rename Complete"
End Sub

