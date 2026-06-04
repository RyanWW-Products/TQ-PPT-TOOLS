Attribute VB_Name = "ObjectSearch"
Sub SelectObjectsByName(control As IRibbonControl)
    Dim shp As Shape
    Dim sld As slide
    Dim searchTerm As String
    Dim matches As New Collection
    Dim i As Long
    
    ' Get the active slide
    Set sld = ActiveWindow.View.slide

    ' Ask the user for a search term
    searchTerm = InputBox("Enter part of the name to search for:", "Select Objects by Name")
    
    ' Ensure input is valid
    If Trim(searchTerm) = "" Then
        MsgBox "Invalid search term. Operation cancelled.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Loop through each shape in the slide
    For Each shp In sld.Shapes
        ' Check if shape name contains the search term (case-insensitive)
        If InStr(1, shp.Name, searchTerm, vbTextCompare) > 0 Then
            matches.Add shp.Name  ' Collect the shape name
        End If
    Next shp
    
    ' Now select all matches at once
    If matches.count > 0 Then
        Dim shapeNames() As Variant
        ReDim shapeNames(1 To matches.count)
        
        ' Fill the array with matching shape names
        For i = 1 To matches.count
            shapeNames(i) = matches(i)
        Next i
        
        ' Select them in one go using a ShapeRange
        sld.Shapes.range(shapeNames).Select
        
    Else
        ' If no matches found, notify the user
        MsgBox "No objects found with the term '" & searchTerm & "'.", vbInformation, "No Matches"
    End If
End Sub

