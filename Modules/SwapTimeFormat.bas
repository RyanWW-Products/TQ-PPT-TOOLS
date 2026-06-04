Attribute VB_Name = "SwapTimeFormat"
Sub SwapTimeFormats(control As IRibbonControl)
    Dim slide As slide
    Set slide = ActiveWindow.View.slide

    ProcessShapes slide.Shapes
End Sub

Sub ProcessShapes(shapesCollection As Object)
    Dim shp As Shape
    For Each shp In shapesCollection
        If shp.Type = msoGroup Then
            ' Recursively process grouped shapes
            ProcessShapes shp.GroupItems
        ElseIf shp.HasTextFrame Then
            shp.TextFrame.TextRange.text = SwapFormat(shp.TextFrame.TextRange.text)
        ElseIf shp.HasTable Then
            Dim tbl As Table
            Dim row As Integer, col As Integer
            Set tbl = shp.Table
            For row = 1 To tbl.Rows.count
                For col = 1 To tbl.Columns.count
                    tbl.cell(row, col).Shape.TextFrame.TextRange.text = SwapFormat(tbl.cell(row, col).Shape.TextFrame.TextRange.text)
                Next col
            Next row
        End If
    Next shp
End Sub

Function SwapFormat(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Pattern for Format A (e.g., 5:00 PM)
    regex.Pattern = "(\b\d{1,2}):(\d{2})\s?(AM|PM)\b"
    If regex.Test(text) Then
        SwapFormat = Format(TimeValue(text), "HH:MM")
    Else
        ' Pattern for Format B (e.g., 17:00)
        regex.Pattern = "\b(\d{1,2}):(\d{2})\b"
        If regex.Test(text) Then
            SwapFormat = Format(TimeValue(text), "hh:MM AM/PM")
        Else
            SwapFormat = text
        End If
    End If
End Function

