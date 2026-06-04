Attribute VB_Name = "SwapMonthFormat"
Sub SwapMonthFormats(control As IRibbonControl)
    Dim slide As slide
    Set slide = ActiveWindow.View.slide

    ProcessMonthShapes slide.Shapes
End Sub

Sub ProcessMonthShapes(shapesCollection As Object)
    Dim shp As Shape
    For Each shp In shapesCollection
        If shp.Type = msoGroup Then
            ' Recursively process grouped shapes
            ProcessMonthShapes shp.GroupItems
        ElseIf shp.HasTextFrame Then
            shp.TextFrame.TextRange.text = SwapDateFormat(shp.TextFrame.TextRange.text)
        ElseIf shp.HasTable Then
            Dim tbl As Table
            Dim row As Integer, col As Integer
            Set tbl = shp.Table
            For row = 1 To tbl.Rows.count
                For col = 1 To tbl.Columns.count
                    tbl.cell(row, col).Shape.TextFrame.TextRange.text = SwapDateFormat(tbl.cell(row, col).Shape.TextFrame.TextRange.text)
                Next col
            Next row
        End If
    Next shp
End Sub

Function SwapDateFormat(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Pattern for Format A (e.g., Oct 2000)
    regex.Pattern = "\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s(\d{4})\b"
    If regex.Test(text) Then
        SwapDateFormat = Replace(text, " ", ". ") ' Convert to Format B
    Else
        ' Pattern for Format B (e.g., Oct. 2000)
        regex.Pattern = "\b(Jan\.|Feb\.|Mar\.|Apr\.|May|Jun\.|Jul\.|Aug\.|Sep\.|Oct\.|Nov\.|Dec\.)\s(\d{4})\b"
        If regex.Test(text) Then
            SwapDateFormat = ConvertToFormatC(text) ' Convert to Format C
        Else
            ' Assume Format C and convert to Format A
            SwapDateFormat = ConvertToFormatA(text)
        End If
    End If
End Function

Function ConvertToFormatC(text As String) As String
    Dim monthNames As Object
    Set monthNames = CreateObject("Scripting.Dictionary")
    monthNames.Add "Jan.", "January"
    monthNames.Add "Feb.", "February"
    monthNames.Add "Mar.", "March"
    monthNames.Add "Apr.", "April"
    monthNames.Add "May.", "May"
    monthNames.Add "Jun.", "June"
    monthNames.Add "Jul.", "July"
    monthNames.Add "Aug.", "August"
    monthNames.Add "Sep.", "September"
    monthNames.Add "Oct.", "October"
    monthNames.Add "Nov.", "November"
    monthNames.Add "Dec.", "December"

    Dim month As String
    month = Split(text, ". ")(0) & "."
    If monthNames.exists(month) Then
        ConvertToFormatC = Replace(text, month, monthNames.item(month)) & " "
    Else
        ConvertToFormatC = text
    End If
End Function

Function ConvertToFormatA(text As String) As String
    Dim monthNames As Object
    Set monthNames = CreateObject("Scripting.Dictionary")
    monthNames.Add "January", "Jan"
    monthNames.Add "February", "Feb"
    monthNames.Add "March", "Mar"
    monthNames.Add "April", "Apr"
    monthNames.Add "May", "May "
    monthNames.Add "June", "Jun"
    monthNames.Add "July", "Jul"
    monthNames.Add "August", "Aug"
    monthNames.Add "September", "Sep"
    monthNames.Add "October", "Oct"
    monthNames.Add "November", "Nov"
    monthNames.Add "December", "Dec"

    Dim month As String
    month = Split(text, " ")(0)
    If monthNames.exists(month) Then
        ConvertToFormatA = monthNames.item(month) & " " & Split(text, " ")(1)
    Else
        ConvertToFormatA = text
    End If
End Function

