Attribute VB_Name = "CalendarGenerator"
Function IsValidDate(DateString As String) As Boolean
    Dim DateParts As Variant
    If InStr(DateString, "/") > 0 Then
        DateParts = Split(DateString, "/")
        If UBound(DateParts) = 1 Then
            If IsNumeric(DateParts(0)) And IsNumeric(DateParts(1)) Then
                If Len(DateParts(0)) <= 2 And Len(DateParts(1)) = 4 Then
                    If DateParts(0) >= 1 And DateParts(0) <= 12 Then
                        IsValidDate = True
                    End If
                End If
            End If
        End If
    End If
End Function

Sub CalendarGenerator(control As IRibbonControl)
    Dim ppt As PowerPoint.Application
    Dim presentation As PowerPoint.presentation
    Dim slide As PowerPoint.slide
    Dim calendarTable As PowerPoint.Shape
    Dim i As Integer
    Dim startDay As Integer
    Dim daysInMonth As Integer
    Dim inputDate As Date
    Dim inputString As String
    Dim templatePath As String

    templatePath = Environ("APPDATA") & "\Microsoft\AddIns\Trial Ex Addin\CalendarTemplate.pptx"
    
    ' Get the PowerPoint application, presentation, and slide
    Set ppt = PowerPoint.Application
    Set presentation = ppt.ActivePresentation
    Set slide = presentation.Slides.Add(presentation.Slides.count + 1, ppLayoutBlank)
    
    ' Prompt the user for the month and year
    Do
        inputString = InputBox("Please enter the month and year (MM/YYYY):", "Input Date")
        If IsValidDate(inputString) Then Exit Do
        MsgBox "Invalid date. Please enter the date in the format MM/YYYY.", vbOKOnly + vbExclamation, "Invalid Input"
    Loop
    
    inputDate = DateSerial(Right(inputString, 4), Left(inputString, 2), 1)
    
    ' Open the template presentation and copy the table
    Dim templatePresentation As PowerPoint.presentation
    Set templatePresentation = ppt.Presentations.Open(templatePath, WithWindow:=msoFalse)
    templatePresentation.Slides(1).Shapes(1).Copy
    templatePresentation.Close
    
' Paste the table to the new slide
slide.Shapes.Paste
Set calendarTable = slide.Shapes(slide.Shapes.count)
calendarTable.Name = "CalendarTable"
    
    ' Get the start day of the week and the total number of days in the month
    startDay = Weekday(inputDate, vbSunday)
    daysInMonth = day(DateAdd("m", 1, inputDate) - 1)
    
    ' Check if the 6th row is needed, and if so, add it
    If daysInMonth + startDay - 1 > 36 Then
        calendarTable.Table.Rows.Add
    End If
    
    ' Set the calendar days
    For i = 1 To daysInMonth
        Dim row As Integer
        Dim column As Integer
        row = ((startDay - 2 + i) \ 7) + 1
        column = ((startDay - 2 + i) Mod 7) + 1

        With calendarTable.Table.cell(row, column).Shape.TextFrame.TextRange
            .text = i
            .ParagraphFormat.Alignment = ppAlignLeft
            .Paragraphs(1).ParagraphFormat.Alignment = ppAlignLeft
            .Font.Size = 14
            .parent.VerticalAnchor = msoAnchorTop
        End With
    Next i

    ' Add date as a custom property to the slide
    AddCustomProperty slide, "CalendarMonth", Left(inputString, 2)
    AddCustomProperty slide, "CalendarYear", Right(inputString, 4)
End Sub

Sub AddCustomProperty(slide As PowerPoint.slide, propName As String, propValue As String)
    With slide.Tags
        .Add propName, propValue
    End With
End Sub

Function GetCustomProperty(slide As PowerPoint.slide, propName As String) As String
    On Error Resume Next
    GetCustomProperty = slide.Tags(propName)
    On Error GoTo 0
End Function




