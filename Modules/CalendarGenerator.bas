Attribute VB_Name = "CalendarGenerator"
Option Explicit

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
    Dim pres As PowerPoint.presentation
    Dim tpl As PowerPoint.presentation
    Dim templatePath As String
    Dim choice As VbMsgBoxResult
    Dim startDate As Date, endDate As Date
    Dim cancelled As Boolean
    Dim d As Date, cnt As Integer

    Set ppt = PowerPoint.Application
    If ppt.Presentations.count = 0 Then
        MsgBox "Please open a presentation first.", vbExclamation, "Calendar Generator"
        Exit Sub
    End If
    Set pres = ppt.ActivePresentation

    ' Single calendar, or a range (each month on its own slide)?
    choice = MsgBox("Create a RANGE of calendars (each month on its own slide)?" & vbCrLf & vbCrLf & _
                  "Yes = a range (enter a start and end month)" & vbCrLf & _
                  "No  = a single month", _
                  vbYesNoCancel + vbQuestion, "Calendar Generator")
    If choice = vbCancel Then Exit Sub

    If choice = vbYes Then
        startDate = PromptMonthYear("Enter the START month and year (MM/YYYY):", cancelled)
        If cancelled Then Exit Sub
        endDate = PromptMonthYear("Enter the END month and year (MM/YYYY):", cancelled)
        If cancelled Then Exit Sub
        If endDate < startDate Then        ' tolerate reversed entry
            Dim tmp As Date
            tmp = startDate: startDate = endDate: endDate = tmp
        End If
        Dim nMonths As Integer
        nMonths = DateDiff("m", startDate, endDate) + 1
        If MsgBox("This will create " & nMonths & " calendars, each on its own slide." & vbCrLf & _
                  "Continue?", vbYesNo + vbQuestion, "Calendar Generator") <> vbYes Then Exit Sub
    Else
        startDate = PromptMonthYear("Please enter the month and year (MM/YYYY):", cancelled)
        If cancelled Then Exit Sub
        endDate = startDate
    End If

    ' Open the template ONCE for the whole run
    templatePath = Environ("APPDATA") & "\Microsoft\AddIns\Trial Ex Addin\CalendarTemplate.pptx"
    If Dir(templatePath) = "" Then
        MsgBox "Calendar template not found:" & vbCrLf & vbCrLf & templatePath, vbExclamation, "Calendar Generator"
        Exit Sub
    End If

    On Error GoTo CleanFail
    Set tpl = ppt.Presentations.Open(templatePath, WithWindow:=msoFalse)

    d = startDate
    Do While d <= endDate
        CreateCalendarSlide pres, tpl, d
        cnt = cnt + 1
        d = DateAdd("m", 1, d)
    Loop

    tpl.Close
    Set tpl = Nothing
    If cnt > 1 Then MsgBox cnt & " calendars created.", vbInformation, "Calendar Generator"
    Exit Sub
CleanFail:
    On Error Resume Next
    If Not tpl Is Nothing Then tpl.Close
    MsgBox "Calendar creation failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Calendar Generator"
End Sub

' Prompt for MM/YYYY; returns the 1st of that month. Sets cancelled=True if the
' user cancels the InputBox.
Private Function PromptMonthYear(prompt As String, ByRef cancelled As Boolean) As Date
    Dim s As String
    cancelled = False
    Do
        s = InputBox(prompt, "Input Date")
        If s = "" Then cancelled = True: Exit Function
        If IsValidDate(s) Then
            PromptMonthYear = DateSerial(CInt(Right(s, 4)), CInt(Left(s, 2)), 1)
            Exit Function
        End If
        MsgBox "Invalid date. Please enter the date in the format MM/YYYY.", vbOKOnly + vbExclamation, "Invalid Input"
    Loop
End Function

' Build one calendar slide for the given month, copying the table from the
' already-open template.
Private Sub CreateCalendarSlide(pres As PowerPoint.presentation, tpl As PowerPoint.presentation, monthDate As Date)
    Dim slide As PowerPoint.slide
    Dim calendarTable As PowerPoint.Shape
    Dim startDay As Integer, daysInMonth As Integer, i As Integer
    Dim row As Integer, column As Integer

    Set slide = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)

    tpl.Slides(1).Shapes(1).Copy
    slide.Shapes.Paste
    Set calendarTable = slide.Shapes(slide.Shapes.count)
    calendarTable.Name = "CalendarTable"

    startDay = Weekday(monthDate, vbSunday)
    daysInMonth = day(DateAdd("m", 1, monthDate) - 1)

    ' Add the 6th row only when the month needs it
    If daysInMonth + startDay - 1 > 36 Then calendarTable.Table.Rows.Add

    For i = 1 To daysInMonth
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

    AddCustomProperty slide, "CalendarMonth", CStr(month(monthDate))
    AddCustomProperty slide, "CalendarYear", CStr(year(monthDate))
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
