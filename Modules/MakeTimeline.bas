Attribute VB_Name = "MakeTimeline"
Option Explicit

' Fixed datebar band height (matches the Timeline maker's BAND_HEIGHT) so every datebar
' group has the same vertical size regardless of column count / font scaling.
Public Const BAR_HEIGHT As Single = 30

Sub CreateTimeline(control As IRibbonControl)
    Dim timelineType As String
    Dim timelineColor As String
    Dim startDate As Date, endDate As Date
    Dim isTwoBar As Boolean
    Dim SlideIndex As Integer
    Dim cancelFlag As Boolean ' Flag to check if user cancelled the input

    ' Create and show the UserForm
    Dim frm As TimelineSettings
    Set frm = New TimelineSettings
    frm.Caption = "DateBar Settings"      ' renamed at runtime (no form re-import needed)
    frm.Show

    If frm.Tag = "OK" Then
        ' Retrieve the values from the UserForm
        timelineType = frm.timelineType
        timelineColor = frm.timelineColor
        ' ... (Continue with the rest of your UserForm handling code)
    Else
        ' User clicked Cancel or closed the form
        Exit Sub
    End If

    Unload frm

    ' Get the start and end times or dates based on the timeline type
    Select Case timelineType
        Case "Hours"
            startDate = InputHour("Enter the start hour (HH):", cancelFlag)
            If cancelFlag Then Exit Sub

            endDate = InputHour("Enter the end hour (HH):", cancelFlag)
            If cancelFlag Then Exit Sub

            If endDate < startDate Then endDate = DateAdd("d", 1, endDate)

        Case "Days"
            startDate = inputDate("Enter the start date (MM/DD/YYYY):", cancelFlag)
            If cancelFlag Then Exit Sub

            endDate = inputDate("Enter the end date (MM/DD/YYYY):", cancelFlag)
            If cancelFlag Then Exit Sub

            isTwoBar = MsgBox("Use a two-bar layout for days?", vbYesNo) = vbYes

        Case "Months"
            startDate = InputMonthYear("Enter the start month/year (MM/YYYY):", cancelFlag)
            If cancelFlag Then Exit Sub

            endDate = InputMonthYear("Enter the end month/year (MM/YYYY):", cancelFlag)
            If cancelFlag Then Exit Sub

            isTwoBar = MsgBox("Use a two-bar layout for months?", vbYesNo) = vbYes

        Case "Years"
            startDate = InputYear("Enter the start year (YYYY):", cancelFlag)
            If cancelFlag Then Exit Sub

            endDate = InputYear("Enter the end year (YYYY):", cancelFlag)
            If cancelFlag Then Exit Sub

    End Select

    ' Draw the datebar on the CURRENTLY ACTIVE slide (previously this added a new slide)
    If ActivePresentation.Slides.count = 0 Then
        SlideIndex = ActivePresentation.Slides.Add(1, ppLayoutBlank).SlideIndex
    ElseIf Not (ActiveWindow.View.slide Is Nothing) Then
        SlideIndex = ActiveWindow.View.slide.SlideIndex
    Else
        SlideIndex = 1
    End If

    ' Creating the datebar on the active slide
    If isTwoBar And (timelineType = "Days" Or timelineType = "Months") Then
        CreateTwoBarTimelineTable SlideIndex, startDate, endDate, timelineColor, timelineType
    Else
        CreateTimelineTable SlideIndex, timelineType, startDate, endDate, timelineColor
    End If

    MsgBox "DateBar created."
End Sub

Function IsValidTimelineType(timelineType As String) As Boolean
    IsValidTimelineType = (timelineType = "Hours" Or timelineType = "Days" Or timelineType = "Months" Or timelineType = "Years")
    If Not IsValidTimelineType Then
        MsgBox "Invalid date-bar type. Please enter one of the following: Hours, Days, Months, Years."
    End If
End Function
Function inputDate(prompt As String, ByRef wasCancelled As Boolean) As Date
    Dim userInput As String
    wasCancelled = False ' Initialize flag

    Do
        userInput = InputBox(prompt, "Input Date")
        If userInput = "" Then ' User pressed Cancel
            wasCancelled = True
            Exit Function
        ElseIf IsDate(userInput) Then
            inputDate = CDate(userInput)
            Exit Function
        Else
            MsgBox "Invalid input. Please enter a valid date (MM/DD/YYYY)."
        End If
    Loop
End Function
Function InputHour(prompt As String, ByRef wasCancelled As Boolean) As Date
    Dim userInput As String
    Dim hour As Integer
    wasCancelled = False ' Initialize flag

    Do
        userInput = InputBox(prompt, "Input Hour")
        If userInput = "" Then ' User pressed Cancel
            wasCancelled = True
            Exit Function
        ElseIf IsNumeric(userInput) Then
            hour = CInt(userInput)
            If hour >= 0 And hour < 24 Then
                InputHour = DateSerial(1, 1, 1) + TimeSerial(hour, 0, 0)
                Exit Function
            End If
        End If
        MsgBox "Invalid input. Please enter a valid hour (HH)."
    Loop
End Function
Function InputMonthYear(prompt As String, ByRef wasCancelled As Boolean) As Date
    Dim userInput As String
    Dim monthYear() As String
    Dim month As Integer
    Dim year As Integer
    wasCancelled = False ' Initialize flag

    Do
        userInput = InputBox(prompt, "Input Month/Year")
        If userInput = "" Then ' User pressed Cancel
            wasCancelled = True
            Exit Function
        ElseIf UBound(Split(userInput, "/")) = 1 Then
            monthYear = Split(userInput, "/")
            month = CInt(monthYear(0))
            year = CInt(monthYear(1))
            If month >= 1 And month <= 12 Then
                InputMonthYear = DateSerial(year, month, 1)
                Exit Function
            End If
        End If
        MsgBox "Invalid input. Please enter a valid month/year (MM/YYYY)."
    Loop
End Function
Function InputYear(prompt As String, ByRef wasCancelled As Boolean) As Date
    Dim userInput As String
    wasCancelled = False ' Initialize flag

    Do
        userInput = InputBox(prompt, "Input Year")
        If userInput = "" Then ' User pressed Cancel
            wasCancelled = True
            Exit Function
        ElseIf IsNumeric(userInput) And Len(userInput) = 4 Then
            InputYear = DateSerial(CInt(userInput), 1, 1)
            Exit Function
        Else
            MsgBox "Invalid input. Please enter a valid year (YYYY)."
        End If
    Loop
End Function
Function GetSegmentCount(timelineType As String, startDate As Date, endDate As Date) As Integer
    Select Case timelineType
        Case "Hours"
            GetSegmentCount = DateDiff("h", startDate, endDate)
        Case "Days"
            GetSegmentCount = DateDiff("d", startDate, endDate) + 1
        Case "Months"
            GetSegmentCount = DateDiff("m", startDate, endDate) + 1
        Case "Years"
            GetSegmentCount = DateDiff("yyyy", startDate, endDate) + 1
    End Select
End Function
Function GetSegmentLabel(timelineType As String, segmentDate As Date) As String
    Select Case timelineType
        Case "Hours"
            GetSegmentLabel = Format(segmentDate, "h:mm AM/PM")
        Case "Days"
            GetSegmentLabel = Format(segmentDate, "mm/dd/yyyy")
        Case "Months"
            GetSegmentLabel = Format(segmentDate, "mmm yyyy")
        Case "Years"
            GetSegmentLabel = Format(segmentDate, "yyyy")
    End Select
End Function
Sub CreateTimelineTable(SlideIndex As Integer, timelineType As String, startDate As Date, endDate As Date, timelineColor As String)
    Dim pptSlide As slide
    Dim pptTable As Table
    Dim colCount As Integer
    Dim i As Integer
    Dim slideWidth As Single, tableWidth As Single
    Dim segmentDate As Date, cellDates() As Date

    Set pptSlide = ActivePresentation.Slides(SlideIndex)
slideWidth = pptSlide.parent.PageSetup.slideWidth
colCount = GetSegmentCount(timelineType, startDate, endDate)
ReDim cellDates(1 To colCount)
Set pptTable = pptSlide.Shapes.AddTable(1, colCount).Table

tableWidth = slideWidth
segmentDate = startDate
For i = 1 To colCount
    pptTable.cell(1, i).Shape.TextFrame.TextRange.text = GetSegmentLabel(timelineType, segmentDate)
    pptTable.Columns(i).Width = tableWidth / colCount
    FormatCell pptTable.cell(1, i), timelineColor, False, colCount ' Corrected call to FormatCell
    cellDates(i) = segmentDate
    segmentDate = DateAdd(GetIntervalType(timelineType), 1, segmentDate)
Next i
    pptTable.parent.Left = (slideWidth - tableWidth) / 2
    pptTable.parent.Top = pptSlide.parent.PageSetup.slideHeight / 3

    FormatTableBorders pptTable, RGB(0, 0, 0)
    FinalizeDateBar pptSlide, Array(pptTable.parent), timelineType, Array(cellDates)
End Sub

' Convert the bar table(s) to objects (fixed BAR_HEIGHT), group them into one "Datebar NNN",
' and stamp the unit so Date Snap etc. recognize it. tables = array of table shapes.
Private Function FinalizeDateBar(ByVal sld As slide, ByVal tables As Variant, ByVal unitType As String, Optional ByVal tablesDates As Variant) As Shape
    Dim ti As Long, coll As Collection, s As Variant, tShape As Shape
    Dim names As Collection, arr() As String, i As Long, grp As Shape, hasDates As Boolean
    hasDates = Not IsMissing(tablesDates)
    Set names = New Collection
    For ti = LBound(tables) To UBound(tables)
        Set tShape = tables(ti)
        If hasDates Then
            Set coll = ConvertTableToTextBoxes(tShape, BAR_HEIGHT, tablesDates(ti))
        Else
            Set coll = ConvertTableToTextBoxes(tShape, BAR_HEIGHT)
        End If
        For Each s In coll
            names.Add s.Name
        Next s
    Next ti
    If names.count = 0 Then Exit Function
    ReDim arr(1 To names.count)
    For i = 1 To names.count
        arr(i) = names(i)
    Next i
    If names.count = 1 Then
        Set grp = sld.Shapes(arr(1))
    Else
        Set grp = sld.Shapes.Range(arr).Group
    End If
    grp.Name = NextDateBarName(sld)
    grp.Tags.Add "TLBar", "1"
    grp.Tags.Add "TLType", unitType
    With grp.Shadow
        .Visible = msoTrue
        .Type = msoShadow21
        .IncrementOffsetX 3
        .IncrementOffsetY 3
    End With
    Set FinalizeDateBar = grp
End Function
Function GetIntervalType(timelineType As String) As String
    Select Case timelineType
        Case "Hours": GetIntervalType = "h"
        Case "Days": GetIntervalType = "d"
        Case "Months": GetIntervalType = "m"
        Case "Years": GetIntervalType = "yyyy"
    End Select
End Function
    Sub FormatCell(cell As cell, colorName As String, isBottomBar As Boolean, colCount As Integer)
    Dim baseFontSize As Integer
    baseFontSize = 14 ' Base font size

    Dim topColor As Long, bottomColor As Long

    ' Determine the RGB color based on the color name
    If isBottomBar Then
        ' Adjust these colors based on your requirement
        Select Case colorName
            Case "DarkerGreen"
                topColor = DarkenColor(RGB(0, 176, 80), 0.4)
                bottomColor = DarkenColor(RGB(0, 70, 32), 0.4)
            Case "DarkerGray"
                topColor = DarkenColor(RGB(166, 166, 166), 0.4)
                bottomColor = DarkenColor(RGB(38, 38, 38), 0.4)
            ' Add more cases as needed
        End Select
    Else
        ' For the top bar, use the normal colors
        Select Case colorName
            Case "Green"
                topColor = RGB(0, 176, 80)
                bottomColor = RGB(0, 70, 32)
            Case "Gray"
                topColor = RGB(166, 166, 166)
                bottomColor = RGB(38, 38, 38)
            ' Add more cases as needed
        End Select
    End If

    ' Apply the gradient fill
    With cell.Shape.Fill
        .TwoColorGradient Style:=msoGradientVertical, Variant:=1
        .GradientStops(1).Color.RGB = topColor
        .GradientStops(2).Color.RGB = bottomColor
        .GradientAngle = 90
    End With

    ' Format text frame
    With cell.Shape.TextFrame
        .TextRange.Font.Size = Max(10, baseFontSize / Max(1, Sqr(colCount))) ' Adjusting size based on number of cells
        .TextRange.Font.Name = "Arial"
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = RGB(255, 255, 255) ' White color
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
    End With
End Sub

Function Max(ByVal v1 As Integer, ByVal v2 As Integer) As Integer
    If v1 > v2 Then Max = v1 Else Max = v2
End Function
Sub FormatTableBorders(pptTable As Table, borderColor As Long)
    If pptTable Is Nothing Then Exit Sub

    Dim i As Integer, j As Integer
    For i = 1 To pptTable.Rows.count
        For j = 1 To pptTable.Columns.count
            With pptTable.cell(i, j).Borders(ppBorderTop)
                .ForeColor.RGB = borderColor
                .Weight = 1
            End With
            With pptTable.cell(i, j).Borders(ppBorderBottom)
                .ForeColor.RGB = borderColor
                .Weight = 1
            End With
            With pptTable.cell(i, j).Borders(ppBorderLeft)
                .ForeColor.RGB = borderColor
                .Weight = 1
            End With
            With pptTable.cell(i, j).Borders(ppBorderRight)
                .ForeColor.RGB = borderColor
                .Weight = 1
            End With
        Next j
    Next i
End Sub
Sub CreateTwoBarTimelineTable(SlideIndex As Integer, startDate As Date, endDate As Date, timelineColor As String, timelineType As String)
    Dim pptSlide As slide
    Dim pptTableTop As Table, pptTableBottom As Table
    Dim slideWidth As Single
    Dim i As Integer
    Dim topBarCount As Integer, bottomBarCount As Integer
    Dim currentYear As Integer, totalMonths As Integer, monthsInYear As Integer
    Dim segmentStartDate As Date, segmentEndDate As Date
    Dim totalDays As Integer, daysInSegment As Integer
    Dim topColorString As String, bottomColorString As String
    Dim topDates() As Date, bottomDates() As Date

    Set pptSlide = ActivePresentation.Slides(SlideIndex)
    slideWidth = pptSlide.parent.PageSetup.slideWidth

    ' Determine the number of segments (columns) for each bar
    If timelineType = "Months" Then
        topBarCount = year(endDate) - year(startDate) + 1
        bottomBarCount = DateDiff("m", startDate, endDate) + 1
        totalMonths = bottomBarCount
    ElseIf timelineType = "Days" Then
        totalDays = DateDiff("d", startDate, endDate) + 1
        topBarCount = DateDiff("m", startDate, endDate) + 1
        bottomBarCount = totalDays
    End If
    ReDim topDates(1 To topBarCount)
    ReDim bottomDates(1 To bottomBarCount)

    ' Assigning color strings based on the timelineColor
    topColorString = IIf(timelineColor = "Green", "Green", "Gray")
    bottomColorString = IIf(timelineColor = "Green", "DarkerGreen", "DarkerGray")

    ' Create and format the top bar
Set pptTableTop = pptSlide.Shapes.AddTable(1, topBarCount).Table

If timelineType = "Days" Then
    segmentStartDate = startDate
    Dim totalWidth As Single
    totalWidth = slideWidth  ' Total width of the timeline

    For i = 1 To topBarCount
        ' Calculate the end date for each segment
        segmentEndDate = DateSerial(year(segmentStartDate), month(segmentStartDate) + 1, 0)
        If segmentEndDate > endDate Then
            segmentEndDate = endDate
        End If

        ' Set the text for each cell
        pptTableTop.cell(1, i).Shape.TextFrame.TextRange.text = Format(segmentStartDate, "mmm yyyy")
        topDates(i) = DateSerial(year(segmentStartDate), month(segmentStartDate), 1)

        ' Calculate the width for each segment
        Dim segmentDays As Integer
        segmentDays = DateDiff("d", segmentStartDate, segmentEndDate) + 1
        pptTableTop.Columns(i).Width = (segmentDays / totalDays) * totalWidth

        ' Move to the next segment
        segmentStartDate = DateSerial(year(segmentEndDate), month(segmentEndDate) + 1, 1)

        ' Format the cell
        FormatCell pptTableTop.cell(1, i), topColorString, False, topBarCount
    Next i
End If


    If timelineType = "Months" Then
        currentYear = year(startDate)
        For i = 1 To topBarCount
            pptTableTop.cell(1, i).Shape.TextFrame.TextRange.text = CStr(currentYear)
            topDates(i) = DateSerial(currentYear, 1, 1)
            monthsInYear = IIf(currentYear = year(endDate), month(endDate), 12) - IIf(currentYear = year(startDate), month(startDate) - 1, 0)
            pptTableTop.Columns(i).Width = (monthsInYear / totalMonths) * slideWidth
            currentYear = currentYear + 1
            FormatCell pptTableTop.cell(1, i), topColorString, False, topBarCount
        Next i
    End If
    pptTableTop.parent.Name = "TopBar"
    pptTableTop.parent.Left = 0
    pptTableTop.parent.Top = pptSlide.parent.PageSetup.slideHeight / 4

    ' Create and format the bottom bar
    Set pptTableBottom = pptSlide.Shapes.AddTable(1, bottomBarCount).Table
    segmentStartDate = startDate
    For i = 1 To bottomBarCount
        If timelineType = "Months" Then
            pptTableBottom.cell(1, i).Shape.TextFrame.TextRange.text = Format(DateAdd("m", i - 1, startDate), "mmm")
            bottomDates(i) = DateAdd("m", i - 1, startDate)
        ElseIf timelineType = "Days" Then
            pptTableBottom.cell(1, i).Shape.TextFrame.TextRange.text = Format(segmentStartDate, "d")
            bottomDates(i) = segmentStartDate
            segmentStartDate = DateAdd("d", 1, segmentStartDate)
        End If
        pptTableBottom.Columns(i).Width = slideWidth / bottomBarCount
        FormatCell pptTableBottom.cell(1, i), bottomColorString, True, bottomBarCount
    Next i
    pptTableBottom.parent.Name = "BottomBar"
    pptTableBottom.parent.Left = 0
    pptTableTop.Rows(1).Height = BAR_HEIGHT          ' fixed height so the two rows align + stay consistent
    pptTableBottom.Rows(1).Height = BAR_HEIGHT
    pptTableBottom.parent.Top = pptTableTop.parent.Top + BAR_HEIGHT

    ' Formatting table borders
    FormatTableBorders pptTableTop, RGB(0, 0, 0)
    FormatTableBorders pptTableBottom, RGB(0, 0, 0)
    FinalizeDateBar pptSlide, Array(pptTableTop.parent, pptTableBottom.parent), timelineType, Array(topDates, bottomDates)
End Sub

Function DarkenColor(col As Long, ByVal darkenPercent As Single) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = col Mod 256
    g = (col \ 256) Mod 256
    b = (col \ 65536) Mod 256

    r = r * (1 - darkenPercent)
    g = g * (1 - darkenPercent)
    b = b * (1 - darkenPercent)

    DarkenColor = RGB(r, g, b)
End Function

