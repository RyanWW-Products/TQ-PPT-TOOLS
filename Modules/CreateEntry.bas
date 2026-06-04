Attribute VB_Name = "CreateEntry"
Option Explicit

Sub AddEntry(control As IRibbonControl)
    Dim inputDays As String
    inputDays = InputBox("Enter the day or range of days (DD, DD-DD, or DD,DD,DD...):")

    Dim days() As String
    days = Split(inputDays, ",")
    Dim dayRange As String
    Dim i As Integer

    For i = LBound(days) To UBound(days)
        dayRange = days(i)
        ' Process each day or range of days
        If InStr(dayRange, "-") > 0 Then
            ' Range of days
            ProcessDayRange dayRange
        Else
            ' Single day
            If IsNumeric(dayRange) Then
                CreateCalendarEntry CInt(dayRange)
            Else
                MsgBox "Invalid input for day: " & dayRange, vbOKOnly + vbExclamation, "Invalid Day Input"
            End If
        End If
    Next i
End Sub

Sub ProcessDayRange(range As String)
    Dim dayParts() As String
    dayParts = Split(range, "-")
    If UBound(dayParts) = 1 Then
        If IsNumeric(dayParts(0)) And IsNumeric(dayParts(1)) Then
            Dim j As Integer
            For j = CInt(dayParts(0)) To CInt(dayParts(1))
                CreateCalendarEntry j
            Next j
        Else
            MsgBox "Invalid range: " & range, vbOKOnly + vbExclamation, "Invalid Range"
        End If
    Else
        MsgBox "Invalid range format: " & range, vbOKOnly + vbExclamation, "Invalid Format"
    End If
End Sub

Sub CreateCalendarEntry(ByVal entryDay As Integer)
    Dim ppt As PowerPoint.Application
    Dim slide As PowerPoint.slide
    Dim calendarMonth As Integer
    Dim calendarYear As Integer
    Dim row As Integer
    Dim column As Integer
    Dim ShapeIndex As Integer
    Dim entryShape As PowerPoint.Shape

    Set ppt = PowerPoint.Application
    Set slide = ppt.ActiveWindow.View.slide

    ' Retrieve the month and year from the slide's tags
    If IsNumeric(slide.Tags("CalendarMonth")) Then
        calendarMonth = CInt(slide.Tags("CalendarMonth"))
    Else
        MsgBox "The 'CalendarMonth' tag is missing or not a number. Please check the tag settings."
        Exit Sub
    End If

    If IsNumeric(slide.Tags("CalendarYear")) Then
        calendarYear = CInt(slide.Tags("CalendarYear"))
    Else
        MsgBox "The 'CalendarYear' tag is missing or not a number. Please check the tag settings."
        Exit Sub
    End If

    ' Determine the row and column for the entry day in the calendar
    Dim firstDayOfMonth As Date
    firstDayOfMonth = DateSerial(calendarYear, calendarMonth, 1)
    Dim startDayOfWeek As Integer
    startDayOfWeek = Weekday(firstDayOfMonth, vbSunday)
    row = ((entryDay + startDayOfWeek - 2) \ 7) + 1
    column = ((entryDay + startDayOfWeek - 2) Mod 7) + 1

    ' Determine the suffix for the new entry
    Dim suffix As String
    suffix = NextAvailableSuffix(slide, entryDay)

    If suffix = "" Then
        MsgBox "Maximum number of entries for a single day reached."
        Exit Sub
    End If

  ' Calculate the size and position of the new entry shape
Dim cellShape As PowerPoint.Shape
Set cellShape = slide.Shapes("CalendarTable").Table.cell(row, column).Shape
Dim newEntry As PowerPoint.Shape
Set newEntry = slide.Shapes.AddShape(msoShapeRectangle, _
    cellShape.Left, cellShape.Top, cellShape.Width, cellShape.Height)

' Name the new entry with the correct suffix
newEntry.Name = "Entry_" & entryDay & "_" & suffix

' Set the initial color for the new entry
Dim totalEntries As Integer
totalEntries = CountDayEntries(slide, entryDay)
newEntry.Fill.ForeColor.RGB = GetColorByIndex(totalEntries)



    ' Add example text in Arial, black font
    With newEntry.TextFrame.TextRange
        .text = "Example text"
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Color.RGB = RGB(0, 0, 0) ' Black color
    End With
    

    ' Adjust the size and position based on the total entries for the day
    Select Case totalEntries
        Case 1
            ' First entry; takes up the whole cell
            newEntry.Width = cellShape.Width
            newEntry.Height = cellShape.Height
        Case 2 To 4
            ' Resize and reposition existing entries and the new entry
            ResizeAndPositionEntries slide, row, column, totalEntries, entryDay
        Case Else
            MsgBox "Error in entry positioning."
            newEntry.Delete
            Exit Sub
    End Select
End Sub

Function NextAvailableSuffix(slide As PowerPoint.slide, ByVal day As Integer) As String
    Dim ascii As Integer
    For ascii = Asc("A") To Asc("D")
        Dim suffix As String
        suffix = Chr(ascii)
        If Not ShapeExists(slide, "Entry_" & day & "_" & suffix) Then
            NextAvailableSuffix = suffix
            Exit Function
        End If
    Next ascii
    NextAvailableSuffix = "" ' No available suffix
End Function


Function CountDayEntries(slide As PowerPoint.slide, ByVal entryDay As Integer) As Integer
    Dim shp As PowerPoint.Shape
    Dim count As Integer
    count = 0

    For Each shp In slide.Shapes
        If InStr(1, shp.Name, "Entry_" & entryDay & "_") > 0 Then
            count = count + 1
        End If
    Next shp

    Debug.Print "CountDayEntries: Total shapes counted for day " & entryDay & ": " & count
    CountDayEntries = count
End Function


Function ShapeExists(slide As PowerPoint.slide, ByVal shapeName As String) As Boolean
    On Error Resume Next
    Dim shp As PowerPoint.Shape
    Set shp = slide.Shapes(shapeName)
    If Not shp Is Nothing Then
        ShapeExists = True
    Else
        ShapeExists = False
    End If
    On Error GoTo 0
End Function

Sub ResizeAndPositionEntries(slide As PowerPoint.slide, ByVal row As Integer, ByVal column As Integer, ByVal totalEntries As Integer, ByVal entryDay As Integer)
    Dim cellShape As PowerPoint.Shape
    Set cellShape = slide.Shapes("CalendarTable").Table.cell(row, column).Shape

    Dim entryName As String
    Dim entryShape As PowerPoint.Shape
    Dim i As Integer

    For i = 1 To totalEntries
        entryName = "Entry_" & entryDay & "_" & Chr(64 + i)
        If ShapeExists(slide, entryName) Then
            Set entryShape = slide.Shapes(entryName)

            ' Assign color to the shape based on the suffix
            entryShape.Fill.ForeColor.RGB = GetColorByIndex(i)

            Select Case totalEntries
                Case 1
                    ' First entry; takes up the whole cell
                    entryShape.Width = cellShape.Width
                    entryShape.Height = cellShape.Height

                Case 2
                    ' Halve the height for both entries
                    entryShape.Height = cellShape.Height / 2
                    entryShape.Top = cellShape.Top + (i - 1) * (cellShape.Height / 2)
                    entryShape.TextFrame.TextRange.Font.Size = 8

               Case 3
    ' Divide the cell into three parts; adjust two existing and place the new one
    entryShape.Width = cellShape.Width / 2
    entryShape.Height = cellShape.Height / 2
    If i <= 2 Then
        ' Position first two entries in the top half, side by side
        entryShape.Top = cellShape.Top
        entryShape.Left = cellShape.Left + (i - 1) * (cellShape.Width / 2)
    Else
        ' Position the third entry in the bottom half, spanning full width
        entryShape.Top = cellShape.Top + cellShape.Height / 2
        entryShape.Left = cellShape.Left
        entryShape.Width = cellShape.Width ' Span the full width for the third entry
    End If
    entryShape.TextFrame.TextRange.Font.Size = 7


               Case 4
    ' Adjust all four entries to each corner
    entryShape.Width = cellShape.Width / 2
    entryShape.Height = cellShape.Height / 2

    ' Determine the position based on the entry index (i)
    Select Case i
        Case 1  ' Top left
            entryShape.Top = cellShape.Top
            entryShape.Left = cellShape.Left
        Case 2  ' Top right
            entryShape.Top = cellShape.Top
            entryShape.Left = cellShape.Left + cellShape.Width / 2
        Case 3  ' Bottom left
            entryShape.Top = cellShape.Top + cellShape.Height / 2
            entryShape.Left = cellShape.Left
        Case 4  ' Bottom right
            entryShape.Top = cellShape.Top + cellShape.Height / 2
            entryShape.Left = cellShape.Left + cellShape.Width / 2
    End Select

    entryShape.TextFrame.TextRange.Font.Size = 7


                Case Else
                    MsgBox "Error in entry positioning."
                    entryShape.Delete
            End Select
        Else
            MsgBox "Shape named '" & entryName & "' not found.", vbExclamation
        End If
    Next i
End Sub

Function GetColorByIndex(ByVal index As Integer) As Long
    ' Redefine array structure
    Dim colors(1 To 4) As Long

    ' Define colors
    colors(1) = RGB(173, 216, 230) ' Light Blue
    colors(2) = RGB(144, 238, 144) ' Light Green
    colors(3) = RGB(255, 165, 0)   ' Orange
    colors(4) = RGB(255, 182, 193) ' Light Pink

    ' Cycle through the colors
    GetColorByIndex = colors((index - 1) Mod 4 + 1)
End Function








