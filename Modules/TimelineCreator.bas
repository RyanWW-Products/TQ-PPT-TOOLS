Attribute VB_Name = "TimelineCreator"
Option Explicit

' ============================================================================
' TIMELINE CREATOR  -  Excel -> year-banded chronological timeline
' ============================================================================
' Reads a standardized Excel template (Date / Time / Description columns matched
' by HEADER TEXT) and draws a year-banded timeline. Every position is computed
' at runtime; only the style values below are hardcoded (spec section 3).
'
' Fit behaviour:
'   * Autofit  -> shrink fonts/metrics so everything fits on ONE slide.
'   * Multi    -> split year-columns across multiple slides when asked.
'   * If even the minimum scale can't fit, it REPORTS (never clips).
'
' Idempotent: re-running deletes prior "EventCard*" shapes on the slide and any
' extra slides this tool created (tagged), then redraws -- no data loss.
'
' Entry point (ribbon button "Import Timeline"):  BuildTimeline
' ============================================================================

' ---- Style constants (section 3 -- the ONLY hardcoded layout values) -------
Private Const SHAPE_PREFIX  As String = "EventCard"
Private Const SLIDE_TAG     As String = "TLIMPORTER"

Private Const MARGIN        As Single = 36      ' slide edge margin (pt)
Private Const BAND_TOP      As Single = 72      ' top of the year band
Private Const BAND_HEIGHT   As Single = 30      ' year band height
Private Const BAND_GAP      As Single = 24      ' gap below band before first card
Private Const BOX_WIDTH     As Single = 168
Private Const DATE_HEIGHT   As Single = 21
Private Const ROW_GAP       As Single = 22
Private Const LINE_INSET    As Single = 18      ' where the leader line meets the box
Private Const COL_PAD       As Single = 8       ' inner padding each side of a column
Private Const MIN_COL_W     As Single = 132     ' below this, split years across slides

Private Const FS_DATE       As Single = 12      ' Arial 12 bold  (date box)
Private Const FS_DESC       As Single = 10      ' Arial 10       (description)
Private Const FS_YEAR       As Single = 14      ' Arial 14 bold  (year band)
Private Const MIN_SCALE     As Single = 0.55    ' autofit floor

Private Const BORDER_PT     As Single = 1.5

' ---- Event record (section 2) ----------------------------------------------
' (Module-level Type MUST sit in the declarations section, before any procedure.)
Private Type TLEvent
    SortKey   As Double      ' constructed date serial, for sorting
    RawDate   As Date
    Year      As Integer
    YearFrac  As Double      ' (DayOfYear-1)/(DaysInYear-1); leader-line X only
    DateLabel As String
    Desc      As String
    OrigIndex As Long
    DescH     As Single      ' measured description height at base font
End Type

' Colors (navy 043D66 + white)
Private Function NAVY() As Long
    NAVY = RGB(4, 61, 102)
End Function
Private Function WHITE() As Long
    WHITE = RGB(255, 255, 255)
End Function

' ============================================================================
' ENTRY POINT
' ============================================================================
Public Sub BuildTimeline(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveTargetSlide()
    If sld Is Nothing Then
        MsgBox "Open a presentation and select a slide first.", vbExclamation, "Timeline Creator"
        Exit Sub
    End If

    ' Offer template creation up front
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Build a timeline from your Excel data?" & vbCrLf & vbCrLf & _
                 "Yes  = pick your filled-in spreadsheet" & vbCrLf & _
                 "No   = create a blank template to fill in first", _
                 vbYesNoCancel + vbQuestion, "Timeline Creator")
    If ans = vbCancel Then Exit Sub
    If ans = vbNo Then CreateTimelineTemplate: Exit Sub

    Dim filePath As String
    filePath = PickExcelFile()
    If filePath = "" Then Exit Sub

    Dim events() As TLEvent, n As Long, errLog As String
    n = ReadEvents(filePath, events, errLog)
    If n = 0 Then
        MsgBox "No usable events were found in that file." & vbCrLf & vbCrLf & errLog, _
               vbExclamation, "Timeline Creator"
        Exit Sub
    End If

    ' Options
    Dim includeEmpty As Boolean, allowMulti As Boolean, animate As Boolean
    includeEmpty = (MsgBox("Show empty years (gaps with no events) as columns too?", _
                    vbYesNo + vbQuestion, "Timeline Creator") = vbYes)
    allowMulti = (MsgBox("If it won't fit on one slide, split it across multiple slides?" & vbCrLf & vbCrLf & _
                    "Yes = split across slides" & vbCrLf & "No = shrink to fit one slide", _
                    vbYesNo + vbQuestion, "Timeline Creator") = vbYes)
    animate = (MsgBox("Add a fade-in (on click, in chronological order) to each entry?", _
                    vbYesNo + vbQuestion, "Timeline Creator") = vbYes)

    ' Idempotent: clear our prior output
    ClearPriorOutput sld

    Dim summary As String
    summary = DrawTimeline(sld, events, n, includeEmpty, allowMulti, animate)

    ReportSummary summary, errLog
    Exit Sub
Fail:
    MsgBox "Timeline build failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Timeline Creator"
End Sub

' ============================================================================
' READ + PARSE  (sections 1, 2)
' ============================================================================
Private Function ReadEvents(ByVal filePath As String, ByRef events() As TLEvent, _
                            ByRef errLog As String) As Long
    Dim xl As Object, wb As Object, ws As Object
    Dim colDate As Long, colTime As Long, colDesc As Long
    Dim lastCol As Long, c As Long, hdr As String
    Dim r As Long, n As Long
    Dim rawDate As Variant, rawTime As Variant, desc As String
    Dim ev As TLEvent

    Set xl = CreateObject("Excel.Application")     ' late-bound (section 8)
    xl.Visible = False
    On Error GoTo CleanFail
    Set wb = xl.Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    ' Match columns by header text (section 1) -- scan row 1
    lastCol = ws.Cells(1, ws.Columns.count).End(-4159).Column   ' xlToLeft
    If lastCol < 1 Or lastCol > 200 Then lastCol = 50
    For c = 1 To lastCol
        hdr = LCase$(Trim$(CStr(ws.Cells(1, c).Value & "")))
        Select Case hdr
            Case "date":        colDate = c
            Case "time":        colTime = c
            Case "description", "desc", "event": colDesc = c
        End Select
    Next c
    If colDate = 0 Or colDesc = 0 Then
        errLog = "Couldn't find the required 'Date' and 'Description' columns in row 1."
        wb.Close False: xl.Quit
        ReadEvents = 0: Exit Function
    End If

    ReDim events(1 To 1000)
    r = 2
    Do
        rawDate = ws.Cells(r, colDate).Value
        desc = Trim$(CStr(ws.Cells(r, colDesc).Value & ""))
        ' stop at first fully blank row (section 1)
        If (Len(Trim$(CStr(rawDate & ""))) = 0) And (Len(desc) = 0) Then Exit Do

        If Len(desc) = 0 Then
            errLog = errLog & "Row " & r & ": blank description, skipped." & vbCrLf
        ElseIf Not ParseDateCell(rawDate, ev) Then
            errLog = errLog & "Row " & r & ": unparseable date '" & CStr(rawDate & "") & "', skipped." & vbCrLf
        Else
            If colTime > 0 Then
                rawTime = ws.Cells(r, colTime).Value
                If Len(Trim$(CStr(rawTime & ""))) > 0 Then ev.DateLabel = ev.DateLabel & "  " & FormatTime(rawTime)
            End If
            ev.Desc = desc
            ev.OrigIndex = r
            n = n + 1
            If n > UBound(events) Then ReDim Preserve events(1 To UBound(events) + 1000)
            events(n) = ev
        End If
        r = r + 1
        If r > 100000 Then Exit Do                 ' safety
    Loop

    wb.Close False: xl.Quit
    Set xl = Nothing

    If n = 0 Then ReadEvents = 0: Exit Function
    ReDim Preserve events(1 To n)
    SortEvents events, n                            ' by date, OrigIndex tiebreak (section 2)
    ReadEvents = n
    Exit Function
CleanFail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    errLog = "Excel read error: " & Err.Description
    ReadEvents = 0
End Function

' Parse a real date, or "April 2023" (month), or "2023" (year). Sets the record's
' RawDate/Year/YearFrac/SortKey/DateLabel. Returns False if unparseable.
Private Function ParseDateCell(ByVal v As Variant, ByRef ev As TLEvent) As Boolean
    Dim s As String, mo As Integer, yr As Integer, d As Date
    ParseDateCell = False

    If IsDate(v) And (VarType(v) = vbDate) Then
        d = CDate(v)
        ev.RawDate = d
        ev.DateLabel = Format$(d, "mmm d, yyyy")
        GoTo Finish
    End If

    s = Trim$(CStr(v & ""))
    If Len(s) = 0 Then Exit Function

    ' "YYYY"
    If s Like "####" And IsNumeric(s) Then
        yr = CInt(s)
        If yr >= 1000 And yr <= 9999 Then
            d = DateSerial(yr, 1, 1)
            ev.RawDate = d
            ev.DateLabel = CStr(yr)
            GoTo Finish
        End If
    End If

    ' "MonthName YYYY" / "Mon YYYY"
    Dim parts() As String
    parts = Split(s, " ")
    If UBound(parts) >= 1 Then
        mo = MonthFromName(parts(0))
        If mo > 0 And IsNumeric(parts(UBound(parts))) Then
            yr = CInt(parts(UBound(parts)))
            If yr >= 1000 And yr <= 9999 Then
                d = DateSerial(yr, mo, 1)
                ev.RawDate = d
                ev.DateLabel = Format$(d, "mmmm yyyy")
                GoTo Finish
            End If
        End If
    End If

    ' last resort: let VBA try
    If IsDate(s) Then
        d = CDate(s)
        ev.RawDate = d
        ev.DateLabel = Format$(d, "mmm d, yyyy")
        GoTo Finish
    End If
    Exit Function

Finish:
    ev.Year = Year(ev.RawDate)
    ev.SortKey = CDbl(ev.RawDate)
    ev.YearFrac = YearFraction(ev.RawDate)
    ParseDateCell = True
End Function

Private Function MonthFromName(ByVal s As String) As Integer
    Dim m As Integer
    s = LCase$(Trim$(s))
    For m = 1 To 12
        If s = LCase$(Format$(DateSerial(2000, m, 1), "mmmm")) _
           Or s = LCase$(Format$(DateSerial(2000, m, 1), "mmm")) Then
            MonthFromName = m: Exit Function
        End If
    Next m
    MonthFromName = 0
End Function

' (DayOfYear-1)/(DaysInYear-1)  (section 2.4)
Private Function YearFraction(ByVal d As Date) As Double
    Dim doy As Long, diy As Long
    doy = DateDiff("d", DateSerial(Year(d), 1, 1), d) + 1
    diy = DateDiff("d", DateSerial(Year(d), 1, 1), DateSerial(Year(d), 12, 31)) + 1
    If diy <= 1 Then YearFraction = 0 Else YearFraction = (doy - 1) / (diy - 1)
End Function

Private Function FormatTime(ByVal v As Variant) As String
    On Error Resume Next
    If IsDate(v) Then FormatTime = Format$(CDate(v), "h:mm AM/PM") Else FormatTime = CStr(v & "")
End Function

Private Sub SortEvents(ByRef ev() As TLEvent, ByVal n As Long)
    Dim i As Long, j As Long, t As TLEvent
    For i = 1 To n - 1
        For j = i + 1 To n
            If (ev(j).SortKey < ev(i).SortKey) Or _
               (ev(j).SortKey = ev(i).SortKey And ev(j).OrigIndex < ev(i).OrigIndex) Then
                t = ev(i): ev(i) = ev(j): ev(j) = t
            End If
        Next j
    Next i
End Sub

' ============================================================================
' LAYOUT + DRAW  (sections 4-7) with autofit + multi-slide
' ============================================================================
Private Function DrawTimeline(ByVal firstSlide As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                              ByVal includeEmpty As Boolean, ByVal allowMulti As Boolean, _
                              ByVal animate As Boolean) As String
    Dim sld As slide, sw As Single, sh As Single
    sw = firstSlide.parent.PageSetup.slideWidth
    sh = firstSlide.parent.PageSetup.slideHeight

    ' distinct year columns (section 4)
    Dim years() As Integer, nYears As Long
    nYears = DistinctYears(ev, n, includeEmpty, years)

    ' pre-measure description heights at base font (section 6.1, off-slide)
    Dim i As Long
    For i = 1 To n
        ev(i).DescH = MeasureDescHeight(firstSlide, ev(i).Desc, BOX_WIDTH - 2 * COL_PAD, "Arial", FS_DESC)
    Next i

    ' how many year-columns fit per slide (width); the rest paginate when allowed
    Dim usableW As Single, colsPerSlide As Long
    usableW = sw - 2 * MARGIN
    colsPerSlide = nYears
    If allowMulti Then
        colsPerSlide = CLng(MaxS(1, Int(usableW / MIN_COL_W)))
        If colsPerSlide > nYears Then colsPerSlide = nYears
    End If

    Dim pageCount As Long, drawn As Long, overflowMsg As String
    Dim startCol As Long, p As Long
    pageCount = -Int(-nYears / colsPerSlide)        ' ceil
    Set sld = firstSlide

    For p = 1 To pageCount
        startCol = (p - 1) * colsPerSlide
        Dim pageYears() As Integer, pn As Long, k As Long
        pn = CLng(MinS(colsPerSlide, nYears - startCol))
        ReDim pageYears(1 To pn)
        For k = 1 To pn: pageYears(k) = years(startCol + k): Next k

        If p > 1 Then Set sld = NewPageSlide(firstSlide)
        drawn = drawn + DrawPage(sld, ev, n, pageYears, pn, sw, sh, animate, overflowMsg)
    Next p

    DrawTimeline = "Events drawn: " & drawn & " of " & n & vbCrLf & _
                   "Year columns: " & nYears & "   Slides: " & pageCount & overflowMsg
End Function

' Draw one slide's worth of year-columns; returns events drawn. Autofits vertically.
Private Function DrawPage(ByVal sld As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                          ByRef pageYears() As Integer, ByVal pn As Long, _
                          ByVal sw As Single, ByVal sh As Single, _
                          ByVal animate As Boolean, ByRef overflowMsg As String) As Long
    Dim colW As Single, availH As Single, bandCenterY As Single
    colW = (sw - 2 * MARGIN) / pn
    bandCenterY = BAND_TOP + BAND_HEIGHT / 2
    availH = sh - (BAND_TOP + BAND_HEIGHT + BAND_GAP) - MARGIN

    ' required height of the tallest column on this page -> autofit scale
    Dim ci As Long, maxH As Single, colH As Single, i As Long
    For ci = 1 To pn
        colH = 0
        For i = 1 To n
            If ev(i).Year = pageYears(ci) Then colH = colH + DATE_HEIGHT + ev(i).DescH + ROW_GAP
        Next i
        If colH > maxH Then maxH = colH
    Next ci

    Dim sc As Single
    sc = 1
    If maxH > availH And maxH > 0 Then sc = availH / maxH
    If sc < MIN_SCALE Then
        sc = MIN_SCALE
        overflowMsg = vbCrLf & "NOTE: content exceeds one slide even at minimum size on some columns." & _
                      vbCrLf & "Consider enabling multi-slide split or trimming descriptions."
    End If
    If sc > 1 Then sc = 1

    Dim drawn As Long, animSeq As Long
    ' year band (section 5) + each column (section 6)
    For ci = 1 To pn
        Dim colLeft As Single
        colLeft = MARGIN + (ci - 1) * colW
        DrawYearBox sld, pageYears(ci), colLeft, colW, sc
        drawn = drawn + DrawColumn(sld, ev, n, pageYears(ci), colLeft, colW, bandCenterY, _
                                   availH, sc, animate, animSeq)
    Next ci

    ApplyZOrder sld                                 ' section 6.7
    DrawPage = drawn
End Function

' One year-column: stacked cards + date-accurate leader lines (section 6)
Private Function DrawColumn(ByVal sld As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                            ByVal yr As Integer, ByVal colLeft As Single, ByVal colW As Single, _
                            ByVal bandCenterY As Single, ByVal availH As Single, ByVal sc As Single, _
                            ByVal animate As Boolean, ByRef animSeq As Long) As Long
    Dim innerL As Single, innerR As Single, boxW As Single
    innerL = colLeft + COL_PAD
    innerR = colLeft + colW - COL_PAD
    boxW = BOX_WIDTH * sc
    If boxW > (colW - 2 * COL_PAD) Then boxW = colW - 2 * COL_PAD

    Dim cursorY As Single, prevLeft As Single, drawn As Long, i As Long, lastDate As Double
    cursorY = BAND_TOP + BAND_HEIGHT + BAND_GAP * sc
    prevLeft = innerL
    lastDate = -1

    For i = 1 To n
        If ev(i).Year = yr Then
            Dim dateX As Single, boxLeft As Single, dH As Single, descH As Single
            dH = DATE_HEIGHT * sc
            descH = ev(i).DescH * sc

            ' date-accurate leader-line X (section 6.4) -- NOT the box center
            dateX = innerL + ClampD(ev(i).YearFrac, 0, 1) * (innerR - innerL)

            ' box left so the line meets it LINE_INSET from its edge; monotonic stagger (6.3)
            boxLeft = dateX - LINE_INSET * sc
            If boxLeft < prevLeft Then boxLeft = prevLeft        ' monotonic
            If boxLeft > innerR - boxW Then boxLeft = innerR - boxW
            If boxLeft < innerL Then boxLeft = innerL
            prevLeft = boxLeft

            ' the two boxes (section 6.2)
            Dim dateMidY As Single, dShp As Shape, descShp As Shape
            Set dShp = PlaceDateBox(sld, ev(i).DateLabel, boxLeft, cursorY, boxW, dH, sc)
            dateMidY = cursorY + dH / 2
            Set descShp = PlaceDescBox(sld, ev(i).Desc, boxLeft, cursorY + dH, boxW, descH, sc)

            ' leader line: vertical at the date-accurate X, band -> date-box middle (6.5)
            If ev(i).SortKey <> lastDate Then
                DrawLeaderLine sld, dateX, bandCenterY, dateMidY
            End If
            lastDate = ev(i).SortKey

            If animate Then
                AddFade sld, dShp, True
                AddFade sld, descShp, False
                animSeq = animSeq + 1
            End If
            cursorY = cursorY + dH + descH + ROW_GAP * sc
            drawn = drawn + 1
        End If
    Next i
    DrawColumn = drawn
End Function

Private Sub DrawYearBox(ByVal sld As slide, ByVal yr As Integer, ByVal colLeft As Single, _
                        ByVal colW As Single, ByVal sc As Single)
    Dim w As Single, sh As Shape
    w = MinS(colW - 2 * COL_PAD, 120 * sc)
    Set sh = sld.Shapes.AddShape(msoShapeRectangle, colLeft + (colW - w) / 2, BAND_TOP, w, BAND_HEIGHT * sc)
    StyleBox sh, NAVY(), WHITE(), CStr(yr), "Arial", FS_YEAR * sc, True, True
    sh.Name = SHAPE_PREFIX & "_Year_" & yr & "_" & sld.SlideIndex
End Sub

Private Function PlaceDateBox(ByVal sld As slide, ByVal label As String, ByVal x As Single, ByVal y As Single, _
                              ByVal w As Single, ByVal h As Single, ByVal sc As Single) As Shape
    Dim sh As Shape
    Set sh = sld.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    StyleBox sh, NAVY(), WHITE(), label, "Arial", FS_DATE * sc, True, True
    sh.Name = SHAPE_PREFIX & "_Date_" & sld.Shapes.count
    Set PlaceDateBox = sh
End Function

Private Function PlaceDescBox(ByVal sld As slide, ByVal text As String, ByVal x As Single, ByVal y As Single, _
                              ByVal w As Single, ByVal h As Single, ByVal sc As Single) As Shape
    Dim sh As Shape
    Set sh = sld.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    StyleBox sh, WHITE(), NAVY(), text, "Arial", FS_DESC * sc, False, False
    sh.TextFrame.WordWrap = msoTrue
    sh.TextFrame2.VerticalAnchor = msoAnchorTop
    sh.Name = SHAPE_PREFIX & "_Desc_" & sld.Shapes.count
    Set PlaceDescBox = sh
End Function

Private Sub DrawLeaderLine(ByVal sld As slide, ByVal x As Single, ByVal y1 As Single, ByVal y2 As Single)
    Dim ln As Shape
    Set ln = sld.Shapes.AddLine(x, y1, x, y2)
    ln.line.ForeColor.RGB = NAVY()
    ln.line.Weight = BORDER_PT
    ln.Name = SHAPE_PREFIX & "_Line_" & sld.Shapes.count
End Sub

' Shared box styling
Private Sub StyleBox(ByVal sh As Shape, ByVal fill As Long, ByVal txtOrBorder As Long, _
                     ByVal text As String, ByVal fName As String, ByVal fSize As Single, _
                     ByVal bold As Boolean, ByVal centered As Boolean)
    sh.fill.ForeColor.RGB = fill
    sh.line.ForeColor.RGB = NAVY()
    sh.line.Weight = BORDER_PT
    With sh.TextFrame
        .MarginLeft = 4: .MarginRight = 4: .MarginTop = 2: .MarginBottom = 2
        With .TextRange
            .text = text
            .Font.Name = fName
            .Font.Size = fSize
            .Font.bold = IIf(bold, msoTrue, msoFalse)
            If fill = NAVY() Then .Font.Color.RGB = WHITE() Else .Font.Color.RGB = NAVY()
            .ParagraphFormat.Alignment = IIf(centered, ppAlignCenter, ppAlignLeft)
        End With
        .VerticalAnchor = msoAnchorMiddle
    End With
End Sub

' z-order without naming the band: push lines behind the boxes (section 6.7)
Private Sub ApplyZOrder(ByVal sld As slide)
    Dim sh As Shape
    For Each sh In sld.Shapes
        If sh.Name Like SHAPE_PREFIX & "_Line_*" Then sh.ZOrder msoSendToBack
    Next sh
End Sub

' ============================================================================
' MEASURE / SLIDES / CLEAR / TEMPLATE / PICKER / REPORT  (section 8)
' ============================================================================
Private Function MeasureDescHeight(ByVal sld As slide, ByVal text As String, ByVal w As Single, _
                                   ByVal fName As String, ByVal fSize As Single) As Single
    Dim sh As Shape
    Set sh = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, -3000, -3000, w, 20)
    With sh.TextFrame
        .WordWrap = msoTrue
        .MarginLeft = 4: .MarginRight = 4: .MarginTop = 2: .MarginBottom = 2
        .TextRange.text = text
        .TextRange.Font.Name = fName
        .TextRange.Font.Size = fSize
        .AutoSize = ppAutoSizeShapeToFitText
    End With
    MeasureDescHeight = sh.Height
    sh.Delete
End Function

Private Function ActiveTargetSlide() As slide
    On Error Resume Next
    If ActivePresentation.Slides.count = 0 Then Exit Function
    Set ActiveTargetSlide = ActiveWindow.View.slide
    If ActiveTargetSlide Is Nothing Then Set ActiveTargetSlide = ActivePresentation.Slides(1)
End Function

Private Function NewPageSlide(ByVal refSlide As slide) As slide
    Dim s As slide
    Set s = ActivePresentation.Slides.Add(ActivePresentation.Slides.count + 1, ppLayoutBlank)
    s.Tags.Add SLIDE_TAG, "1"
    Set NewPageSlide = s
End Function

Private Sub ClearPriorOutput(ByVal sld As slide)
    ' delete prior cards on the active slide
    Dim i As Long
    For i = sld.Shapes.count To 1 Step -1
        If sld.Shapes(i).Name Like SHAPE_PREFIX & "*" Then sld.Shapes(i).Delete
    Next i
    ' delete extra slides this tool created previously
    Dim s As Long
    For s = ActivePresentation.Slides.count To 1 Step -1
        If ActivePresentation.Slides(s).SlideIndex <> sld.SlideIndex Then
            If ActivePresentation.Slides(s).Tags(SLIDE_TAG) = "1" Then ActivePresentation.Slides(s).Delete
        End If
    Next s
End Sub

Private Function PickExcelFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select your timeline spreadsheet"
    fd.Filters.Clear
    fd.Filters.Add "Excel files", "*.xlsx; *.xlsm; *.xls"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then PickExcelFile = fd.SelectedItems(1)
End Function

Private Sub CreateTimelineTemplate()
    Dim fd As FileDialog, savePath As String
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "Save the blank timeline template"
    fd.InitialFileName = "Timeline Template.xlsx"
    If fd.Show <> -1 Then Exit Sub
    savePath = fd.SelectedItems(1)

    Dim xl As Object, wb As Object, ws As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "Date"
    ws.Range("B1").Value = "Time"
    ws.Range("C1").Value = "Description"
    ws.Range("A1:C1").Font.bold = True
    ws.Range("A2").Value = "1/15/2023":  ws.Range("C2").Value = "Example: exact date"
    ws.Range("A3").Value = "April 2023":  ws.Range("C3").Value = "Example: month + year"
    ws.Range("A4").Value = "2024":        ws.Range("C4").Value = "Example: year only"
    ws.Columns("A:C").AutoFit
    On Error Resume Next
    wb.SaveAs savePath
    wb.Close True
    xl.Quit
    On Error GoTo 0
    MsgBox "Blank template saved:" & vbCrLf & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "Fill in the Date and Description columns (Time optional), then run Import Timeline again.", _
           vbInformation, "Timeline Creator"
End Sub

Private Sub ReportSummary(ByVal summary As String, ByVal errLog As String)
    Dim msg As String
    msg = summary
    If Len(errLog) > 0 Then msg = msg & vbCrLf & vbCrLf & "Skipped / notes:" & vbCrLf & errLog
    MsgBox msg, vbInformation, "Timeline Creator"
End Sub

Private Function ClampD(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then
        ClampD = lo
    ElseIf v > hi Then
        ClampD = hi
    Else
        ClampD = v
    End If
End Function

Private Function MinS(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then MinS = a Else MinS = b
End Function
Private Function MaxS(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxS = a Else MaxS = b
End Function

' Optional fade entrance (section 7). Columns draw earliest-year first and each
' column is in date order, so draw order is already chronological.
Private Sub AddFade(ByVal sld As slide, ByVal shp As Shape, ByVal firstOfCard As Boolean)
    On Error Resume Next
    Dim trig As Long
    trig = IIf(firstOfCard, msoAnimTriggerOnPageClick, msoAnimTriggerWithPrevious)
    sld.TimeLine.MainSequence.AddEffect shp, msoAnimEffectFade, , trig
End Sub

Private Function DistinctYears(ByRef ev() As TLEvent, ByVal n As Long, ByVal includeEmpty As Boolean, _
                               ByRef years() As Integer) As Long
    Dim minY As Integer, maxY As Integer, i As Long, cnt As Long, y As Integer
    minY = 32767: maxY = -32767
    For i = 1 To n
        If ev(i).Year < minY Then minY = ev(i).Year
        If ev(i).Year > maxY Then maxY = ev(i).Year
    Next i

    If includeEmpty Then
        cnt = maxY - minY + 1
        ReDim years(1 To cnt)
        For y = minY To maxY: years(y - minY + 1) = y: Next y
        DistinctYears = cnt
    Else
        ReDim years(1 To (maxY - minY + 1))
        For y = minY To maxY
            For i = 1 To n
                If ev(i).Year = y Then
                    cnt = cnt + 1: years(cnt) = y: Exit For
                End If
            Next i
        Next y
        ReDim Preserve years(1 To cnt)
        DistinctYears = cnt
    End If
End Function
