Attribute VB_Name = "TimelineCreator"
Option Explicit

' ============================================================================
' TIMELINE CREATOR  -  Excel -> DateBar-style chronological timeline
' ============================================================================
' Reads a standardized Excel template (Date / Time / Description columns matched
' by HEADER TEXT) and draws a timeline of the chosen unit (Years / Months / Days
' / Hours), reusing the DateBar look:
'   * The band is a gray/green gradient bar grouped + named "BottomBar".
'   * Each event is a Make-Entry-style entry (navy date box + white entry box,
'     GroupStyle-tagged, drop-shadowed) with a date-accurate LeadingLine.
'
' Gaps:  OFF -> contiguous bar (every unit first..last);  ON -> compact (only
'        units that have entries).  Unit count is capped at MAX_COLS.
' Fit:   autofit shrinks to one slide; optional multi-slide split by columns.
' Idempotent: re-running clears prior "EventCard*" shapes + the "BottomBar".
'
' NOTE: two-bar (Days->months / Months->years top row) is a follow-up pass.
' Entry point (ribbon "Make Timeline"):  BuildTimeline
' ============================================================================

Private Const SHAPE_PREFIX  As String = "EventCard"
Private Const SLIDE_TAG     As String = "TLIMPORTER"

Private Const MARGIN        As Single = 36
Private Const BAND_TOP      As Single = 44.64      ' 0.62" from the top of the slide (0.62 * 72)
Private Const BAND_HEIGHT   As Single = 30
Private Const BAND_GAP      As Single = 24
Private Const BOX_WIDTH     As Single = 168
Private Const DATE_HEIGHT   As Single = 21
Private Const ROW_GAP       As Single = 22
Private Const LINE_INSET    As Single = 18
Private Const COL_PAD       As Single = 8
Private Const MIN_COL_W     As Single = 132
Private Const LANE_GAP      As Single = 14        ' horizontal gap between packed lanes within a column

Private Const FS_DATE       As Single = 12
Private Const FS_DESC       As Single = 12        ' MUST match CreateTimelineEntry's drawn font (12*scale), or autofit under-measures and columns overflow
Private Const FS_LABEL      As Single = 12      ' fixed datebar label size (NOT scaled by autofit)
Private Const MIN_SCALE     As Single = 0.55
Private Const BORDER_PT     As Single = 1.5
Private Const MAX_COLS      As Long = 75        ' PowerPoint table / practical column ceiling

' ---- Weighted timeline (tiered, proportional column widths) -----------------
Private Const TIER_REL_THRESH As Single = 0.5   ' a tier break needs a >=50% relative jump...
Private Const TIER_ABS_GATE   As Single = 0.3   ' ...AND the jump must be >=30% of the full count range
Private Const TIER_MIN_FRAC   As Single = 0.14  ' a column is never thinner than ~14% (auto-backs off when many cols)

' ---- Event record (Type MUST sit before any procedure) ---------------------
Private Type TLEvent
    SortKey   As Double
    RawDate   As Date
    DateLabel As String
    Desc      As String
    OrigIndex As Long
    DescH     As Single
    UnitStart As Date        ' the date truncated to the chosen unit (column key)
    UnitFrac  As Double       ' position within the unit, 0..1 (leader-line X)
    Bates     As String       ' Bates / document number, parsed + held (not yet rendered)
End Type

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

    ' Timeline unit + bar color: auto-suggest, with a manual override (the DateBar dialog)
    Dim timelineType As String, timelineColor As String, detected As String
    detected = DetectType(events, n)
    If MsgBox("This looks like a " & UnitWord(detected) & " timeline." & vbCrLf & vbCrLf & _
              "Yes = continue with " & UnitWord(detected) & vbCrLf & _
              "No  = choose the type and color myself", _
              vbYesNo + vbQuestion, "Timeline Type") = vbYes Then
        timelineType = detected
        timelineColor = "Green"
    Else
        Dim frm As TimelineSettings
        Set frm = New TimelineSettings
        frm.Caption = "Timeline Type"
        frm.Show
        If frm.Tag <> "OK" Then Unload frm: Exit Sub
        timelineType = frm.timelineType
        timelineColor = frm.timelineColor
        Unload frm
    End If

    Dim allowGaps As Boolean, allowMulti As Boolean, doWipe As Boolean, allowWeighted As Boolean
    allowGaps = (MsgBox("Allow breaks / gaps in the timeline?" & vbCrLf & vbCrLf & _
                 "Yes = compact (only units that have entries)" & vbCrLf & _
                 "No  = continuous (every " & LCase$(UnitWord(timelineType)) & " from first to last)", _
                 vbYesNo + vbQuestion, "Timeline Creator") = vbYes)
    allowMulti = (MsgBox("If it won't fit on one slide, split it across multiple slides?", _
                 vbYesNo + vbQuestion, "Timeline Creator") = vbYes)
    doWipe = (MsgBox("Add a top-to-bottom wipe entrance to each entry (on click)?", _
                 vbYesNo + vbQuestion, "Timeline Creator") = vbYes)
    allowWeighted = (MsgBox("Weight column widths by how many events each " & _
                 LCase$(UnitWord(timelineType)) & " has?" & vbCrLf & vbCrLf & _
                 "Yes = busier units get wider columns (tiered)" & vbCrLf & _
                 "No  = every column the same width", _
                 vbYesNo + vbQuestion, "Timeline Creator") = vbYes)

    Dim summary As String
    If RenderTimeline(sld, events, n, timelineType, timelineColor, allowGaps, _
                      allowMulti, doWipe, allowWeighted, summary) Then ReportSummary summary, errLog
    Exit Sub
Fail:
    MsgBox "Timeline build failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Timeline Creator"
End Sub

' Shared render: assign units, build columns, clear, draw, and stash the source on
' the slide so "Toggle Weighted Timeline" can re-flow it without re-importing.
' Returns False (with its own message) if the column count exceeds the limit.
Private Function RenderTimeline(ByVal sld As slide, ByRef events() As TLEvent, ByVal n As Long, _
                               ByVal timelineType As String, ByVal timelineColor As String, _
                               ByVal allowGaps As Boolean, ByVal allowMulti As Boolean, _
                               ByVal doWipe As Boolean, ByVal allowWeighted As Boolean, _
                               ByRef summary As String) As Boolean
    Dim i As Long
    For i = 1 To n
        events(i).UnitStart = UnitStartOf(events(i).RawDate, timelineType)
        events(i).UnitFrac = UnitFracOf(events(i).RawDate, timelineType)
    Next i

    Dim cols() As Date, nCols As Long
    nCols = ComputeColumns(events, n, timelineType, allowGaps, cols)
    If nCols > MAX_COLS Then
        MsgBox "That would need " & nCols & " " & LCase$(UnitWord(timelineType)) & " columns, " & _
               "above the limit of " & MAX_COLS & "." & vbCrLf & vbCrLf & _
               "Choose a coarser unit (e.g. Months instead of Days), or turn ON " & _
               """Allow breaks / gaps"" to compact it.", vbExclamation, "Timeline Creator"
        RenderTimeline = False
        Exit Function
    End If

    ClearPriorOutput sld
    summary = DrawTimeline(sld, events, n, cols, nCols, timelineType, timelineColor, _
                           allowMulti, doWipe, allowWeighted)
    StoreTimelineState sld, events, n, timelineType, timelineColor, allowGaps, allowMulti, doWipe, allowWeighted
    RenderTimeline = True
End Function

' ============================================================================
' READ + PARSE
' ============================================================================
Private Function ReadEvents(ByVal filePath As String, ByRef events() As TLEvent, _
                            ByRef errLog As String) As Long
    Dim xl As Object, wb As Object, ws As Object
    Dim colDate As Long, colTime As Long, colDesc As Long, colBates As Long
    Dim lastCol As Long, c As Long, hdr As String
    Dim r As Long, n As Long
    Dim rawDate As Variant, rawTime As Variant, desc As String
    Dim ev As TLEvent

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    On Error GoTo CleanFail
    Set wb = xl.Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    lastCol = ws.Cells(1, ws.Columns.count).End(-4159).Column
    If lastCol < 1 Or lastCol > 200 Then lastCol = 50
    For c = 1 To lastCol
        hdr = LCase$(Trim$(CStr(ws.Cells(1, c).Value & "")))
        Select Case hdr
            Case "date":        colDate = c
            Case "time":        colTime = c
            Case "description", "desc", "event": colDesc = c
            Case "bates", "bates number", "bates no", "bates #", "bates range", _
                 "document number", "document no", "doc number", "doc no", "doc #", _
                 "document", "doc", "control number", "production number": colBates = c
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
            If colBates > 0 Then ev.Bates = Trim$(CStr(ws.Cells(r, colBates).Value & "")) Else ev.Bates = vbNullString
            ev.OrigIndex = r
            n = n + 1
            If n > UBound(events) Then ReDim Preserve events(1 To UBound(events) + 1000)
            events(n) = ev
        End If
        r = r + 1
        If r > 100000 Then Exit Do
    Loop

    wb.Close False: xl.Quit
    Set xl = Nothing

    If n = 0 Then ReadEvents = 0: Exit Function
    ReDim Preserve events(1 To n)
    SortEvents events, n
    ReadEvents = n
    Exit Function
CleanFail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    errLog = "Excel read error: " & Err.Description
    ReadEvents = 0
End Function

Private Function ParseDateCell(ByVal v As Variant, ByRef ev As TLEvent) As Boolean
    Dim s As String, mo As Integer, yr As Integer, d As Date
    ParseDateCell = False

    If IsDate(v) And (VarType(v) = vbDate) Then
        d = CDate(v)
        ev.RawDate = d
        ev.DateLabel = Format$(d, "mmm d")
        GoTo Finish
    End If

    s = Trim$(CStr(v & ""))
    If Len(s) = 0 Then Exit Function

    If s Like "####" And IsNumeric(s) Then
        yr = CInt(s)
        If yr >= 1000 And yr <= 9999 Then
            d = DateSerial(yr, 1, 1)
            ev.RawDate = d
            ev.DateLabel = CStr(yr)
            GoTo Finish
        End If
    End If

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

    If IsDate(s) Then
        d = CDate(s)
        ev.RawDate = d
        ev.DateLabel = Format$(d, "mmm d")
        GoTo Finish
    End If
    Exit Function

Finish:
    ev.SortKey = CDbl(ev.RawDate)
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

Private Function FormatTime(ByVal v As Variant) As String
    On Error Resume Next
    If IsNumeric(v) Then                       ' Excel/CSV store a clock time as a 0..1 fraction of a day
        FormatTime = Format$(CDate(CDbl(v)), "h:mm AM/PM")
    ElseIf IsDate(v) Then
        FormatTime = Format$(CDate(v), "h:mm AM/PM")
    Else
        FormatTime = CStr(v & "")
    End If
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
' UNIT MODEL  (Years / Months / Days / Hours)
' ============================================================================
Private Function UnitWord(ByVal t As String) As String
    Select Case t
        Case "Hours": UnitWord = "Hours"
        Case "Days":  UnitWord = "Days"
        Case "Months": UnitWord = "Months"
        Case Else:    UnitWord = "Years"
    End Select
End Function

' Guess the most likely unit from the data's granularity.
Private Function DetectType(ByRef ev() As TLEvent, ByVal n As Long) As String
    Dim yrs As Object, mos As Object, dys As Object, i As Long
    Set yrs = CreateObject("Scripting.Dictionary")
    Set mos = CreateObject("Scripting.Dictionary")
    Set dys = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        yrs(Year(ev(i).RawDate)) = 1
        mos(Year(ev(i).RawDate) * 100 + Month(ev(i).RawDate)) = 1
        dys(CLng(Int(CDbl(ev(i).RawDate)))) = 1
    Next i
    If yrs.count >= 2 Then
        DetectType = "Years"
    ElseIf mos.count >= 2 Then
        DetectType = "Months"
    ElseIf dys.count >= 2 Then
        DetectType = "Days"
    Else
        DetectType = "Hours"
    End If
End Function

Private Function IntervalCode(ByVal t As String) As String
    Select Case t
        Case "Hours": IntervalCode = "h"
        Case "Days":  IntervalCode = "d"
        Case "Months": IntervalCode = "m"
        Case Else:    IntervalCode = "yyyy"
    End Select
End Function

' Truncate a date down to the start of its unit (the column key).
Private Function UnitStartOf(ByVal d As Date, ByVal t As String) As Date
    Select Case t
        Case "Hours": UnitStartOf = Int(CDbl(d)) + TimeSerial(Hour(d), 0, 0)
        Case "Days":  UnitStartOf = DateSerial(Year(d), Month(d), Day(d))
        Case "Months": UnitStartOf = DateSerial(Year(d), Month(d), 1)
        Case Else:    UnitStartOf = DateSerial(Year(d), 1, 1)
    End Select
End Function

' Position of a date within its unit, 0..1 (governs the leader-line X only).
Private Function UnitFracOf(ByVal d As Date, ByVal t As String) As Double
    Dim u As Date, span As Double
    u = UnitStartOf(d, t)
    span = CDbl(DateAdd(IntervalCode(t), 1, u)) - CDbl(u)
    If span <= 0 Then UnitFracOf = 0 Else UnitFracOf = ClampD((CDbl(d) - CDbl(u)) / span, 0, 1)
End Function

' Build the list of column unit-starts: distinct (compact) or contiguous.
Private Function ComputeColumns(ByRef ev() As TLEvent, ByVal n As Long, ByVal t As String, _
                                ByVal allowGaps As Boolean, ByRef cols() As Date) As Long
    Dim minU As Date, maxU As Date, i As Long
    minU = ev(1).UnitStart: maxU = ev(1).UnitStart
    For i = 1 To n
        If ev(i).UnitStart < minU Then minU = ev(i).UnitStart
        If ev(i).UnitStart > maxU Then maxU = ev(i).UnitStart
    Next i

    Dim cnt As Long
    If allowGaps Then
        ' compact: distinct unit-starts (events are already date-sorted)
        ReDim cols(1 To n)
        For i = 1 To n
            If cnt = 0 Then
                cnt = 1: cols(1) = ev(i).UnitStart
            ElseIf ev(i).UnitStart <> cols(cnt) Then
                cnt = cnt + 1: cols(cnt) = ev(i).UnitStart
            End If
        Next i
        ReDim Preserve cols(1 To cnt)
    Else
        ' contiguous: every unit from minU to maxU
        Dim d As Date
        cnt = 0: d = minU
        Do While d <= maxU
            cnt = cnt + 1
            If cnt > MAX_COLS + 1 Then Exit Do          ' stop runaway; caller alerts on > MAX_COLS
            d = DateAdd(IntervalCode(t), 1, d)
        Loop
        ReDim cols(1 To cnt)
        d = minU
        For i = 1 To cnt
            cols(i) = d
            d = DateAdd(IntervalCode(t), 1, d)
        Next i
    End If
    ComputeColumns = cnt
End Function

' ============================================================================
' DRAW
' ============================================================================
Private Function DrawTimeline(ByVal firstSlide As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                              ByRef cols() As Date, ByVal nCols As Long, _
                              ByVal t As String, ByVal barColor As String, _
                              ByVal allowMulti As Boolean, ByVal doWipe As Boolean, _
                              ByVal allowWeighted As Boolean) As String
    Dim sw As Single, sh As Single
    sw = firstSlide.parent.PageSetup.slideWidth
    sh = firstSlide.parent.PageSetup.slideHeight

    Dim i As Long
    For i = 1 To n
        ev(i).DescH = MeasureDescHeight(firstSlide, ev(i).Desc, BOX_WIDTH - 2 * COL_PAD, "Arial", FS_DESC)
    Next i

    Dim colsPerSlide As Long
    colsPerSlide = nCols
    If allowMulti Then
        colsPerSlide = CLng(MaxS(1, Int(sw / MIN_COL_W)))
        If colsPerSlide > nCols Then colsPerSlide = nCols
    End If

    Dim pageCount As Long, drawn As Long, overflowMsg As String, p As Long, startCol As Long
    pageCount = -Int(-nCols / colsPerSlide)
    Dim sld As slide
    Set sld = firstSlide

    For p = 1 To pageCount
        startCol = (p - 1) * colsPerSlide
        Dim pn As Long, k As Long
        pn = CLng(MinS(colsPerSlide, nCols - startCol))
        Dim pageCols() As Date
        ReDim pageCols(1 To pn)
        For k = 1 To pn: pageCols(k) = cols(startCol + k): Next k
        If p > 1 Then Set sld = NewPageSlide(firstSlide)
        drawn = drawn + DrawPage(sld, ev, n, pageCols, pn, t, barColor, sw, sh, doWipe, allowWeighted, overflowMsg)
    Next p

    DrawTimeline = "Events drawn: " & drawn & " of " & n & vbCrLf & _
                   LCase$(UnitWord(t)) & " columns: " & nCols & "   Slides: " & pageCount & overflowMsg
End Function

Private Function DrawPage(ByVal sld As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                          ByRef pageCols() As Date, ByVal pn As Long, ByVal t As String, _
                          ByVal barColor As String, ByVal sw As Single, ByVal sh As Single, _
                          ByVal doWipe As Boolean, ByVal allowWeighted As Boolean, _
                          ByRef overflowMsg As String) As Long
    Dim availH As Single
    availH = sh - (BAND_TOP + BAND_HEIGHT + BAND_GAP) - MARGIN

    ' event count per column (drives weighting)
    Dim ci As Long, i As Long
    Dim cnt() As Long
    ReDim cnt(1 To pn)
    For ci = 1 To pn
        For i = 1 To n
            If ev(i).UnitStart = pageCols(ci) Then cnt(ci) = cnt(ci) + 1
        Next i
    Next ci

    ' per-column width (fraction of slide width): uniform, or tiered-by-count when weighted.
    ' Build cumulative left edges so every downstream X is a lookup, not (ci-1)*colW.
    Dim frac() As Single, colWArr() As Single, colLArr() As Single, acc As Single
    ReDim colWArr(1 To pn): ReDim colLArr(1 To pn)
    If allowWeighted Then
        ColumnWeights cnt, pn, frac
    Else
        ReDim frac(1 To pn)
        For ci = 1 To pn: frac(ci) = 1! / pn: Next ci
    End If
    acc = 0
    For ci = 1 To pn
        colWArr(ci) = frac(ci) * sw
        colLArr(ci) = acc
        acc = acc + colWArr(ci)
    Next ci

    ' per-column lanes + effective height for vertical autofit. A column that would
    ' overflow at full size AND is wide enough for more than one box-lane gets packed
    ' into lanes (column-major by date), shrinking its effective height toward the bar.
    Dim lanes() As Long, maxH As Single, singleH As Single, effH As Single, lc As Long
    ReDim lanes(1 To pn)
    maxH = 0
    For ci = 1 To pn
        singleH = 0
        For i = 1 To n
            If ev(i).UnitStart = pageCols(ci) Then singleH = singleH + DATE_HEIGHT + ev(i).DescH + ROW_GAP
        Next i
        lc = LanesFor(colWArr(ci))
        If lc > 1 And singleH > availH Then
            lanes(ci) = lc
            effH = PackedColHeight(ev, n, pageCols(ci), cnt(ci), lc)
        Else
            lanes(ci) = 1
            effH = singleH
        End If
        If effH > maxH Then maxH = effH
    Next ci

    Dim sc As Single
    sc = 1
    If maxH > availH And maxH > 0 Then sc = availH / maxH
    If sc < MIN_SCALE Then
        sc = MIN_SCALE
        overflowMsg = vbCrLf & "NOTE: content exceeds one slide even at minimum size on some columns." & _
                      vbCrLf & "Consider multi-slide split, gaps, or trimming descriptions."
    End If
    If sc > 1 Then sc = 1

    Dim bandBottom As Single
    bandBottom = BAND_TOP + BAND_HEIGHT          ' band is fixed height; only the entries below scale

    DrawBar sld, pageCols, pn, colWArr, colLArr, sc, t, barColor

    Dim drawn As Long, animSeq As Long
    For ci = 1 To pn
        drawn = drawn + DrawColumn(sld, ev, n, pageCols(ci), colLArr(ci), colWArr(ci), bandBottom, sc, doWipe, animSeq, lanes(ci), t)
    Next ci

    DrawPage = drawn
End Function

' The DateBar-style band: gradient cell per column, grouped + named "BottomBar".
Private Sub DrawBar(ByVal sld As slide, ByRef pageCols() As Date, ByVal pn As Long, _
                    ByRef colWArr() As Single, ByRef colLArr() As Single, ByVal sc As Single, _
                    ByVal t As String, ByVal barColor As String)
    Dim cellNames() As String, ci As Long, sh As Shape, isGreen As Boolean
    isGreen = (LCase$(barColor) = "green")
    ReDim cellNames(1 To pn)
    For ci = 1 To pn
        Set sh = sld.Shapes.AddShape(msoShapeRectangle, colLArr(ci), BAND_TOP, colWArr(ci), BAND_HEIGHT)
        StyleBandCell sh, GetSegmentLabel(t, pageCols(ci)), sc, isGreen
        sh.Name = SHAPE_PREFIX & "_BandCell_" & ci & "_" & sld.SlideIndex
        sh.Tags.Add "TLCellDate", CStr(CDbl(pageCols(ci)))   ' precise date per cell (for Date Snap)
        cellNames(ci) = sh.Name
    Next ci

    ' Tear cells adjacent to a gap (compact mode only; contiguous has no gaps).
    ' boundaryX = the shared edge between column ci and ci+1 = colLArr(ci+1).
    For ci = 1 To pn - 1
        If DateAdd(IntervalCode(t), 1, pageCols(ci)) <> pageCols(ci + 1) Then
            ApplyTear sld, cellNames(ci), True, colLArr(ci + 1), BAND_HEIGHT / 30!       ' TearA: left cell's right edge
            ApplyTear sld, cellNames(ci + 1), False, colLArr(ci + 1), BAND_HEIGHT / 30!  ' TearB (rot 180): right cell's left edge
        End If
    Next ci

    Dim grp As Shape
    If pn = 1 Then
        Set grp = sld.Shapes(cellNames(1))
    Else
        Set grp = sld.Shapes.Range(cellNames).Group
    End If
    grp.Name = NextDateBarName(sld)
    grp.Tags.Add "TLBar", "1"                        ' durable marker so lookups don't depend on the name
    grp.Tags.Add "TLType", t                        ' stamp the unit on the bar (shape tag -> survives copy/paste)
    With grp.Shadow                                 ' #1 drop shadow on the whole band
        .Visible = msoTrue
        .Type = msoShadow21
        .IncrementOffsetX 3
        .IncrementOffsetY 3
    End With
End Sub

Private Sub StyleBandCell(ByVal sh As Shape, ByVal label As String, ByVal sc As Single, ByVal isGreen As Boolean)
    With sh.fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientVertical, 1
        If isGreen Then
            .GradientStops(1).Color.RGB = RGB(0, 176, 80)
            .GradientStops(2).Color.RGB = RGB(0, 70, 32)
        Else
            .GradientStops(1).Color.RGB = RGB(166, 166, 166)
            .GradientStops(2).Color.RGB = RGB(38, 38, 38)
        End If
        .GradientAngle = 90
    End With
    sh.line.ForeColor.RGB = RGB(0, 0, 0)
    sh.line.Weight = 1
    With sh.TextFrame
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .text = label
            .Font.Name = "Arial"
            .Font.Size = FS_LABEL
            .Font.bold = msoTrue
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    ' text drop shadow (best-effort; VBA can be finicky here)
    On Error Resume Next
    With sh.TextFrame2.TextRange.Font.Shadow
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 3
        .Transparency = 0.5
        .Size = 100
        .OffsetX = 1.5
        .OffsetY = 1.5
    End With
    On Error GoTo 0
End Sub

' Subtract the tear freeform from a cell's edge (cell stays primary -> keeps fill + text).
Private Sub ApplyTear(ByVal sld As slide, ByVal cellName As String, ByVal isTearA As Boolean, _
                      ByVal boundaryX As Single, ByVal sc As Single, _
                      Optional ByVal bandTop As Single = BAND_TOP)
    Dim tear As Shape, shp As Shape, before As Object
    On Error GoTo Bail
    ' snapshot existing names so we can find the merged result afterwards
    Set before = CreateObject("Scripting.Dictionary")
    For Each shp In sld.Shapes
        before(shp.Name) = 1
    Next shp
    Set tear = BuildTear(sld, isTearA, boundaryX, sc, bandTop)
    tear.Name = "TLTear_tmp"
    before("TLTear_tmp") = 1
    sld.Shapes.Range(Array(cellName, "TLTear_tmp")).MergeShapes msoMergeSubtract, sld.Shapes(cellName)
    ' MergeShapes renames the result -> restore the cell name so grouping still works
    For Each shp In sld.Shapes
        If Not before.Exists(shp.Name) Then shp.Name = cellName: Exit For
    Next shp
    Exit Sub
Bail:
    On Error Resume Next
    If Not tear Is Nothing Then tear.Delete       ' a failed tear must not abort the build
End Sub

' Recreate the user's tear freeform at the cell boundary, scaled. TearA cuts the left
' cell's right edge; TearB cuts the right cell's left edge. In Tear Example.pptx TearB
' is authored rotated 180 deg (its jagged edge sits upper-right), so we rebuild its path
' and rotate the shape 180 to match. Node coords are normalized 0..1 of the shape box;
' offsets/sizes/rotation come from the example (authored for a 30pt-tall cell).
Private Function BuildTear(ByVal sld As slide, ByVal isTearA As Boolean, _
                           ByVal boundaryX As Single, ByVal sc As Single, _
                           Optional ByVal bandTop As Single = BAND_TOP) As Shape
    Dim ff As FreeformBuilder, sh As Shape, L As Single, T As Single, w As Single, h As Single
    If isTearA Then
        w = 29.2 * sc: h = 61.9 * sc
        L = boundaryX - 11 * sc: T = bandTop - 16.1 * sc
        Set ff = sld.Shapes.BuildFreeform(msoEditingCorner, L + 0.5739 * w, T)
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + w, T + 0.6106 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + 0.6717 * w, T + 0.6106 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + w, T + h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L, T + 0.997 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + 0.3412 * w, T + 0.5682 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + 0.1606 * w, T + 0.5682 * h
        ff.AddNodes msoSegmentCurve, msoEditingAuto, L + 0.2984 * w, T + 0.3667 * h, _
                    L + 0.4426 * w, T + 0.1894 * h, L + 0.5739 * w, T
        Set BuildTear = ff.ConvertToShape
    Else
        w = 26.9 * sc: h = 57.8 * sc
        L = boundaryX - 18.1 * sc: T = bandTop - 9 * sc
        Set ff = sld.Shapes.BuildFreeform(msoEditingCorner, L + 0.5522 * w, T)
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + w, T + 0.5828 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + 0.6443 * w, T + 0.5828 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + w, T + h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L, T + 0.9968 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + 0.2722 * w, T + 0.5763 * h
        ff.AddNodes msoSegmentLine, msoEditingAuto, L + 0.0696 * w, T + 0.5763 * h
        ff.AddNodes msoSegmentCurve, msoEditingAuto, L + 0.2537 * w, T + 0.3604 * h, _
                    L + 0.3682 * w, T + 0.2159 * h, L + 0.5522 * w, T
        Set sh = ff.ConvertToShape
        sh.Rotation = 180          ' TearB is authored rotated 180 deg in the example
        Set BuildTear = sh
    End If
End Function

' One unit-column of entries. lanes=1 -> the original date-accurate single stack;
' lanes>1 -> pack entries into that many box-lanes (column-major by date) so a busy,
' wide column rises toward the bar instead of running off the bottom of the slide.
Private Function DrawColumn(ByVal sld As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                            ByVal colUnit As Date, ByVal colLeft As Single, ByVal colW As Single, _
                            ByVal bandBottom As Single, ByVal sc As Single, _
                            ByVal doWipe As Boolean, ByRef animSeq As Long, ByVal lanes As Long, _
                            ByVal t As String) As Long
    Dim innerL As Single, innerR As Single, boxW As Single, drawn As Long, i As Long
    Dim dateX As Single, boxLeft As Single, boxesH As Single, grp As Shape
    Dim cursorY As Single, prevLeft As Single
    Dim nc As Long, laneW As Single, rpl As Long, k As Long, laneIdx As Long, laneCenter As Single
    Dim laneY() As Single, j As Long

    innerL = colLeft + COL_PAD
    innerR = colLeft + colW - COL_PAD
    boxW = BOX_WIDTH * sc
    If boxW > (colW - 2 * COL_PAD) Then boxW = colW - 2 * COL_PAD
    If boxW < 1 Then boxW = 1                        ' narrow weighted columns must stay positive

    If lanes <= 1 Then
        ' --- date-accurate single stack (original behavior) ---
        cursorY = bandBottom + BAND_GAP * sc
        prevLeft = innerL
        For i = 1 To n
            If ev(i).UnitStart = colUnit Then
                dateX = innerL + ClampD(ev(i).UnitFrac, 0, 1) * (innerR - innerL)
                boxLeft = dateX - LINE_INSET * sc
                If boxLeft < prevLeft Then boxLeft = prevLeft
                If boxLeft > innerR - boxW Then boxLeft = innerR - boxW
                If boxLeft < innerL Then boxLeft = innerL
                prevLeft = boxLeft
                Set grp = CreateTimelineEntry(sld, ev(i).DateLabel, ev(i).Desc, boxLeft, cursorY, boxW, _
                                              dateX, bandBottom, True, sc, boxesH, ev(i).RawDate, t)
                grp.Tags.Add "TLENTRY", "1"
                grp.Tags.Add "TLFullDate", CStr(CDbl(ev(i).RawDate))   ' label may omit the year; keep the real date for Date Snap
                grp.ZOrder msoSendToBack
                If doWipe Then
                    animSeq = animSeq + 1
                    AddWipe sld, grp, True
                End If
                cursorY = cursorY + boxesH + ROW_GAP * sc
                drawn = drawn + 1
            End If
        Next i
    Else
        ' --- multi-lane packing (column-major by date); leader line drops at the lane center ---
        For i = 1 To n
            If ev(i).UnitStart = colUnit Then nc = nc + 1
        Next i
        laneW = (colW - 2 * COL_PAD) / lanes
        rpl = -Int(-nc / lanes)                     ' ceil(nc / lanes) = rows per lane
        If rpl < 1 Then rpl = 1
        ReDim laneY(0 To lanes - 1)
        For j = 0 To lanes - 1: laneY(j) = bandBottom + BAND_GAP * sc: Next j
        k = 0
        For i = 1 To n
            If ev(i).UnitStart = colUnit Then
                laneIdx = k \ rpl
                If laneIdx > lanes - 1 Then laneIdx = lanes - 1
                laneCenter = innerL + (laneIdx + 0.5) * laneW
                boxLeft = laneCenter - boxW / 2
                Set grp = CreateTimelineEntry(sld, ev(i).DateLabel, ev(i).Desc, boxLeft, laneY(laneIdx), boxW, _
                                              laneCenter, bandBottom, True, sc, boxesH, ev(i).RawDate, t)
                grp.Tags.Add "TLENTRY", "1"
                grp.Tags.Add "TLFullDate", CStr(CDbl(ev(i).RawDate))   ' label may omit the year; keep the real date for Date Snap
                grp.ZOrder msoSendToBack
                If doWipe Then
                    animSeq = animSeq + 1
                    AddWipe sld, grp, True
                End If
                laneY(laneIdx) = laneY(laneIdx) + boxesH + ROW_GAP * sc
                k = k + 1
                drawn = drawn + 1
            End If
        Next i
    End If
    DrawColumn = drawn
End Function

' How many BOX_WIDTH-wide lanes fit across a column of width colW (at least 1).
Private Function LanesFor(ByVal colW As Single) As Long
    Dim k As Long
    k = Int((colW - 2 * COL_PAD + LANE_GAP) / (BOX_WIDTH + LANE_GAP))
    If k < 1 Then k = 1
    LanesFor = k
End Function

' Height of the tallest lane when a column's entries are packed column-major into
' `lanes` lanes (matches DrawColumn's assignment so the autofit scale is correct).
Private Function PackedColHeight(ByRef ev() As TLEvent, ByVal n As Long, ByVal colUnit As Date, _
                                 ByVal nc As Long, ByVal lanes As Long) As Single
    Dim laneH() As Single, rpl As Long, k As Long, i As Long, laneIdx As Long, mx As Single, j As Long
    ReDim laneH(0 To lanes - 1)
    rpl = -Int(-nc / lanes)
    If rpl < 1 Then rpl = 1
    k = 0
    For i = 1 To n
        If ev(i).UnitStart = colUnit Then
            laneIdx = k \ rpl
            If laneIdx > lanes - 1 Then laneIdx = lanes - 1
            laneH(laneIdx) = laneH(laneIdx) + DATE_HEIGHT + ev(i).DescH + ROW_GAP
            k = k + 1
        End If
    Next i
    mx = 0
    For j = 0 To lanes - 1
        If laneH(j) > mx Then mx = laneH(j)
    Next j
    PackedColHeight = mx
End Function

' Top-to-bottom wipe entrance (#6). The LeadingLine comes in with the entry group.
Private Sub AddWipe(ByVal sld As slide, ByVal shp As Shape, ByVal firstOfCard As Boolean)
    On Error Resume Next
    Dim eff As Effect, trig As Long
    trig = IIf(firstOfCard, msoAnimTriggerOnPageClick, msoAnimTriggerWithPrevious)
    Set eff = sld.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectWipe, , trig)
    eff.EffectParameters.Direction = msoAnimDirectionTop
End Sub

' ============================================================================
' MEASURE / SLIDES / CLEAR / TEMPLATE / PICKER / REPORT / UTIL
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
    Dim i As Long
    For i = sld.Shapes.count To 1 Step -1
        If (sld.Shapes(i).Tags("TLENTRY") = "1") Or IsDateBar(sld.Shapes(i)) _
           Or (sld.Shapes(i).Name Like SHAPE_PREFIX & "*") Then sld.Shapes(i).Delete
    Next i
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
    fd.Title = "Select your timeline file (Excel or CSV)"
    fd.Filters.Clear
    fd.Filters.Add "Timeline data (Excel / CSV)", "*.xlsx; *.xlsm; *.xls; *.csv"
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
    ws.Name = "Timeline"
    ws.Range("A1").Value = "Date"
    ws.Range("B1").Value = "Time"
    ws.Range("C1").Value = "Description"
    ws.Range("D1").Value = "Bates"
    ws.Range("A1:D1").Font.bold = True
    ws.Range("A2").Value = "1/15/2023":  ws.Range("C2").Value = "Example: exact date":     ws.Range("D2").Value = "ABC-000123"
    ws.Range("A3").Value = "April 2023":  ws.Range("C3").Value = "Example: month + year"
    ws.Range("A4").Value = "2024":        ws.Range("C4").Value = "Example: year only"
    ws.Columns("A:D").AutoFit

    ' second sheet: how to convert any source timeline with Copilot (prompt travels with the template)
    Dim ins As Object
    Set ins = wb.Worksheets.Add(After:=ws)
    ins.Name = "Convert with Copilot"
    ins.Range("A1").Value = "Convert any timeline (PDF / Word / Excel) into the 'Timeline' sheet using Microsoft Copilot:"
    ins.Range("A1").Font.bold = True
    ins.Range("A3").Value = "1.  In Microsoft Copilot (Chat, or Copilot in Word), attach or open your source timeline."
    ins.Range("A4").Value = "2.  Copy the prompt in the box below and paste it into Copilot."
    ins.Range("A5").Value = "3.  Paste Copilot's table into the 'Timeline' sheet under the headers (row 1), then run Make Timeline."
    ins.Range("A7").Value = "PROMPT  (click the cell below and copy it):"
    ins.Range("A7").Font.bold = True
    ins.Range("A8").Value = Replace(CopilotPromptText(), vbCrLf, Chr(10))
    ins.Range("A8").WrapText = True
    ins.Range("A8").VerticalAlignment = -4160      ' xlTop
    ins.Columns("A").ColumnWidth = 95
    Dim ok As Boolean
    On Error Resume Next
    wb.SaveAs savePath
    ok = (Err.Number = 0)
    Err.Clear
    wb.Close ok                 ' save-on-close only if SaveAs worked; avoids a re-save prompt on failure
    xl.Quit
    Set xl = Nothing
    On Error GoTo 0

    If ok Then
        MsgBox "Blank template saved:" & vbCrLf & vbCrLf & savePath & vbCrLf & vbCrLf & _
               "Fill in the Date and Description columns (Time and Bates optional), then run Make Timeline again." & vbCrLf & _
               "See the 'Convert with Copilot' tab to turn any PDF/Word/Excel timeline into this format.", _
               vbInformation, "Timeline Creator"
    Else
        MsgBox "Couldn't save the template to:" & vbCrLf & vbCrLf & savePath & vbCrLf & vbCrLf & _
               "Check the location isn't read-only or open elsewhere, then try again.", _
               vbExclamation, "Timeline Creator"
    End If
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

' ============================================================================
' WEIGHTED TIMELINE  -  tiered, proportional column widths
' ============================================================================
' Group columns into up to 3 size tiers by the GAPS in their event counts
' (roughly-equal counts share a tier + a width), weight each tier by its MEAN
' count (so a severe split makes the busy tier dominate), then enforce an
' order-preserving minimum width. Fills frac() (sums to 1).
Private Sub ColumnWeights(ByRef cnt() As Long, ByVal pn As Long, ByRef frac() As Single)
    ReDim frac(1 To pn)
    Dim i As Long
    If pn <= 1 Then
        If pn = 1 Then frac(1) = 1!
        Exit Sub
    End If

    ' all-equal -> one uniform tier
    Dim allEq As Boolean
    allEq = True
    For i = 2 To pn
        If cnt(i) <> cnt(1) Then allEq = False: Exit For
    Next i
    If allEq Then
        For i = 1 To pn: frac(i) = 1! / pn: Next i
        Exit Sub
    End If

    ' sort column indices ascending by count (carry original index in ord)
    Dim ord() As Long, a As Long, b As Long, keyi As Long
    ReDim ord(1 To pn)
    For i = 1 To pn: ord(i) = i: Next i
    For a = 2 To pn
        keyi = ord(a): b = a - 1
        Do While b >= 1
            If cnt(ord(b)) <= cnt(keyi) Then Exit Do
            ord(b + 1) = ord(b): b = b - 1
        Loop
        ord(b + 1) = keyi
    Next a
    Dim sv() As Long
    ReDim sv(1 To pn)
    For i = 1 To pn: sv(i) = cnt(ord(i)): Next i

    ' choose up to 2 tier breaks: a gap must clear BOTH the absolute-range gate and
    ' the relative-jump threshold; keep the two strongest by relative jump.
    Dim rng As Single
    rng = sv(pn) - sv(1): If rng < 1 Then rng = 1
    Dim bi1 As Long, bi2 As Long, bs1 As Single, bs2 As Single
    bi1 = 0: bi2 = 0: bs1 = 0: bs2 = 0
    Dim cur As Single, relG As Single, absF As Single
    For i = 1 To pn - 1
        cur = sv(i): If cur < 1 Then cur = 1
        relG = (sv(i + 1) - cur) / cur
        absF = (sv(i + 1) - sv(i)) / rng
        If absF >= TIER_ABS_GATE And relG >= TIER_REL_THRESH Then
            If relG > bs1 Then
                bs2 = bs1: bi2 = bi1: bs1 = relG: bi1 = i
            ElseIf relG > bs2 Then
                bs2 = relG: bi2 = i
            End If
        End If
    Next i

    ' walk sorted order assigning tier rank, incrementing after each chosen break
    Dim tier() As Long, tr As Long
    ReDim tier(1 To pn)
    tr = 0
    For i = 1 To pn
        tier(i) = tr
        If (bi1 > 0 And i = bi1) Or (bi2 > 0 And i = bi2) Then tr = tr + 1
    Next i
    Dim numT As Long
    numT = tr + 1

    ' tier weight = mean count of its members (proportional -> dramatic on severe splits)
    Dim tsum() As Double, tnum() As Long
    ReDim tsum(0 To numT - 1): ReDim tnum(0 To numT - 1)
    For i = 1 To pn
        tsum(tier(i)) = tsum(tier(i)) + sv(i)
        tnum(tier(i)) = tnum(tier(i)) + 1
    Next i

    Dim w() As Single, tot As Single
    ReDim w(1 To pn)
    For i = 1 To pn
        w(ord(i)) = tsum(tier(i)) / tnum(tier(i))    ' map tier weight back to original column order
    Next i
    tot = 0
    For i = 1 To pn: tot = tot + w(i): Next i
    For i = 1 To pn: frac(i) = w(i) / tot: Next i

    ApplyMinFloor frac, pn
End Sub

' Raise any column below the (capacity-aware) minimum up to it, taking the shortfall
' proportionally from the columns above the minimum -> ordering and within-tier
' equality are preserved.
Private Sub ApplyMinFloor(ByRef frac() As Single, ByVal pn As Long)
    Dim effMin As Single, i As Long, deficit As Single, excess As Single, takeRatio As Single
    effMin = TIER_MIN_FRAC
    If 0.6! / pn < effMin Then effMin = 0.6! / pn    ' back off when many columns
    For i = 1 To pn
        If frac(i) < effMin Then
            deficit = deficit + (effMin - frac(i))
        Else
            excess = excess + (frac(i) - effMin)
        End If
    Next i
    If deficit <= 0! Or excess <= 0! Then Exit Sub
    takeRatio = deficit / excess
    If takeRatio > 1! Then takeRatio = 1!
    For i = 1 To pn
        If frac(i) < effMin Then
            frac(i) = effMin
        Else
            frac(i) = frac(i) - (frac(i) - effMin) * takeRatio
        End If
    Next i
End Sub

' ============================================================================
' TIMELINE EDITS  (ribbon: Tear Gap, Toggle Weighted Timeline)
' ============================================================================
' Tear one edge of the SELECTED shape: build the tear scaled to the shape's height
' and boolean-subtract it (the shape keeps its fill / text / name). Reuses the
' timeline tear via ApplyTear's optional bandTop anchor.
Public Sub TearGap(control As IRibbonControl)
    On Error GoTo Fail
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select one shape to tear first.", vbExclamation, "Tear Gap": Exit Sub
    End If
    If ActiveWindow.Selection.ShapeRange.count <> 1 Then
        MsgBox "Select exactly one shape.", vbExclamation, "Tear Gap": Exit Sub
    End If
    Dim shp As Shape, sld As slide
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    Set sld = shp.Parent

    Select Case shp.Type        ' boolean ops need a real geometric shape
        Case msoPicture, msoLinkedPicture, msoGroup, msoPlaceholder, msoChart, msoTable, msoSmartArt, msoEmbeddedOLEObject, msoLinkedOLEObject
            MsgBox "That shape can't be torn - pictures, groups, tables, placeholders and charts can't be boolean-subtracted." & vbCrLf & vbCrLf & _
                   "Use an autoshape, a text box, or a Table-to-Object cell.", vbExclamation, "Tear Gap": Exit Sub
    End Select
    If shp.Rotation <> 0 Then
        MsgBox "Rotated shapes aren't supported (the tear uses the bounding-box edge). Set rotation to 0 first.", _
               vbExclamation, "Tear Gap": Exit Sub
    End If

    Dim ans As VbMsgBoxResult
    ans = MsgBox("Tear which side of the shape?" & vbCrLf & vbCrLf & _
                 "Yes = LEFT edge" & vbCrLf & "No  = RIGHT edge", _
                 vbYesNoCancel + vbQuestion, "Tear Gap")
    If ans = vbCancel Then Exit Sub

    Dim isTearA As Boolean, boundaryX As Single
    If ans = vbYes Then
        isTearA = False: boundaryX = shp.Left                 ' TearB (rot 180) cuts a LEFT edge
    Else
        isTearA = True: boundaryX = shp.Left + shp.Width       ' TearA cuts a RIGHT edge
    End If

    ApplyTear sld, shp.Name, isTearA, boundaryX, shp.Height / BAND_HEIGHT, shp.Top
    Exit Sub
Fail:
    MsgBox "Tear Gap failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Tear Gap"
End Sub

' Re-flow the timeline on the slide with weighting flipped, from the source data
' stashed at build time (no re-import, no hand-moving the torn band).
Public Sub ToggleWeightedTimeline(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = FindTimelineSlide()
    If sld Is Nothing Then
        MsgBox "No imported timeline found to toggle." & vbCrLf & vbCrLf & _
               "Create one with Make Timeline first (weighting can also be chosen at creation).", _
               vbExclamation, "Weighted Timeline": Exit Sub
    End If

    Dim ev() As TLEvent, n As Long, t As String, color As String
    Dim gaps As Boolean, multi As Boolean, wipe As Boolean, weighted As Boolean
    If Not LoadTimelineState(sld, ev, n, t, color, gaps, multi, wipe, weighted) Then
        MsgBox "Couldn't read the stored timeline data on this slide.", vbExclamation, "Weighted Timeline": Exit Sub
    End If

    weighted = Not weighted
    Dim summary As String
    If RenderTimeline(sld, ev, n, t, color, gaps, multi, wipe, weighted, summary) Then
        MsgBox "Weighted timeline turned " & IIf(weighted, "ON", "OFF") & "." & vbCrLf & vbCrLf & summary, _
               vbInformation, "Weighted Timeline"
    End If
    Exit Sub
Fail:
    MsgBox "Toggle failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Weighted Timeline"
End Sub

' --- source/state persistence (slide tags) so the toggle can re-flow in place ----
Private Sub StoreTimelineState(ByVal sld As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                               ByVal t As String, ByVal color As String, ByVal gaps As Boolean, _
                               ByVal multi As Boolean, ByVal wipe As Boolean, ByVal weighted As Boolean)
    Dim s As String, i As Long
    For i = 1 To n        ' strip the delimiter chars from text so records can't collide
        s = s & CStr(CDbl(ev(i).RawDate)) & Chr(1) _
              & Replace(Replace(ev(i).DateLabel, Chr(1), ""), Chr(2), "") & Chr(1) _
              & Replace(Replace(ev(i).Desc, Chr(1), ""), Chr(2), "") & Chr(1) _
              & Replace(Replace(ev(i).Bates, Chr(1), ""), Chr(2), "") & Chr(2)
    Next i
    SetTag sld, "TLDATA_EVENTS", s
    SetTag sld, "TLDATA_N", CStr(n)
    SetTag sld, "TLDATA_TYPE", t
    SetTag sld, "TLDATA_COLOR", color
    SetTag sld, "TLDATA_GAPS", IIf(gaps, "1", "0")
    SetTag sld, "TLDATA_MULTI", IIf(multi, "1", "0")
    SetTag sld, "TLDATA_WIPE", IIf(wipe, "1", "0")
    SetTag sld, "TLDATA_WEIGHTED", IIf(weighted, "1", "0")
End Sub

Private Function LoadTimelineState(ByVal sld As slide, ByRef ev() As TLEvent, ByRef n As Long, _
                                  ByRef t As String, ByRef color As String, ByRef gaps As Boolean, _
                                  ByRef multi As Boolean, ByRef wipe As Boolean, ByRef weighted As Boolean) As Boolean
    On Error GoTo CleanFail        ' any parse error on corrupt/edited state -> return False (friendly caller message)
    If Len(sld.Tags("TLDATA_N")) = 0 Then Exit Function
    n = CLng(sld.Tags("TLDATA_N"))
    If n < 1 Then Exit Function
    ReDim ev(1 To n)
    Dim recs() As String, f() As String, i As Long
    recs = Split(sld.Tags("TLDATA_EVENTS"), Chr(2))
    If UBound(recs) < n - 1 Then Exit Function       ' record count disagrees with TLDATA_N -> bail cleanly
    For i = 1 To n
        f = Split(recs(i - 1), Chr(1))
        If UBound(f) < 2 Then Exit Function
        ev(i).RawDate = CDate(CDbl(f(0)))
        ev(i).DateLabel = f(1)
        ev(i).Desc = f(2)
        If UBound(f) >= 3 Then ev(i).Bates = f(3)     ' tolerate older 3-field saved state
        ev(i).SortKey = CDbl(ev(i).RawDate)
        ev(i).OrigIndex = i
    Next i
    t = sld.Tags("TLDATA_TYPE")
    color = sld.Tags("TLDATA_COLOR")
    gaps = (sld.Tags("TLDATA_GAPS") = "1")
    multi = (sld.Tags("TLDATA_MULTI") = "1")
    wipe = (sld.Tags("TLDATA_WIPE") = "1")
    weighted = (sld.Tags("TLDATA_WEIGHTED") = "1")
    LoadTimelineState = True
    Exit Function
CleanFail:
    LoadTimelineState = False
End Function

Private Function FindTimelineSlide() As slide
    Dim a As slide, s As slide
    Set a = ActiveTargetSlide()
    If Not a Is Nothing Then
        If Len(a.Tags("TLDATA_N")) > 0 Then Set FindTimelineSlide = a: Exit Function
    End If
    For Each s In ActivePresentation.Slides
        If Len(s.Tags("TLDATA_N")) > 0 Then Set FindTimelineSlide = s: Exit Function
    Next s
End Function

Private Sub SetTag(ByVal sld As slide, ByVal tagKey As String, ByVal tagVal As String)
    On Error Resume Next
    sld.Tags.Delete tagKey
    On Error GoTo 0
    sld.Tags.Add tagKey, tagVal
End Sub

' ============================================================================
' DATEBAR LOOKUP / UNIT RECOGNITION  (shared by both makers + the edit tools)
' ============================================================================
' The bar is a grouped object named "Datebar NNN", tagged TLBar="1" and TLType=<unit>.
' Everything finds it by tag / name pattern (not the old literal "BottomBar"), so a
' rename or regroup no longer breaks recognition. Legacy "BottomBar"/"TopBar" still match.
Public Function IsDateBar(ByVal shp As Shape) As Boolean
    If ShapeTagVal(shp, "TLBar") = "1" Then IsDateBar = True: Exit Function
    If shp.Name Like "Datebar*" Then IsDateBar = True: Exit Function
    If shp.Name = "BottomBar" Or shp.Name = "TopBar" Then IsDateBar = True
End Function

Public Function FindDateBar(ByVal sld As slide) As Shape
    Dim shp As Shape
    For Each shp In sld.Shapes
        If ShapeTagVal(shp, "TLBar") = "1" Then Set FindDateBar = shp: Exit Function
    Next shp
    For Each shp In sld.Shapes
        If shp.Name Like "Datebar*" Then Set FindDateBar = shp: Exit Function
    Next shp
    For Each shp In sld.Shapes
        If shp.Name = "BottomBar" Then Set FindDateBar = shp: Exit Function
    Next shp
    For Each shp In sld.Shapes
        If shp.Name = "TopBar" Then Set FindDateBar = shp: Exit Function
    Next shp
End Function

Public Function NextDateBarName(ByVal sld As slide) As String
    Dim shp As Shape, mx As Long, suff As String, n As Long
    For Each shp In sld.Shapes
        If Left$(shp.Name, 8) = "Datebar " Then
            suff = Mid$(shp.Name, 9)
            If IsNumeric(suff) Then
                n = CLng(suff)
                If n > mx Then mx = n
            End If
        End If
    Next shp
    NextDateBarName = "Datebar " & Format$(mx + 1, "000")
End Function

Private Function ShapeTagVal(ByVal shp As Shape, ByVal key As String) As String
    On Error Resume Next
    ShapeTagVal = shp.Tags(key)
    On Error GoTo 0
End Function

' Robust unit: the bar's TLType tag -> stored slide state -> inferred from the bar labels.
Public Function BarUnit(ByVal sld As slide, ByVal bar As Shape) As String
    Dim t As String
    Dim ev() As TLEvent, n As Long, color As String
    Dim gaps As Boolean, multi As Boolean, wipe As Boolean, weighted As Boolean, tlSlide As slide
    If Not bar Is Nothing Then t = ShapeTagVal(bar, "TLType")
    If t <> "" Then BarUnit = t: Exit Function
    Set tlSlide = FindTimelineSlide()
    If Not tlSlide Is Nothing Then
        If LoadTimelineState(tlSlide, ev, n, t, color, gaps, multi, wipe, weighted) Then
            If t <> "" Then BarUnit = t: Exit Function
        End If
    End If
    If Not bar Is Nothing Then t = InferBarUnit(bar)
    BarUnit = t
End Function

' Read the bar's cell labels and guess the unit. Two-bar aware: a bottom row of bare
' day numbers -> Days, bare month names -> Months, checked before the coarser top signals.
Public Function InferBarUnit(ByVal bar As Shape) As String
    Dim labels As Collection, i As Long, s As String, anyText As Boolean
    Dim hasTime As Boolean, hasFullDate As Boolean, hasMonYear As Boolean
    Dim hasBareDay As Boolean, hasMonName As Boolean, allYears As Boolean
    Set labels = New Collection
    CollectBarLabels bar, labels
    allYears = True
    For i = 1 To labels.count
        s = Trim$(CStr(labels(i)))
        If s <> "" Then
            anyText = True
            If LabelIsTime(s) Then hasTime = True
            If LabelIsFullDate(s) Then hasFullDate = True
            If LabelIsMonthYear(s) Then hasMonYear = True
            If LabelIsBareDay(s) Then hasBareDay = True
            If LabelIsMonthName(s) Then hasMonName = True
            If Not LabelIsYear(s) Then allYears = False
        End If
    Next i
    If Not anyText Then Exit Function
    If hasTime Then InferBarUnit = "Hours": Exit Function
    If hasFullDate Then InferBarUnit = "Days": Exit Function
    If hasBareDay Then InferBarUnit = "Days": Exit Function
    If hasMonName Then InferBarUnit = "Months": Exit Function
    If hasMonYear Then InferBarUnit = "Months": Exit Function
    If allYears Then InferBarUnit = "Years"
End Function

Private Sub CollectBarLabels(ByVal shp As Shape, ByRef labels As Collection)
    Dim it As Shape, s As String
    If shp Is Nothing Then Exit Sub
    If shp.Type = msoGroup Then
        For Each it In shp.GroupItems
            CollectBarLabels it, labels
        Next it
    Else
        s = ""
        On Error Resume Next
        s = shp.TextFrame.TextRange.text
        On Error GoTo 0
        If Trim$(s) <> "" Then labels.Add s
    End If
End Sub

Private Function LabelIsTime(ByVal s As String) As Boolean
    s = UCase$(s)
    LabelIsTime = (InStr(s, ":") > 0) And (InStr(s, "AM") > 0 Or InStr(s, "PM") > 0)
End Function

Private Function LabelIsYear(ByVal s As String) As Boolean
    Dim y As Long
    s = Trim$(s)
    If Len(s) <> 4 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    y = CLng(s)
    LabelIsYear = (y >= 1000 And y <= 2999)
End Function

Private Function LabelIsBareDay(ByVal s As String) As Boolean
    Dim d As Long
    s = Trim$(s)
    If Len(s) < 1 Or Len(s) > 2 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    d = CLng(s)
    LabelIsBareDay = (d >= 1 And d <= 31)
End Function

Private Function LabelIsFullDate(ByVal s As String) As Boolean
    LabelIsFullDate = (InStr(s, "/") > 0) And IsDate(s)
End Function

Private Function LabelIsMonthYear(ByVal s As String) As Boolean
    Dim p() As String
    s = Trim$(s)
    If InStr(s, "/") > 0 Then Exit Function
    p = Split(s, " ")
    If UBound(p) <> 1 Then Exit Function
    LabelIsMonthYear = IsMonthWordTL(p(0)) And LabelIsYear(p(1))
End Function

Private Function LabelIsMonthName(ByVal s As String) As Boolean
    LabelIsMonthName = IsMonthWordTL(Trim$(s))
End Function

Private Function IsMonthWordTL(ByVal tok As String) As Boolean
    Select Case LCase$(Trim$(tok))
        Case "jan", "january", "feb", "february", "mar", "march", "apr", "april", _
             "may", "jun", "june", "jul", "july", "aug", "august", "sep", "sept", _
             "september", "oct", "october", "nov", "november", "dec", "december"
            IsMonthWordTL = True
    End Select
End Function

' ============================================================================
' COPILOT PROMPT  -  convert any source timeline into the import format
' ============================================================================
' The single source of truth for the prompt (used by the ribbon button and the
' "Convert with Copilot" sheet of the blank template). Tuned to exactly what
' ReadEvents/ParseDateCell accept, including the Bates column.
Private Function CopilotPromptText() As String
    Dim s As String, nl As String
    nl = vbCrLf
    s = "Convert the attached timeline into a fixed spreadsheet format. Extract EVERY dated event and output ONE table with exactly these four columns and this header row:" & nl & nl
    s = s & "Date | Time | Description | Bates" & nl & nl
    s = s & "Rules:" & nl
    s = s & "- One row per distinct event, sorted oldest to newest." & nl
    s = s & "- Date: use the most precise the source supports - a full date as M/D/YYYY (e.g. 3/15/2020); only a month and year as ""Month YYYY"" (e.g. April 2023); only a year as the 4-digit year (e.g. 2024). Never leave Date blank; omit events that have no date at all." & nl
    s = s & "- Time: only if a clock time is given (e.g. 2:30 PM); otherwise leave blank." & nl
    s = s & "- Description: a concise one-line summary of the event in plain text, with no line breaks or bullet characters." & nl
    s = s & "- Bates: the Bates number / document number / control number associated with the event if one is present (e.g. ABC-000123, or a range ABC-000123-000130); otherwise leave blank. Copy it exactly - do not alter or invent it." & nl
    s = s & "- For a date range, use the start date (one row)." & nl
    s = s & "- Do not invent events, dates, or Bates numbers. Output ONLY the table - no commentary and no extra columns."
    CopilotPromptText = s
End Function

' Ribbon: copy the Copilot prompt to the clipboard and show the 3 steps.
Public Sub CopilotTimelinePrompt(control As IRibbonControl)
    Dim copied As Boolean
    On Error Resume Next
    CopyToClipboard CopilotPromptText()
    copied = (Err.Number = 0)
    On Error GoTo 0

    Dim msg As String
    If copied Then
        msg = "A Copilot prompt has been copied to your clipboard." & vbCrLf & vbCrLf
    Else
        msg = "Couldn't reach the clipboard automatically - you can copy the prompt from the " & _
              """Convert with Copilot"" sheet of a blank template instead (Make Timeline -> No)." & vbCrLf & vbCrLf
    End If
    msg = msg & "1.  In Microsoft Copilot (Chat, or Copilot in Word), attach or open your source timeline (PDF / Word / Excel)." & vbCrLf & _
                "2.  Paste the prompt (Ctrl+V) and send it." & vbCrLf & _
                "3.  Paste Copilot's table into a blank timeline template, save it, then run Make Timeline on that file."
    MsgBox msg, vbInformation, "Convert a Timeline with Copilot"
End Sub

' Put text on the Windows clipboard (late-bound MSForms.DataObject; no project reference needed).
Private Sub CopyToClipboard(ByVal s As String)
    Dim dobj As Object
    Set dobj = GetObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")   ' MSForms.DataObject
    dobj.SetText s
    dobj.PutInClipboard
End Sub

' ============================================================================
' SPACE EVENLY WITH LEADERS
' ============================================================================
' Distribute the selected entries with EQUAL GAPS, measured on their BOX bounds
' only (the leader lines are ignored so they can't skew the spacing the way
' PowerPoint's native Distribute does). Each whole entry - box + leader - moves
' together. Axis is auto-chosen as the one the boxes are more spread along.
Public Sub SpaceEntriesEvenly(control As IRibbonControl)
    On Error GoTo Fail
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select the entries you want to space (at least 3).", vbExclamation, "Space Evenly with Leaders"
        Exit Sub
    End If
    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    Dim cnt As Long
    cnt = sr.count
    If cnt < 3 Then
        MsgBox "Select at least 3 entries to space them evenly.", vbExclamation, "Space Evenly with Leaders"
        Exit Sub
    End If

    Dim bTop() As Single, bLeft() As Single, bH() As Single, bW() As Single, idx() As Long
    ReDim bTop(1 To cnt): ReDim bLeft(1 To cnt): ReDim bH(1 To cnt): ReDim bW(1 To cnt): ReDim idx(1 To cnt)
    Dim i As Long, topY As Single, leftX As Single, botY As Single, rgtX As Single
    For i = 1 To cnt
        BoxRectOf sr(i), topY, leftX, botY, rgtX
        bTop(i) = topY: bLeft(i) = leftX: bH(i) = botY - topY: bW(i) = rgtX - leftX
        idx(i) = i
    Next i

    ' choose the axis the boxes are more spread along
    Dim minT As Single, maxB As Single, minL As Single, maxR As Single
    minT = bTop(1): maxB = bTop(1) + bH(1): minL = bLeft(1): maxR = bLeft(1) + bW(1)
    For i = 2 To cnt
        If bTop(i) < minT Then minT = bTop(i)
        If bTop(i) + bH(i) > maxB Then maxB = bTop(i) + bH(i)
        If bLeft(i) < minL Then minL = bLeft(i)
        If bLeft(i) + bW(i) > maxR Then maxR = bLeft(i) + bW(i)
    Next i
    Dim vertical As Boolean
    vertical = ((maxB - minT) >= (maxR - minL))

    ' sort idx() by box position along the chosen axis (insertion sort)
    Dim a As Long, keyIdx As Long, pa As Single, pb As Single
    For a = 2 To cnt
        keyIdx = idx(a): i = a - 1
        Do While i >= 1
            If vertical Then
                pa = bTop(idx(i)): pb = bTop(keyIdx)
            Else
                pa = bLeft(idx(i)): pb = bLeft(keyIdx)
            End If
            If pa <= pb Then Exit Do
            idx(i + 1) = idx(i): i = i - 1
        Loop
        idx(i + 1) = keyIdx
    Next a

    ' equal-gap distribution; move each whole entry so its BOX lands on target
    Dim sumSize As Single, gap As Single, cursor As Single, j As Long
    sumSize = 0
    For i = 1 To cnt
        If vertical Then sumSize = sumSize + bH(idx(i)) Else sumSize = sumSize + bW(idx(i))
    Next i
    If vertical Then
        gap = ((maxB - minT) - sumSize) / (cnt - 1)
        cursor = minT
    Else
        gap = ((maxR - minL) - sumSize) / (cnt - 1)
        cursor = minL
    End If
    For i = 1 To cnt
        j = idx(i)
        If vertical Then
            sr(j).Top = sr(j).Top + (cursor - bTop(j))
            cursor = cursor + bH(j) + gap
        Else
            sr(j).Left = sr(j).Left + (cursor - bLeft(j))
            cursor = cursor + bW(j) + gap
        End If
    Next i
    Exit Sub
Fail:
    MsgBox "Space Evenly failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Space Evenly with Leaders"
End Sub

' Bounding rect of an entry's BOX portion (every child except the LeadingLine);
' falls back to the whole shape if it isn't a group or has only a leader.
Private Sub BoxRectOf(ByVal shp As Shape, ByRef topY As Single, ByRef leftX As Single, _
                      ByRef botY As Single, ByRef rgtX As Single)
    Dim it As Shape, got As Boolean
    If shp.Type = msoGroup Then
        For Each it In shp.GroupItems
            If Not (it.Name Like "LeadingLine*") Then
                If Not got Then
                    topY = it.Top: leftX = it.Left: botY = it.Top + it.Height: rgtX = it.Left + it.Width
                    got = True
                Else
                    If it.Top < topY Then topY = it.Top
                    If it.Left < leftX Then leftX = it.Left
                    If it.Top + it.Height > botY Then botY = it.Top + it.Height
                    If it.Left + it.Width > rgtX Then rgtX = it.Left + it.Width
                End If
            End If
        Next it
    End If
    If Not got Then
        topY = shp.Top: leftX = shp.Left: botY = shp.Top + shp.Height: rgtX = shp.Left + shp.Width
    End If
End Sub

' ============================================================================
' Z AXIS FIX  -  put every leader line behind every entry box
' ============================================================================
' Reassesses the entries currently on the active slide (the user may have moved or
' added some) and restacks the whole entry GROUPS so each lower entry's leader line
' tucks behind the boxes above it. Groups are preserved (no ungrouping).
Public Sub ZAxisFix(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveTargetSlide()
    If sld Is Nothing Then
        MsgBox "Open a presentation and select a slide first.", vbExclamation, "Z Axis Fix"
        Exit Sub
    End If
    If sld.Shapes.count = 0 Then
        MsgBox "No timeline entries found on this slide.", vbExclamation, "Z Axis Fix"
        Exit Sub
    End If

    Dim shp As Shape, ents() As Shape, tops() As Single, cnt As Long
    Dim topY As Single, leftX As Single, botY As Single, rgtX As Single
    ReDim ents(1 To sld.Shapes.count)
    ReDim tops(1 To sld.Shapes.count)
    cnt = 0
    For Each shp In sld.Shapes
        If IsEntryGroup(shp) Then
            BoxRectOf shp, topY, leftX, botY, rgtX
            cnt = cnt + 1
            Set ents(cnt) = shp
            tops(cnt) = topY
        End If
    Next shp
    If cnt = 0 Then
        MsgBox "No timeline entries (groups with a leader line) found on this slide.", vbExclamation, "Z Axis Fix"
        Exit Sub
    End If

    ' sort entries by box top, topmost first (insertion sort)
    Dim i As Long, j As Long, keyT As Single, keyS As Shape
    For i = 2 To cnt
        keyT = tops(i): Set keyS = ents(i)
        j = i - 1
        Do While j >= 1
            If tops(j) <= keyT Then Exit Do
            tops(j + 1) = tops(j): Set ents(j + 1) = ents(j)
            j = j - 1
        Loop
        tops(j + 1) = keyT: Set ents(j + 1) = keyS
    Next i

    ' SendToBack top-first -> topmost entry ends frontmost among entries, the bar/other
    ' content stays in front, and each lower entry's longer leader line sits behind the
    ' boxes above it.
    For i = 1 To cnt
        ents(i).ZOrder msoSendToBack
    Next i

    MsgBox cnt & " entr" & IIf(cnt = 1, "y", "ies") & " restacked - leader lines are now behind the entry boxes.", _
           vbInformation, "Z Axis Fix"
    Exit Sub
Fail:
    MsgBox "Z Axis Fix failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Z Axis Fix"
End Sub

' ============================================================================
' DATE SNAP  -  align each entry's leader line to its date on the datebar
' ============================================================================
' Moves entries HORIZONTALLY only so each leader line lands where its date falls on
' the datebar. Re-reads the CURRENT datebar cell positions (the bar may have been
' moved/resized) and the CURRENT entries (some may be new/moved). Needs a timeline
' made by Make Timeline (which stores the date unit).
' Two modes. Group Move shifts the whole entry so its leader lands on the date. Line Nudge
' moves only the leader WITHIN the entry's box bounds; entries whose date is past the box edge
' are clamped + reported, with an offer to group-move those the rest of the way.
Public Sub DateSnapGroupMove(control As IRibbonControl)
    DateSnapCore True
End Sub

Public Sub DateSnapLineNudge(control As IRibbonControl)
    DateSnapCore False
End Sub

Private Sub DateSnapCore(ByVal groupMove As Boolean)
    Dim title As String
    title = IIf(groupMove, "Date Snap (Group Move)", "Date Snap (Line Nudge)")
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveTargetSlide()
    If sld Is Nothing Then
        MsgBox "Open a presentation and select a slide first.", vbExclamation, title
        Exit Sub
    End If

    ' Robust unit: bar tag -> stored state -> inferred from the bar -> ask the user.
    Dim bar As Shape, t As String
    Set bar = FindDateBar(sld)
    t = BarUnit(sld, bar)
    If t = "" Then t = AskTimelineUnit()
    If t = "" Then Exit Sub

    ' map each datebar cell's label -> its CURRENT left/width (absolute) so a moved/resized bar still snaps
    Dim cellLefts As Object, cellWidths As Object
    Set cellLefts = CreateObject("Scripting.Dictionary")
    Set cellWidths = CreateObject("Scripting.Dictionary")
    CollectBandCells bar, cellLefts, cellWidths
    If cellLefts.count = 0 Then
        MsgBox "No datebar cells found on this slide (couldn't find a 'Datebar' / 'BottomBar').", vbExclamation, title
        Exit Sub
    End If

    Dim shp As Shape, snapped As Long, missed As Long, d As Date, lab As String
    Dim innerL As Single, innerW As Single, targetX As Single, ln As Shape
    Dim topY As Single, leftX As Single, botY As Single, rgtX As Single, cx As Single, frac As Double
    Dim shortLabels As Collection, shortShapes As Collection, shortTargets As Collection
    Set shortLabels = New Collection: Set shortShapes = New Collection: Set shortTargets = New Collection

    For Each shp In sld.Shapes
        If IsEntryGroup(shp) Then
            If TryEntryDate(shp, d) Then
                If MatchBarCell(d, t, cellLefts, lab, frac) Then
                    innerL = cellLefts(lab) + COL_PAD
                    innerW = cellWidths(lab) - 2 * COL_PAD
                    If innerW < 0 Then innerW = 0
                    targetX = innerL + frac * innerW
                    Set ln = LeadingLineOf(shp)
                    If ln Is Nothing Then
                        missed = missed + 1
                    ElseIf groupMove Then
                        shp.Left = shp.Left + (targetX - ln.Left)
                        snapped = snapped + 1
                    Else
                        BoxRectOf shp, topY, leftX, botY, rgtX     ' Line Nudge: clamp to the box bounds
                        cx = targetX
                        If cx < leftX Then cx = leftX
                        If cx > rgtX Then cx = rgtX
                        ln.Left = cx
                        snapped = snapped + 1
                        If cx <> targetX Then
                            shortLabels.Add Format$(d, "mmm d, yyyy")
                            shortShapes.Add shp
                            shortTargets.Add targetX
                        End If
                    End If
                Else
                    missed = missed + 1
                End If
            Else
                missed = missed + 1
            End If
        End If
    Next shp

    Dim msg As String
    msg = snapped & " entr" & IIf(snapped = 1, "y", "ies") & " " & IIf(groupMove, "snapped", "nudged") & " to the datebar."
    If missed > 0 Then msg = msg & vbCrLf & missed & " skipped (couldn't read the date, or its unit isn't on the bar)."

    If (Not groupMove) And shortLabels.count > 0 Then
        Dim listMsg As String, k As Long, resp As VbMsgBoxResult, gm As Long
        For k = 1 To shortLabels.count
            listMsg = listMsg & vbCrLf & "  - " & shortLabels(k)
        Next k
        resp = MsgBox(shortLabels.count & " entr" & IIf(shortLabels.count = 1, "y", "ies") & _
              " couldn't fully reach the date (the line hit the edge of the entry box):" & vbCrLf & listMsg & _
              vbCrLf & vbCrLf & "Group-move those entries the rest of the way?", _
              vbYesNo + vbQuestion, title)
        If resp = vbYes Then
            For k = 1 To shortShapes.count
                Set shp = shortShapes(k)
                Set ln = LeadingLineOf(shp)
                If Not ln Is Nothing Then
                    BoxRectOf shp, topY, leftX, botY, rgtX
                    ln.Left = (leftX + rgtX) / 2                 ' recenter the leader...
                    shp.Left = shp.Left + (shortTargets(k) - ln.Left)   ' ...then group-move to the date
                    gm = gm + 1
                End If
            Next k
            msg = msg & vbCrLf & gm & " group-moved to finish."
        End If
    End If

    MsgBox msg, vbInformation, title
    Exit Sub
Fail:
    MsgBox title & " failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, title
End Sub

' Last-resort unit prompt when the bar + stored data can't tell us.
Private Function AskTimelineUnit() As String
    Dim s As String
    s = InputBox("Couldn't auto-detect the timeline's date unit." & vbCrLf & _
                 "Enter one of: Years, Months, Days, Hours", "Date Snap")
    Select Case LCase$(Trim$(s))
        Case "years": AskTimelineUnit = "Years"
        Case "months": AskTimelineUnit = "Months"
        Case "days": AskTimelineUnit = "Days"
        Case "hours": AskTimelineUnit = "Hours"
    End Select
End Function

' Match an entry's date to a datebar cell. Tries the unit's own cell first, then progressively
' coarser cells - so a two-bar bar (whose fine row has ambiguous bare labels like "15"/"Jan")
' snaps against the unique coarser cell ("Mar 2023" / "2023") using the fraction within it.
Private Function MatchBarCell(ByVal d As Date, ByVal t As String, ByVal cellLefts As Object, _
                              ByRef matchedLabel As String, ByRef frac As Double) As Boolean
    Dim chain As Variant, i As Long, u As String, lab As String, dk As String
    ' Precise: a cell tagged with this exact date (the finest unit) - exact for single + two-bar.
    dk = DateKeyOf(UnitStartOf(d, t))
    If cellLefts.Exists(dk) Then
        matchedLabel = dk
        frac = UnitFracOf(d, t)
        MatchBarCell = True
        Exit Function
    End If
    ' Fallback (hand-built bars with no date tags): match by label, coarsest unique cell.
    chain = CoarserChain(t)
    For i = LBound(chain) To UBound(chain)
        u = CStr(chain(i))
        lab = Trim$(GetSegmentLabel(u, UnitStartOf(d, u)))
        If cellLefts.Exists(lab) Then
            matchedLabel = lab
            frac = UnitFracOf(d, u)
            MatchBarCell = True
            Exit Function
        End If
    Next i
End Function

Private Function CoarserChain(ByVal t As String) As Variant
    Select Case t
        Case "Hours": CoarserChain = Array("Hours", "Days")
        Case "Days": CoarserChain = Array("Days", "Months")
        Case "Months": CoarserChain = Array("Months", "Years")
        Case Else: CoarserChain = Array("Years")
    End Select
End Function

' Canonical key for a cell's date (used to match entries to date-tagged cells precisely).
Private Function DateKeyOf(ByVal d As Date) As String
    DateKeyOf = "D:" & Format$(d, "yyyy-mm-dd-hh")
End Function

' ============================================================================
' PLACE ENTRY ON TIMELINE  (Make Entry -> "Confirm & Place")
' ============================================================================
' Drop a freshly-made entry under the datebar at its date, then re-space ALL dated entries
' evenly across the bar (date order) and line-nudge each leader to its date.
Public Sub PlaceEntryOnTimeline(ByVal sld As slide, ByVal newEntry As Shape, ByVal d As Date)
    Const ROW_TOP As Single = 82.8           ' 1.15" from the top, just under the band
    Dim bar As Shape, t As String, db As Shape
    On Error GoTo Fail
    Set bar = FindDateBar(sld)
    t = BarUnit(sld, bar)
    If bar Is Nothing Or t = "" Then
        MsgBox "Entry created, but no datebar was found on this slide to place it on.", _
               vbInformation, "Make Entry"
        Exit Sub
    End If
    Set db = FindTaggedDescendant(newEntry, "Date Box")
    If Not db Is Nothing Then newEntry.Top = newEntry.Top + (ROW_TOP - db.Top)
    ReflowTimeline sld, bar, t
    Exit Sub
Fail:
    MsgBox "The entry was created but couldn't be auto-placed:" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Make Entry"
End Sub

' Re-space by date: GROUP-MOVE every dated entry so its leader sits exactly on its date.
' Horizontal only - each entry keeps its vertical position (lanes preserved). A group move
' (not a line nudge) is what lets the leader actually reach the date.
Private Sub ReflowTimeline(ByVal sld As slide, ByVal bar As Shape, ByVal t As String)
    Dim cellLefts As Object, cellWidths As Object, shp As Shape, dd As Date
    Dim lab As String, frac As Double, innerL As Single, innerW As Single, targetX As Single, ln As Shape
    Set cellLefts = CreateObject("Scripting.Dictionary")
    Set cellWidths = CreateObject("Scripting.Dictionary")
    CollectBandCells bar, cellLefts, cellWidths
    For Each shp In sld.Shapes
        If IsEntryGroup(shp) Then
            If TryEntryDate(shp, dd) Then
                If MatchBarCell(dd, t, cellLefts, lab, frac) Then
                    innerL = cellLefts(lab) + COL_PAD
                    innerW = cellWidths(lab) - 2 * COL_PAD
                    If innerW < 0 Then innerW = 0
                    targetX = innerL + frac * innerW
                    Set ln = LeadingLineOf(shp)
                    If Not ln Is Nothing Then shp.Left = shp.Left + (targetX - ln.Left)
                End If
            End If
        End If
    Next shp
End Sub

' --- shared helpers for the entry-edit commands -----------------------------
' An entry = a group with a direct LeadingLine child (Make Entry + imported entries).
Private Function IsEntryGroup(ByVal shp As Shape) As Boolean
    Dim it As Shape
    If shp.Type <> msoGroup Then Exit Function
    For Each it In shp.GroupItems
        If it.Name Like "LeadingLine*" Then IsEntryGroup = True: Exit Function
    Next it
End Function

Private Function LeadingLineOf(ByVal entryGroup As Shape) As Shape
    Dim it As Shape
    If entryGroup.Type <> msoGroup Then Exit Function
    For Each it In entryGroup.GroupItems
        If it.Name Like "LeadingLine*" Then Set LeadingLineOf = it: Exit Function
    Next it
End Function

' Recurse a container, recording each datebar cell by its label -> current Left/Width (absolute).
' Map each datebar cell's label -> current Left/Width (absolute) by walking the bar group's
' leaf cells. Works for both rectangle-cell bars (Timeline maker) and converted-table textbox
' bars (DateBar maker); recurses nested groups (e.g. a two-bar group).
Private Sub CollectBandCells(ByVal bar As Shape, ByVal cellLefts As Object, ByVal cellWidths As Object)
    Dim it As Shape, lab As String, cd As String, dk As String
    If bar Is Nothing Then Exit Sub
    If bar.Type = msoGroup Then
        For Each it In bar.GroupItems
            CollectBandCells it, cellLefts, cellWidths
        Next it
    Else
        lab = ""
        On Error Resume Next
        lab = Trim$(bar.TextFrame.TextRange.text)
        On Error GoTo 0
        If lab <> "" Then
            cellLefts(lab) = bar.Left
            cellWidths(lab) = bar.Width
        End If
        ' also key by the cell's tagged date (precise, disambiguates two-bar bare labels)
        cd = ShapeTagVal(bar, "TLCellDate")
        If cd <> "" Then
            If IsNumeric(cd) Then
                dk = DateKeyOf(CDate(CDbl(cd)))
                cellLefts(dk) = bar.Left
                cellWidths(dk) = bar.Width
            End If
        End If
    End If
End Sub

' Read + parse an entry's date from its "Date Box" text (strips an appended time).
Private Function TryEntryDate(ByVal entryGroup As Shape, ByRef d As Date) As Boolean
    Dim db As Shape, s As String, ev As TLEvent, fd As String
    ' prefer the full date stashed at build time (the visible label may omit the year)
    On Error Resume Next
    fd = entryGroup.Tags("TLFullDate")
    On Error GoTo 0
    If IsNumeric(fd) Then
        d = CDate(CDbl(fd))
        TryEntryDate = True
        Exit Function
    End If
    ' otherwise parse the Date Box text (Make-Entry / manual entries carry a full date there)
    Set db = FindTaggedDescendant(entryGroup, "Date Box")
    If db Is Nothing Then Exit Function
    On Error Resume Next
    s = db.TextFrame.TextRange.text
    On Error GoTo 0
    s = CleanDateText(s)
    If s = "" Then Exit Function
    If ParseDateCell(s, ev) Then
        d = ev.RawDate
        TryEntryDate = True
    End If
End Function

Private Function CleanDateText(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, vbCr): If p > 0 Then s = Left$(s, p - 1)   ' drop a wrapped 2nd line
    p = InStr(s, vbLf): If p > 0 Then s = Left$(s, p - 1)
    Do While InStr(s, "  ") > 0                             ' collapse the "  <time>" gap but KEEP the time (Hours mode needs it)
        s = Replace(s, "  ", " ")
    Loop
    CleanDateText = Trim$(s)
End Function

Private Function FindTaggedDescendant(ByVal shp As Shape, ByVal tagVal As String) As Shape
    Dim it As Shape, found As Shape
    On Error Resume Next
    If shp.Tags("GroupStyle") = tagVal Then Set FindTaggedDescendant = shp: Exit Function
    On Error GoTo 0
    If shp.Type = msoGroup Then
        For Each it In shp.GroupItems
            Set found = FindTaggedDescendant(it, tagVal)
            If Not found Is Nothing Then Set FindTaggedDescendant = found: Exit Function
        Next it
    End If
End Function

' ============================================================================
' PLACE ENTRY AT TOP  -  pull selected entries up under the bar + snap leaders
' ============================================================================
' Moves each selected entry vertically so its date box top sits at 1.15" from the
' slide top, then snaps that entry's leader line to the BottomBar (top -> bar middle,
' bottom kept at the entry) - the same snap "Snap Entry Lines" uses.
Public Sub PlaceEntryAtTop(control As IRibbonControl)
    On Error GoTo Fail
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select one or more timeline entries first.", vbExclamation, "Place Entry at Top"
        Exit Sub
    End If
    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    Dim sld As slide
    Set sld = sr(1).Parent

    ' the datebar to snap leaders to (may have been moved -> use its current position)
    Dim bar As Shape
    Set bar = FindDateBar(sld)
    If bar Is Nothing Then
        MsgBox "No datebar found on this slide to snap leaders to.", vbExclamation, "Place Entry at Top"
        Exit Sub
    End If
    Dim middleY As Single
    middleY = bar.Top + bar.Height / 2

    Const TARGET_TOP As Single = 82.8           ' 1.15" from the slide top (1.15 * 72)

    Dim k As Long, shp As Shape, db As Shape, ln As Shape, lineBottomY As Single, placed As Long
    placed = 0
    For k = 1 To sr.count
        Set shp = sr(k)
        If IsEntryGroup(shp) Then
            Set db = FindTaggedDescendant(shp, "Date Box")
            If Not db Is Nothing Then
                shp.Top = shp.Top + (TARGET_TOP - db.Top)     ' date box top -> 1.15" (vertical only)
                Set ln = LeadingLineOf(shp)
                If Not ln Is Nothing Then
                    lineBottomY = ln.Top + ln.Height
                    ln.Top = middleY                          ' snap leader top to the bar middle...
                    ln.Height = lineBottomY - middleY         ' ...keeping its bottom at the entry
                End If
                placed = placed + 1
            End If
        End If
    Next k

    If placed = 0 Then
        MsgBox "None of the selected shapes are timeline entries (a group with a date box + leader line).", _
               vbExclamation, "Place Entry at Top"
        Exit Sub
    End If
    MsgBox placed & " entr" & IIf(placed = 1, "y", "ies") & " placed at the top; leaders snapped to the bar.", _
           vbInformation, "Place Entry at Top"
    Exit Sub
Fail:
    MsgBox "Place Entry at Top failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Place Entry at Top"
End Sub

' ============================================================================
' AUTO CENTER  (Formats)  -  align entry boxes by how many lines they wrap to
' ============================================================================
' Every "Entry Box" on the active slide: 1-2 wrapped lines -> centered; 3+ -> left.
Public Sub AutoCenter(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveTargetSlide()
    If sld Is Nothing Then
        MsgBox "Open a presentation and select a slide first.", vbExclamation, "Auto Center"
        Exit Sub
    End If
    Dim changed As Long
    changed = 0
    AutoCenterRecurse sld.Shapes, changed
    MsgBox changed & " entry box(es) re-aligned (1-2 lines centered, 3+ left-aligned).", vbInformation, "Auto Center"
    Exit Sub
Fail:
    MsgBox "Auto Center failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Auto Center"
End Sub

Private Sub AutoCenterRecurse(ByVal shapesColl As Object, ByRef changed As Long)
    Dim shp As Shape, lc As Long
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            AutoCenterRecurse shp.GroupItems, changed
        ElseIf IsEntryBox(shp) Then
            lc = LineCountOf(shp)
            If lc > 0 Then
                If lc <= 2 Then
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                Else
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
                End If
                changed = changed + 1
            End If
        End If
    Next shp
End Sub

Private Function IsEntryBox(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If shp.Tags("GroupStyle") <> "Entry Box" Then Exit Function
    If Not shp.HasTextFrame Then Exit Function
    If Not shp.TextFrame.HasText Then Exit Function
    IsEntryBox = True
End Function

' Number of DISPLAYED (wrapped) lines, like .Paragraphs.Count but counting wrap.
Private Function LineCountOf(ByVal shp As Shape) As Long
    On Error Resume Next
    LineCountOf = shp.TextFrame.TextRange.Lines.count
    On Error GoTo 0
End Function
