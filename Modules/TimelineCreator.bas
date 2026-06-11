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
Private Const BAND_TOP      As Single = 72
Private Const BAND_HEIGHT   As Single = 30
Private Const BAND_GAP      As Single = 24
Private Const BOX_WIDTH     As Single = 168
Private Const DATE_HEIGHT   As Single = 21
Private Const ROW_GAP       As Single = 22
Private Const LINE_INSET    As Single = 18
Private Const COL_PAD       As Single = 8
Private Const MIN_COL_W     As Single = 132

Private Const FS_DATE       As Single = 12
Private Const FS_DESC       As Single = 10
Private Const FS_LABEL      As Single = 14
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
    Dim colDate As Long, colTime As Long, colDesc As Long
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
        ev.DateLabel = Format$(d, "mmm d, yyyy")
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
        ev.DateLabel = Format$(d, "mmm d, yyyy")
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

    ' tallest stacked column (for vertical autofit) + event count per column (for weighting)
    Dim ci As Long, maxH As Single, colH As Single, i As Long
    Dim cnt() As Long
    ReDim cnt(1 To pn)
    For ci = 1 To pn
        colH = 0
        For i = 1 To n
            If ev(i).UnitStart = pageCols(ci) Then
                colH = colH + DATE_HEIGHT + ev(i).DescH + ROW_GAP
                cnt(ci) = cnt(ci) + 1
            End If
        Next i
        If colH > maxH Then maxH = colH
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
    bandBottom = BAND_TOP + BAND_HEIGHT * sc

    DrawBar sld, pageCols, pn, colWArr, colLArr, sc, t, barColor

    Dim drawn As Long, animSeq As Long
    For ci = 1 To pn
        drawn = drawn + DrawColumn(sld, ev, n, pageCols(ci), colLArr(ci), colWArr(ci), bandBottom, sc, doWipe, animSeq)
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
        Set sh = sld.Shapes.AddShape(msoShapeRectangle, colLArr(ci), BAND_TOP, colWArr(ci), BAND_HEIGHT * sc)
        StyleBandCell sh, GetSegmentLabel(t, pageCols(ci)), sc, isGreen
        sh.Name = SHAPE_PREFIX & "_BandCell_" & ci & "_" & sld.SlideIndex
        cellNames(ci) = sh.Name
    Next ci

    ' Tear cells adjacent to a gap (compact mode only; contiguous has no gaps).
    ' boundaryX = the shared edge between column ci and ci+1 = colLArr(ci+1).
    For ci = 1 To pn - 1
        If DateAdd(IntervalCode(t), 1, pageCols(ci)) <> pageCols(ci + 1) Then
            ApplyTear sld, cellNames(ci), True, colLArr(ci + 1), sc       ' TearA: left cell's right edge
            ApplyTear sld, cellNames(ci + 1), False, colLArr(ci + 1), sc  ' TearB (rot 180): right cell's left edge
        End If
    Next ci

    Dim grp As Shape
    If pn = 1 Then
        Set grp = sld.Shapes(cellNames(1))
    Else
        Set grp = sld.Shapes.Range(cellNames).Group
    End If
    grp.Name = "BottomBar"
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
            .Font.Size = FS_LABEL * sc
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

' One unit-column of stacked entries with date-accurate leader lines.
Private Function DrawColumn(ByVal sld As slide, ByRef ev() As TLEvent, ByVal n As Long, _
                            ByVal colUnit As Date, ByVal colLeft As Single, ByVal colW As Single, _
                            ByVal bandBottom As Single, ByVal sc As Single, _
                            ByVal doWipe As Boolean, ByRef animSeq As Long) As Long
    Dim innerL As Single, innerR As Single, boxW As Single
    innerL = colLeft + COL_PAD
    innerR = colLeft + colW - COL_PAD
    boxW = BOX_WIDTH * sc
    If boxW > (colW - 2 * COL_PAD) Then boxW = colW - 2 * COL_PAD
    If boxW < 1 Then boxW = 1                        ' narrow weighted columns must stay positive

    Dim cursorY As Single, prevLeft As Single, drawn As Long, i As Long
    cursorY = bandBottom + BAND_GAP * sc
    prevLeft = innerL

    For i = 1 To n
        If ev(i).UnitStart = colUnit Then
            Dim dateX As Single, boxLeft As Single, boxesH As Single, grp As Shape
            dateX = innerL + ClampD(ev(i).UnitFrac, 0, 1) * (innerR - innerL)
            boxLeft = dateX - LINE_INSET * sc
            If boxLeft < prevLeft Then boxLeft = prevLeft
            If boxLeft > innerR - boxW Then boxLeft = innerR - boxW
            If boxLeft < innerL Then boxLeft = innerL
            prevLeft = boxLeft

            ' the real Make-Entry creator: one group, exact naming, black entry text,
            ' with a date-accurate leader line from the band bottom.
            Set grp = CreateTimelineEntry(sld, ev(i).DateLabel, ev(i).Desc, boxLeft, cursorY, boxW, _
                                          dateX, bandBottom, True, sc, boxesH)
            grp.Tags.Add "TLENTRY", "1"
            grp.ZOrder msoSendToBack        ' earlier (upper) entries stay on top -> leader lines tuck behind the boxes above
            If doWipe Then
                animSeq = animSeq + 1
                AddWipe sld, grp, True
            End If

            cursorY = cursorY + boxesH + ROW_GAP * sc
            drawn = drawn + 1
        End If
    Next i
    DrawColumn = drawn
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
        If (sld.Shapes(i).Tags("TLENTRY") = "1") Or (sld.Shapes(i).Name = "BottomBar") _
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
           "Fill in the Date and Description columns (Time optional), then run Make Timeline again.", _
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
              & Replace(Replace(ev(i).Desc, Chr(1), ""), Chr(2), "") & Chr(2)
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
