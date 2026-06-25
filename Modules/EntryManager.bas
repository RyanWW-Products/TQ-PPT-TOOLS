Attribute VB_Name = "EntryManager"
Option Explicit

' ============================================================================
' EntryManager - the engine behind the Calendar "Entry Manager" (frmEntryManager).
' ============================================================================
' Entries stay rectangle shapes named "Entry_<day>_<A|B|C|D>" (max 4/day, the cap
' is SHARED across groups) so the existing calendar positioning keeps working.
' Group identity is a per-shape Tag "EntryGroup"; "EntryDay" is a hardening tag so
' the day survives a hand-rename. Color comes from the group registry (EntryGroups),
' never insertion order. Reuses NextAvailableSuffix / CountDayEntries / ShapeExists
' from CreateEntry.bas.

' ---- entry point -----------------------------------------------------------
Public Sub OpenEntryManager(control As IRibbonControl)
    If Application.Presentations.count = 0 Then
        MsgBox "Please open a presentation first.", vbExclamation, "Entry Manager"
        Exit Sub
    End If
    If Not HasAnyCalendarSlide() Then
        MsgBox "No calendar slides found. Use 'Make Calendar' first.", vbExclamation, "Entry Manager"
        Exit Sub
    End If
    frmEntryManager.Show
End Sub

' ---- slide helpers ---------------------------------------------------------
Public Function GetSlideTag(sld As slide, ByVal name As String) As String
    On Error Resume Next
    GetSlideTag = sld.Tags(name)
    On Error GoTo 0
End Function

Public Function CalYearOf(sld As slide) As Long
    CalYearOf = CLng(Val(GetSlideTag(sld, "CalendarYear")))
End Function

Public Function CalMonthOf(sld As slide) As Long
    CalMonthOf = CLng(Val(GetSlideTag(sld, "CalendarMonth")))
End Function

Public Function IsCalendarSlide(sld As slide) As Boolean
    IsCalendarSlide = (CalYearOf(sld) > 0 And CalMonthOf(sld) >= 1 And CalMonthOf(sld) <= 12)
End Function

Public Function HasAnyCalendarSlide() As Boolean
    Dim sld As slide
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then HasAnyCalendarSlide = True: Exit Function
    Next sld
End Function

Public Function FindMonthSlide(ByVal yr As Long, ByVal mo As Long) As slide
    Dim sld As slide
    For Each sld In ActivePresentation.Slides
        If CalYearOf(sld) = yr And CalMonthOf(sld) = mo Then Set FindMonthSlide = sld: Exit Function
    Next sld
End Function

Public Function FirstCalendarSlide() As slide
    Dim sld As slide
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then Set FirstCalendarSlide = sld: Exit Function
    Next sld
End Function

Public Function DaysInMonth(ByVal yr As Long, ByVal mo As Long) As Long
    DaysInMonth = day(DateAdd("m", 1, DateSerial(yr, mo, 1)) - 1)
End Function

' ---- entry-shape helpers ---------------------------------------------------
Public Function IsEntryShape(shp As Shape) As Boolean
    IsEntryShape = (Left(shp.name, 6) = "Entry_")
End Function

Public Function DayOfEntry(shp As Shape) As Long
    Dim nm As String, p1 As Long, p2 As Long, ds As String
    nm = shp.name
    p1 = InStr(nm, "_")
    p2 = InStr(p1 + 1, nm, "_")
    If p1 > 0 And p2 > p1 Then
        ds = Mid(nm, p1 + 1, p2 - p1 - 1)
        If IsNumeric(ds) Then DayOfEntry = CLng(ds): Exit Function
    End If
    DayOfEntry = CLng(Val(GetShapeTag(shp, "EntryDay")))
End Function

Public Function GroupOfEntry(shp As Shape) As String
    Dim g As String
    g = GetShapeTag(shp, "EntryGroup")
    If g = "" Then g = DEFAULT_GROUP
    GroupOfEntry = g
End Function

Private Function GetShapeTag(shp As Shape, ByVal name As String) As String
    On Error Resume Next
    GetShapeTag = shp.Tags(name)
    On Error GoTo 0
End Function

' Local copy so EntryManager doesn't bind to the (ambiguous) public ShapeExists that
' lives in CreateEntry / SnapLines / ToggleSuffixes. A module-local definition shadows them.
Private Function ShapeExists(sld As slide, ByVal shapeName As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = sld.Shapes(shapeName)
    On Error GoTo 0
    ShapeExists = Not (shp Is Nothing)
End Function

Private Function GroupDayShape(sld As slide, ByVal dy As Long, ByVal groupName As String) As Shape
    Dim suf As Integer, nm As String, shp As Shape
    For suf = Asc("A") To Asc("D")
        nm = "Entry_" & dy & "_" & Chr(suf)
        If ShapeExists(sld, nm) Then
            Set shp = sld.Shapes(nm)
            If StrComp(GroupOfEntry(shp), groupName, vbTextCompare) = 0 Then
                Set GroupDayShape = shp: Exit Function
            End If
        End If
    Next suf
End Function

Private Function GroupHasDay(sld As slide, ByVal dy As Long, ByVal groupName As String) As Boolean
    GroupHasDay = Not (GroupDayShape(sld, dy, groupName) Is Nothing)
End Function

' ---- placement / geometry --------------------------------------------------
' Lay out all of a day's entries (1..4) like the legacy resizer, but color each
' from its OWN group (not insertion order).
Public Sub RepositionDayEntries(sld As slide, ByVal dy As Long, reg As Object)
    Dim yr As Long, mo As Long, startDay As Integer, row As Integer, col As Integer
    Dim cellShape As Shape, ents As Collection, shp As Shape
    Dim suf As Integer, nm As String, total As Integer, i As Integer, c As Long
    yr = CalYearOf(sld): mo = CalMonthOf(sld)
    If yr = 0 Or mo = 0 Then Exit Sub
    startDay = Weekday(DateSerial(yr, mo, 1), vbSunday)
    row = ((dy + startDay - 2) \ 7) + 1
    col = ((dy + startDay - 2) Mod 7) + 1
    Set cellShape = sld.Shapes("CalendarGrid").Table.Cell(row, col).Shape

    Set ents = New Collection
    For suf = Asc("A") To Asc("D")
        nm = "Entry_" & dy & "_" & Chr(suf)
        If ShapeExists(sld, nm) Then ents.Add sld.Shapes(nm)
    Next suf
    total = ents.count

    For i = 1 To total
        Set shp = ents(i)
        Select Case total
            Case 1
                shp.Left = cellShape.Left: shp.Top = cellShape.Top
                shp.Width = cellShape.Width: shp.Height = cellShape.Height
                shp.TextFrame.TextRange.Font.Size = 10
            Case 2
                shp.Left = cellShape.Left: shp.Width = cellShape.Width
                shp.Height = cellShape.Height / 2
                shp.Top = cellShape.Top + (i - 1) * (cellShape.Height / 2)
                shp.TextFrame.TextRange.Font.Size = 8
            Case 3
                shp.Width = cellShape.Width / 2: shp.Height = cellShape.Height / 2
                If i <= 2 Then
                    shp.Top = cellShape.Top
                    shp.Left = cellShape.Left + (i - 1) * (cellShape.Width / 2)
                Else
                    shp.Top = cellShape.Top + cellShape.Height / 2
                    shp.Left = cellShape.Left
                    shp.Width = cellShape.Width
                End If
                shp.TextFrame.TextRange.Font.Size = 7
            Case 4
                shp.Width = cellShape.Width / 2: shp.Height = cellShape.Height / 2
                Select Case i
                    Case 1
                        shp.Top = cellShape.Top: shp.Left = cellShape.Left
                    Case 2
                        shp.Top = cellShape.Top: shp.Left = cellShape.Left + cellShape.Width / 2
                    Case 3
                        shp.Top = cellShape.Top + cellShape.Height / 2: shp.Left = cellShape.Left
                    Case 4
                        shp.Top = cellShape.Top + cellShape.Height / 2
                        shp.Left = cellShape.Left + cellShape.Width / 2
                End Select
                shp.TextFrame.TextRange.Font.Size = 7
        End Select
        c = GroupColor(reg, GroupOfEntry(shp))
        If c = 0 Then c = RGB(200, 200, 200)
        shp.Fill.ForeColor.RGB = c
    Next i
End Sub

' Add one entry for a group. reason returns "duplicate" / "day full (4 max)" on skip.
Public Function AddEntryForGroup(sld As slide, ByVal yr As Long, ByVal mo As Long, ByVal dy As Long, _
        ByVal groupName As String, reg As Object, ByVal label As String, ByRef reason As String) As Boolean
    Dim suffix As String, startDay As Integer, row As Integer, col As Integer
    Dim cellShape As Shape, newEntry As Shape
    reason = ""
    If GroupHasDay(sld, dy, groupName) Then reason = "duplicate": Exit Function
    suffix = NextAvailableSuffix(sld, CInt(dy))
    If suffix = "" Then reason = "day full (4 max)": Exit Function

    startDay = Weekday(DateSerial(yr, mo, 1), vbSunday)
    row = ((dy + startDay - 2) \ 7) + 1
    col = ((dy + startDay - 2) Mod 7) + 1
    Set cellShape = sld.Shapes("CalendarGrid").Table.Cell(row, col).Shape
    Set newEntry = sld.Shapes.AddShape(msoShapeRectangle, cellShape.Left, cellShape.Top, _
                                       cellShape.Width, cellShape.Height)
    newEntry.name = "Entry_" & dy & "_" & suffix
    newEntry.Tags.Add "EntryGroup", groupName
    newEntry.Tags.Add "EntryDay", CStr(dy)
    newEntry.Tags.Add "EntryDesc", label
    newEntry.Tags.Add "ShowDesc", "0"           ' descriptions hidden by default
    With newEntry.TextFrame.TextRange
        .Font.name = "Arial"
        .Font.Size = 10
        .Font.color.RGB = RGB(0, 0, 0)
    End With
    ApplyEntryText newEntry                      ' shows the description only if ShowDesc = 1
    newEntry.ZOrder msoSendToBack                ' sit behind the calendar grid
    RepositionDayEntries sld, dy, reg
    AddEntryForGroup = True
End Function

' Render an entry's visible text: its description if ShowDesc=1, else blank.
Public Sub ApplyEntryText(shp As Shape)
    Dim show As Boolean, desc As String
    show = (GetShapeTag(shp, "ShowDesc") = "1")
    desc = GetShapeTag(shp, "EntryDesc")
    On Error Resume Next
    shp.TextFrame.TextRange.text = IIf(show, desc, "")
    On Error GoTo 0
End Sub

' Rename surviving same-day entries to contiguous A,B,C... (two-pass, collision-free)
' then reflow - REQUIRED after any delete because the layout maps A..D contiguously.
Public Sub CompactDaySuffixes(sld As slide, ByVal dy As Long, reg As Object)
    Dim survivors As Collection, suf As Integer, nm As String, i As Integer
    Set survivors = New Collection
    For suf = Asc("A") To Asc("D")
        nm = "Entry_" & dy & "_" & Chr(suf)
        If ShapeExists(sld, nm) Then survivors.Add sld.Shapes(nm)
    Next suf
    For i = 1 To survivors.count
        survivors(i).name = "Entry_" & dy & "_TMP" & i
    Next i
    For i = 1 To survivors.count
        survivors(i).name = "Entry_" & dy & "_" & Chr(64 + i)
    Next i
    RepositionDayEntries sld, dy, reg
End Sub

Public Sub RemoveEntryShape(shp As Shape, reg As Object)
    Dim sld As slide, dy As Long
    Set sld = shp.Parent
    dy = DayOfEntry(shp)
    shp.Delete
    CompactDaySuffixes sld, dy, reg
End Sub

' ---- bulk add from the parsed buffer ---------------------------------------
Public Sub AddParsedDates(outDates As Collection, ByVal groupName As String, reg As Object, _
        ByRef added As Long, ByRef skippedDup As Long, ByRef skippedFull As Long, ByRef noSlide As Collection)
    Dim it As Variant, yr As Long, mo As Long, dy As Long, k As Long
    Dim sld As slide, reason As String, affected As Object, sk As Variant
    Dim seenNoSlide As Object, key As String
    Set noSlide = New Collection
    Set affected = CreateObject("Scripting.Dictionary")
    Set seenNoSlide = CreateObject("Scripting.Dictionary")
    seenNoSlide.CompareMode = vbTextCompare

    For k = 1 To outDates.count
        it = outDates(k)
        yr = it(0): mo = it(1): dy = it(2)
        Set sld = FindMonthSlide(yr, mo)
        If sld Is Nothing Then
            key = MonthName(mo) & " " & yr
            If Not seenNoSlide.Exists(key) Then
                seenNoSlide(key) = True
                noSlide.Add key
            End If
        Else
            reason = ""
            If AddEntryForGroup(sld, yr, mo, dy, groupName, reg, "Example text", reason) Then
                added = added + 1
                affected(sld.SlideIndex) = True
            ElseIf reason = "duplicate" Then
                skippedDup = skippedDup + 1
            ElseIf reason = "day full (4 max)" Then
                skippedFull = skippedFull + 1
            End If
        End If
    Next k

    For Each sk In affected.Keys
        DrawLegend ActivePresentation.Slides(CLng(sk)), reg
    Next sk
End Sub

' ---- scan existing entries -------------------------------------------------
' Returns a Collection of Array(slideIndex, year, month, day, group, label, shapeName),
' migrating untagged legacy entries into DEFAULT_GROUP on touch.
' Returns Array(slideIndex, year, month, day, group, description, shapeName, showDesc("0"/"1")).
Public Function ScanEntries(reg As Object) As Collection
    Dim col As Collection, sld As slide, shp As Shape
    Dim yr As Long, mo As Long, g As String, desc As String, sd As String
    Set col = New Collection
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then
            yr = CalYearOf(sld): mo = CalMonthOf(sld)
            For Each shp In sld.Shapes
                If IsEntryShape(shp) Then
                    g = GetShapeTag(shp, "EntryGroup")
                    If g = "" Then
                        g = DEFAULT_GROUP
                        shp.Tags.Add "EntryGroup", g
                    End If
                    desc = GetShapeTag(shp, "EntryDesc")
                    sd = GetShapeTag(shp, "ShowDesc")
                    If desc = "" And sd = "" Then
                        ' legacy entry: adopt its current text as the description, keep it shown
                        On Error Resume Next
                        desc = shp.TextFrame.TextRange.text
                        On Error GoTo 0
                        desc = Replace(Replace(desc, vbCr, " "), vbLf, " ")
                        shp.Tags.Add "EntryDesc", desc
                        shp.Tags.Add "ShowDesc", "1"
                        sd = "1"
                    End If
                    If sd = "" Then sd = "0"
                    col.Add Array(sld.SlideIndex, yr, mo, DayOfEntry(shp), g, desc, shp.name, sd)
                End If
            Next shp
        End If
    Next sld
    Set ScanEntries = col
End Function

' ---- per-day toggle (Grid tab) ---------------------------------------------
Public Sub ToggleDay(sld As slide, ByVal yr As Long, ByVal mo As Long, ByVal dy As Long, _
                     ByVal groupName As String, reg As Object, ByRef msg As String)
    Dim shp As Shape, reason As String
    Set shp = GroupDayShape(sld, dy, groupName)
    If Not shp Is Nothing Then
        RemoveEntryShape shp, reg
        msg = "removed"
    Else
        reason = ""
        If AddEntryForGroup(sld, yr, mo, dy, groupName, reg, "Example text", reason) Then
            msg = "added"
        Else
            msg = reason
        End If
    End If
    DrawLegend sld, reg
End Sub

' Is there ANY entry on this day (any group)? (drives the grid's neutral shading)
Public Function DayHasAnyEntry(sld As slide, ByVal dy As Long) As Boolean
    DayHasAnyEntry = (CountDayEntries(sld, CInt(dy)) > 0)
End Function

Public Function DayHasGroup(sld As slide, ByVal dy As Long, ByVal groupName As String) As Boolean
    DayHasGroup = GroupHasDay(sld, dy, groupName)
End Function

' ---- single-entry edits (All Entries tab) ----------------------------------
' Set an entry's description (stored on the shape; shown only if ShowDesc=1).
Public Sub SetEntryLabelText(shp As Shape, ByVal txt As String)
    On Error Resume Next
    shp.Tags.Delete "EntryDesc"
    On Error GoTo 0
    shp.Tags.Add "EntryDesc", txt
    ApplyEntryText shp
End Sub

' Show or hide an entry's description text.
Public Sub SetEntryShowDesc(shp As Shape, ByVal show As Boolean)
    On Error Resume Next
    shp.Tags.Delete "ShowDesc"
    On Error GoTo 0
    shp.Tags.Add "ShowDesc", IIf(show, "1", "0")
    ApplyEntryText shp
End Sub

Public Sub SetEntryGroupOf(shp As Shape, ByVal newGroup As String, reg As Object)
    On Error Resume Next
    shp.Tags.Delete "EntryGroup"
    On Error GoTo 0
    shp.Tags.Add "EntryGroup", newGroup
    shp.Fill.ForeColor.RGB = GroupColor(reg, newGroup)
End Sub

' ---- group operations with shape side-effects ------------------------------
Public Sub RetagGroupShapes(ByVal oldName As String, ByVal newName As String)
    Dim sld As slide, shp As Shape
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then
            For Each shp In sld.Shapes
                If IsEntryShape(shp) Then
                    If StrComp(GroupOfEntry(shp), oldName, vbTextCompare) = 0 Then
                        On Error Resume Next
                        shp.Tags.Delete "EntryGroup"
                        On Error GoTo 0
                        shp.Tags.Add "EntryGroup", newName
                    End If
                End If
            Next shp
        End If
    Next sld
End Sub

Public Sub RecolorGroupShapes(reg As Object, ByVal name As String)
    Dim sld As slide, shp As Shape, c As Long
    c = GroupColor(reg, name)
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then
            For Each shp In sld.Shapes
                If IsEntryShape(shp) Then
                    If StrComp(GroupOfEntry(shp), name, vbTextCompare) = 0 Then
                        shp.Fill.ForeColor.RGB = c
                    End If
                End If
            Next shp
        End If
    Next sld
End Sub

' Delete a group everywhere: either delete its entries (compacting each affected day)
' or reassign them to DEFAULT_GROUP. Removes the registry key, saves, redraws legends.
Public Sub DeleteGroupEverywhere(reg As Object, ByVal name As String, ByVal alsoDeleteEntries As Boolean)
    Dim sld As slide, shp As Shape, i As Long, daysHit As Object, dk As Variant, dnum As Long
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then
            Set daysHit = CreateObject("Scripting.Dictionary")
            For i = sld.Shapes.count To 1 Step -1
                Set shp = sld.Shapes(i)
                If IsEntryShape(shp) Then
                    If StrComp(GroupOfEntry(shp), name, vbTextCompare) = 0 Then
                        If alsoDeleteEntries Then
                            dnum = DayOfEntry(shp)
                            shp.Delete
                            daysHit(dnum) = True
                        Else
                            On Error Resume Next
                            shp.Tags.Delete "EntryGroup"
                            On Error GoTo 0
                            shp.Tags.Add "EntryGroup", DEFAULT_GROUP
                            shp.Fill.ForeColor.RGB = GroupColor(reg, DEFAULT_GROUP)
                        End If
                    End If
                End If
            Next i
            If alsoDeleteEntries Then
                For Each dk In daysHit.Keys
                    CompactDaySuffixes sld, CLng(dk), reg
                Next dk
            End If
        End If
    Next sld
    RemoveGroupKey reg, name
    SaveRegistry reg
    DrawAllLegends reg
End Sub

' ---- legend ----------------------------------------------------------------
Public Sub DrawAllLegends(reg As Object)
    Dim sld As slide
    For Each sld In ActivePresentation.Slides
        If IsCalendarSlide(sld) Then DrawLegend sld, reg
    Next sld
End Sub

' One legend per slide, named "EntryLegend" and tagged EntryLegend=TRUE so it is
' delete-then-rebuilt (idempotent). Shows only the groups present on THIS slide.
Public Sub DrawLegend(sld As slide, reg As Object)
    Dim shp As Shape, i As Long, present As Object, names As Variant
    Dim fontSz As Single, sw As Single, gap As Single, x0 As Single, yy As Single
    Dim tbl As Shape, made As Collection, n As Long, gname As String
    Dim sq As Shape, tb As Shape, arr() As String, j As Long, grp As Shape, g As String

    ' remove any existing legend on this slide
    For i = sld.Shapes.count To 1 Step -1
        Set shp = sld.Shapes(i)
        If GetShapeTag(shp, "EntryLegend") <> "" Then shp.Delete
    Next i
    If Not LegendOn(reg) Then Exit Sub

    Set present = CreateObject("Scripting.Dictionary")
    present.CompareMode = vbTextCompare
    For Each shp In sld.Shapes
        If IsEntryShape(shp) Then
            g = GroupOfEntry(shp)
            If Not present.Exists(g) Then present.Add g, True
        End If
    Next shp
    If present.count = 0 Then Exit Sub

    On Error Resume Next
    Set tbl = sld.Shapes("CalendarGrid")
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub

    names = GroupNamesInOrder(reg)
    fontSz = LegendFontSize(reg)
    sw = fontSz * 1.3
    gap = 4
    x0 = tbl.Left
    yy = tbl.Top + tbl.Height + 6

    Set made = New Collection
    For n = LBound(names) To UBound(names)
        gname = names(n)
        If present.Exists(gname) Then
            Set sq = sld.Shapes.AddShape(msoShapeRectangle, x0, yy, sw, sw)
            sq.Fill.ForeColor.RGB = GroupColor(reg, gname)
            sq.Line.ForeColor.RGB = RGB(0, 0, 0)
            sq.Line.Weight = 0.75
            sq.Tags.Add "EntryLegend", "TRUE"
            made.Add sq
            Set tb = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + sw + gap, yy - 2, 120, sw + 4)
            tb.TextFrame.WordWrap = msoFalse
            With tb.TextFrame.TextRange
                .text = gname
                .Font.name = "Arial"
                .Font.Size = fontSz
                .Font.color.RGB = RGB(0, 0, 0)
            End With
            tb.TextFrame.AutoSize = ppAutoSizeShapeToFitText
            tb.Tags.Add "EntryLegend", "TRUE"
            made.Add tb
            yy = yy + sw + gap
        End If
    Next n

    If made.count >= 2 Then
        ReDim arr(0 To made.count - 1)
        For j = 1 To made.count
            arr(j - 1) = made(j).name
        Next j
        Set grp = sld.Shapes.Range(arr).Group
        grp.name = "EntryLegend"
        grp.Tags.Add "EntryLegend", "TRUE"
    ElseIf made.count = 1 Then
        made(1).name = "EntryLegend"
    End If
End Sub

' ============================================================================
' Buffer parser - sticky year/month grammar.
' ============================================================================
' Fills outDates (each item Array(year, month, day, lineNo)) and outErrors (strings).
Public Sub ParseBuffer(ByVal text As String, ByVal defaultYear As Long, _
                       ByRef outDates As Collection, ByRef outErrors As Collection)
    Dim lines() As String, i As Long, activeYear As Long, activeMonth As Long
    Set outDates = New Collection
    Set outErrors = New Collection
    activeYear = defaultYear
    activeMonth = 0
    text = Replace(text, vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        ParseLine Trim(lines(i)), i + 1, activeYear, activeMonth, outDates, outErrors
    Next i
End Sub

Private Sub ParseLine(ByVal ln As String, ByVal lineNo As Long, ByRef activeYear As Long, _
                      ByRef activeMonth As Long, ByRef outDates As Collection, ByRef outErrors As Collection)
    Dim colonPos As Long, monTok As String, mo As Long, rest As String
    Dim sp As Long, firstTok As String, wholeMo As Long
    If ln = "" Then Exit Sub
    If Left(ln, 1) = ";" Then Exit Sub

    ' (a) exactly 4 digits => year header
    If Len(ln) = 4 And IsAllDigits(ln) Then
        activeYear = CLng(ln)
        Exit Sub
    End If

    ' (b) "<month>: dayspec"
    colonPos = InStr(ln, ":")
    If colonPos > 0 Then
        monTok = Trim(Left(ln, colonPos - 1))
        mo = MonthFromToken(monTok)
        If mo = 0 Then
            outErrors.Add "Line " & lineNo & ": '" & monTok & "' is not a month."
            Exit Sub
        End If
        activeMonth = mo
        rest = Trim(Mid(ln, colonPos + 1))
        If rest <> "" Then EmitDaySpec rest, activeYear, activeMonth, lineNo, outDates, outErrors
        Exit Sub
    End If

    ' (c) a lone month NAME (never a bare number) sets the active month
    If Not IsAllDigits(ln) And InStr(ln, "/") = 0 Then
        wholeMo = MonthFromToken(ln)
        If wholeMo > 0 Then activeMonth = wholeMo: Exit Sub
    End If

    ' (d) inline "MonthName dayspec"
    sp = InStr(ln, " ")
    If sp > 0 Then
        firstTok = Left(ln, sp - 1)
        mo = MonthFromToken(firstTok)
        If mo > 0 Then
            activeMonth = mo
            rest = Trim(Mid(ln, sp + 1))
            EmitDaySpec rest, activeYear, activeMonth, lineNo, outDates, outErrors
            Exit Sub
        End If
    End If

    ' (e) full dates with "/"  (M/D or M/D/YYYY, comma-separated allowed)
    If InStr(ln, "/") > 0 Then
        EmitFullDates ln, activeYear, lineNo, outDates, outErrors
        Exit Sub
    End If

    ' (f) bare dayspec under the active month
    If activeMonth = 0 Then
        outErrors.Add "Line " & lineNo & ": '" & ln & "' has no month yet (add a year and a 'Mar:' header first)."
        Exit Sub
    End If
    EmitDaySpec ln, activeYear, activeMonth, lineNo, outDates, outErrors
End Sub

Private Sub EmitDaySpec(ByVal spec As String, ByVal yr As Long, ByVal mo As Long, ByVal lineNo As Long, _
                        ByRef outDates As Collection, ByRef outErrors As Collection)
    Dim parts() As String, i As Long, p As String, ab() As String
    Dim a As Long, b As Long, d As Long, t As Long, dmax As Long
    dmax = DaysInMonth(yr, mo)
    parts = Split(spec, ",")
    For i = 0 To UBound(parts)
        p = Trim(parts(i))
        If p <> "" Then
            If InStr(p, "-") > 0 Then
                ab = Split(p, "-")
                If UBound(ab) = 1 And IsAllDigits(Trim(ab(0))) And IsAllDigits(Trim(ab(1))) Then
                    a = CLng(Trim(ab(0))): b = CLng(Trim(ab(1)))
                    If a > b Then t = a: a = b: b = t
                    For d = a To b
                        AddDay yr, mo, d, dmax, lineNo, outDates, outErrors
                    Next d
                Else
                    outErrors.Add "Line " & lineNo & ": bad range '" & p & "'."
                End If
            ElseIf IsAllDigits(p) Then
                AddDay yr, mo, CLng(p), dmax, lineNo, outDates, outErrors
            Else
                outErrors.Add "Line " & lineNo & ": '" & p & "' is not a day."
            End If
        End If
    Next i
End Sub

Private Sub EmitFullDates(ByVal ln As String, ByVal yr As Long, ByVal lineNo As Long, _
                          ByRef outDates As Collection, ByRef outErrors As Collection)
    Dim toks() As String, i As Long, tk As String, dp() As String, yv As Long
    toks = Split(ln, ",")
    For i = 0 To UBound(toks)
        tk = Trim(toks(i))
        If tk <> "" Then
            dp = Split(tk, "/")
            If UBound(dp) = 1 Then
                If IsAllDigits(Trim(dp(0))) And IsAllDigits(Trim(dp(1))) Then
                    AddDayFull yr, CLng(Trim(dp(0))), CLng(Trim(dp(1))), lineNo, outDates, outErrors
                Else
                    outErrors.Add "Line " & lineNo & ": bad date '" & tk & "'."
                End If
            ElseIf UBound(dp) = 2 Then
                If IsAllDigits(Trim(dp(0))) And IsAllDigits(Trim(dp(1))) And IsAllDigits(Trim(dp(2))) Then
                    yv = CLng(Trim(dp(2)))
                    If Len(Trim(dp(2))) <= 2 Then yv = 2000 + yv
                    AddDayFull yv, CLng(Trim(dp(0))), CLng(Trim(dp(1))), lineNo, outDates, outErrors
                Else
                    outErrors.Add "Line " & lineNo & ": bad date '" & tk & "'."
                End If
            Else
                outErrors.Add "Line " & lineNo & ": bad date '" & tk & "'."
            End If
        End If
    Next i
End Sub

Private Sub AddDayFull(ByVal yr As Long, ByVal mo As Long, ByVal dy As Long, ByVal lineNo As Long, _
                       ByRef outDates As Collection, ByRef outErrors As Collection)
    If mo < 1 Or mo > 12 Then
        outErrors.Add "Line " & lineNo & ": month " & mo & " is invalid."
        Exit Sub
    End If
    AddDay yr, mo, dy, DaysInMonth(yr, mo), lineNo, outDates, outErrors
End Sub

Private Sub AddDay(ByVal yr As Long, ByVal mo As Long, ByVal dy As Long, ByVal dmax As Long, _
                   ByVal lineNo As Long, ByRef outDates As Collection, ByRef outErrors As Collection)
    If dy < 1 Or dy > dmax Then
        outErrors.Add "Line " & lineNo & ": " & MonthName(mo, True) & " " & dy & " is not a valid day."
        Exit Sub
    End If
    outDates.Add Array(yr, mo, dy, lineNo)
End Sub

' ---- token helpers ---------------------------------------------------------
Public Function MonthFromToken(ByVal tok As String) As Long
    Dim n As Long
    tok = Trim(tok)
    If tok = "" Then Exit Function
    If IsAllDigits(tok) Then
        n = CLng(tok)
        If n >= 1 And n <= 12 Then MonthFromToken = n
        Exit Function
    End If
    If Not IsMonthWord(tok) Then Exit Function
    Select Case LCase(Left(tok, 3))
        Case "jan": MonthFromToken = 1
        Case "feb": MonthFromToken = 2
        Case "mar": MonthFromToken = 3
        Case "apr": MonthFromToken = 4
        Case "may": MonthFromToken = 5
        Case "jun": MonthFromToken = 6
        Case "jul": MonthFromToken = 7
        Case "aug": MonthFromToken = 8
        Case "sep": MonthFromToken = 9
        Case "oct": MonthFromToken = 10
        Case "nov": MonthFromToken = 11
        Case "dec": MonthFromToken = 12
    End Select
End Function

Private Function IsMonthWord(ByVal tok As String) As Boolean
    Select Case LCase(Trim(tok))
        Case "jan", "january", "feb", "february", "mar", "march", "apr", "april", _
             "may", "jun", "june", "jul", "july", "aug", "august", "sep", "sept", _
             "september", "oct", "october", "nov", "november", "dec", "december"
            IsMonthWord = True
    End Select
End Function

Public Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, c As Integer
    s = Trim(s)
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        c = Asc(Mid(s, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    IsAllDigits = True
End Function
