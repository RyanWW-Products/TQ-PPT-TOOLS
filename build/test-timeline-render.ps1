<#
.SYNOPSIS
    Runs timeline import and render regressions in a new hidden presentation
    using the local master's components. Requires desktop-session VBOM access.
.DESCRIPTION
    Opens a read-only master copy with macros disabled to export components and
    references, then imports them into a newly created hidden presentation.
    Imports Modules/*.bas and preserves embedded forms. Appends test-only VBA to
    the fixture TimelineCreator module. Within that copy, ActivePresentation is
    redirected to the fixture deck so hidden tests cannot alter a user's deck.
    Produces a report, PNG previews, and a test-only PPTM under build/Output.
    No Trust Center, add-in installation, source master, or exported VBA changes.
#>
[CmdletBinding()]
param(
    [string]$SourcePptm,
    [string]$CorrectedWorkbook = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'timeline_import_2099_corrected.xlsx'),
    [string]$OriginalWorkbook = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'timeline_import_2099.xlsx')
)
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
if (-not $SourcePptm) {
    $candidate = Get-ChildItem -LiteralPath $repoRoot -Filter 'TrialQuest Addin Master v*.pptm' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $candidate) { throw 'No source PPTM found. Supply -SourcePptm.' }
    $SourcePptm = $candidate.FullName
}
$SourcePptm = (Resolve-Path -LiteralPath $SourcePptm).Path
$CorrectedWorkbook = (Resolve-Path -LiteralPath $CorrectedWorkbook).Path
$OriginalWorkbook = (Resolve-Path -LiteralPath $OriginalWorkbook).Path
$runName = 'timeline-regression-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$outputDirectory = Join-Path (Join-Path $PSScriptRoot 'Output') $runName
[void](New-Item -ItemType Directory -Path $outputDirectory -Force)
$testDeckPath = Join-Path $outputDirectory 'TimelineRegression.pptm'
$reportPath = Join-Path $outputDirectory 'report.txt'
$sourceCopyPath = Join-Path $outputDirectory 'SourceMaster.pptm'
$componentDirectory = Join-Path $outputDirectory 'original-components'
[void](New-Item -ItemType Directory -Path $componentDirectory)
Copy-Item -LiteralPath $SourcePptm -Destination $sourceCopyPath

# This code is appended only to the temporary TimelineCreator code module.
$testCode = @'

Private Function TQTestTargetPresentation() As Presentation
    Set TQTestTargetPresentation = Application.Presentations("TimelineRegression.pptm")
End Function

Private Sub TQTestCheck(ByVal fh As Integer, ByVal condition As Boolean, ByVal message As String, _
                        ByRef checks As Long, ByRef failures As Long)
    checks = checks + 1
    If condition Then
        Print #fh, "PASS | " & message
    Else
        failures = failures + 1
        Print #fh, "FAIL | " & message
    End If
End Sub

Private Function TQTestNewSlide(ByVal name As String) As slide
    Dim deck As Presentation, sld As slide
    Set deck = TQTestTargetPresentation()
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    sld.Name = name
    Set TQTestNewSlide = sld
End Function

Private Sub TQTestEntryBounds(ByVal sld As slide, ByVal caseName As String, ByVal fh As Integer, _
                             ByRef checks As Long, ByRef failures As Long)
    Dim entries As New Collection, entry As Shape, topY As Single, leftX As Single
    Dim bottomY As Single, rightX As Single, sw As Single, sh As Single, allInside As Boolean
    sw = sld.Parent.PageSetup.SlideWidth: sh = sld.Parent.PageSetup.SlideHeight
    allInside = True
    CollectEntryGroups sld.Shapes, entries
    For Each entry In entries
        BoxRectOf entry, topY, leftX, bottomY, rightX
        If leftX < -0.25 Or topY < -0.25 Or rightX > sw + 0.25 Or bottomY > sh + 0.25 Then
            allInside = False
            Print #fh, "BOUNDS | " & caseName & " | slide=" & sld.SlideIndex & " | " & entry.Name & _
                       " | left=" & leftX & " | top=" & topY & " | right=" & rightX & " | bottom=" & bottomY
        End If
    Next entry
    TQTestCheck fh, allInside, caseName & " keeps every entry inside slide " & sld.SlideIndex, checks, failures
End Sub

Private Sub TQTestBarRectangles(ByVal bar As Shape, ByVal caseName As String, ByVal fh As Integer, _
                               ByRef checks As Long, ByRef failures As Long, ByRef cells As Long)
    Dim child As Shape, isRectangle As Boolean
    If bar Is Nothing Then Exit Sub
    If bar.Type = msoGroup Then
        For Each child In bar.GroupItems
            TQTestBarRectangles child, caseName, fh, checks, failures, cells
        Next child
    ElseIf bar.HasTextFrame Then
        If Len(Trim$(bar.TextFrame.TextRange.Text)) > 0 Then
            cells = cells + 1
            If bar.Type = msoAutoShape Then isRectangle = (bar.AutoShapeType = msoShapeRectangle)
            TQTestCheck fh, isRectangle, caseName & " adjacent hour cell stays rectangular (no false tear)", checks, failures
        End If
    End If
End Sub

Private Sub TQTestLeaderGeometry(ByVal sld As slide, ByVal caseName As String, ByVal fh As Integer, _
                                 ByRef checks As Long, ByRef failures As Long, Optional ByVal exactDates As Boolean = False)
    Dim entries As New Collection, entry As Shape, ln As Shape, db As Shape, bar As Shape
    Dim lefts As Object, widths As Object, d As Date, key As String, t As String
    Dim expectedX As Single, anchorsOK As Boolean, contactsOK As Boolean, verticalOK As Boolean, count As Long
    Set bar = FindDateBar(sld)
    If bar Is Nothing Then Exit Sub
    t = BarUnit(sld, bar)
    Set lefts = CreateObject("Scripting.Dictionary"): Set widths = CreateObject("Scripting.Dictionary")
    CollectBandCells bar, lefts, widths
    anchorsOK = True: contactsOK = True: verticalOK = True
    CollectEntryGroups sld.Shapes, entries
    For Each entry In entries
        Set ln = LeadingLineOf(entry)
        If Not ln Is Nothing Then
            count = count + 1
            If Abs(ln.Width) > 0.01 Then
                verticalOK = False
                Print #fh, "VERTICAL | " & caseName & " | " & entry.Name & " | width=" & ln.Width
            End If
            If exactDates Then
            If StoredEntryDate(entry, d) Then
                If ShapeTagVal(entry, "TLTimeKnown") = "0" And t <> "Hours" Then d = Int(CDbl(d)) + 0.5
                key = DateKeyOf(UnitStartOf(d, t))
                If lefts.Exists(key) Then
                    expectedX = CSng(lefts(key)) + COL_PAD + UnitFracOf(d, t) * (CSng(widths(key)) - 2 * COL_PAD)
                    If Abs(ln.Left - expectedX) > 0.2 Then
                        anchorsOK = False
                        Print #fh, "ANCHOR | " & caseName & " | " & entry.Name & " | expected=" & expectedX & " | actual=" & ln.Left
                    End If
                Else
                    anchorsOK = False
                End If
            Else
                anchorsOK = False
            End If
            End If
            Set db = DateBoxOf(entry)
            If db Is Nothing Then
                contactsOK = False
            Else
                If ln.Left < db.Left - 0.2 Or ln.Left > db.Left + db.Width + 0.2 Or _
                   Abs(ln.Top + ln.Height - db.Top - db.Height) > 0.2 Then
                    contactsOK = False
                    Print #fh, "CONTACT | " & caseName & " | " & entry.Name & " | lineX=" & ln.Left & " | boxLeft=" & db.Left & " | boxRight=" & db.Left + db.Width
                End If
            End If
        End If
    Next entry
    TQTestCheck fh, count > 0 And verticalOK, caseName & " every leader is exactly vertical", checks, failures
    If exactDates Then TQTestCheck fh, count > 0 And anchorsOK, caseName & " leaders meet exact time positions on the bar", checks, failures
    TQTestCheck fh, count > 0 And contactsOK, caseName & " leader bottoms meet their date boxes", checks, failures
End Sub

Private Sub TQTestImportRender(ByVal path As String, ByVal caseName As String, ByVal expectedCount As Long, _
                               ByVal expectedUntimed As Long, ByVal outputDir As String, ByVal fh As Integer, _
                               ByRef checks As Long, ByRef failures As Long)
    Dim ev() As TLEvent, restored() As TLEvent, n As Long, restoredN As Long, errors As String
    Dim sld As slide, bar As Shape, entry As Shape, ln As Shape, entries As New Collection
    Dim i As Long, untimed As Long, noLeaders As Long, leaderCount As Long, summary As String
    Dim cols() As Date, colCount As Long, t As String, color As String
    Dim gaps As Boolean, multi As Boolean, wipe As Boolean, weighted As Boolean, same As Boolean
    n = ReadEvents(path, ev, errors)
    Print #fh, "CASE | " & caseName & " | entries=" & n & " | importErrors=" & Replace(errors, vbCrLf, " / ")
    TQTestCheck fh, n = expectedCount, caseName & " imports all expected rows", checks, failures
    TQTestCheck fh, Len(errors) = 0, caseName & " has no import errors", checks, failures
    If n = 0 Then Exit Sub
    TQTestCheck fh, DetectType(ev, n) = "Hours", caseName & " detects Hours", checks, failures
    PrepareEventUnits ev, n, "Hours"
    colCount = ComputeColumns(ev, n, "Hours", False, cols)
    TQTestCheck fh, colCount = 2, caseName & " has exactly two hour cells", checks, failures
    If colCount > 0 Then
        TQTestCheck fh, Abs(CDbl(cols(1)) - CDbl(DateSerial(2099, 1, 1) + TimeSerial(22, 0, 0))) * 86400 < 0.001, _
                    caseName & " begins at 10:00 PM", checks, failures
        TQTestCheck fh, Abs(CDbl(DateAdd("h", 1, cols(colCount))) - CDbl(DateSerial(2099, 1, 2))) * 86400 < 0.001, _
                    caseName & " ends at midnight edge", checks, failures
    End If
    For i = 1 To n
        If Not ev(i).HasTime Then
            untimed = untimed + 1
            TQTestCheck fh, ev(i).RawDate = DateSerial(2099, 1, 1), _
                        caseName & " untimed row " & ev(i).OrigIndex & " retains date only", checks, failures
            TQTestCheck fh, InStr(ev(i).DateLabel, ":") = 0, _
                        caseName & " untimed row " & ev(i).OrigIndex & " shows no invented time", checks, failures
            If i > 1 Then
                TQTestCheck fh, ev(i).OrigIndex > ev(i - 1).OrigIndex, _
                            caseName & " untimed row stays after preceding spreadsheet row", checks, failures
            End If
            If i < n Then
                TQTestCheck fh, ev(i).OrigIndex < ev(i + 1).OrigIndex, _
                            caseName & " untimed row stays before following spreadsheet row", checks, failures
            End If
        End If
    Next i
    TQTestCheck fh, untimed = expectedUntimed, caseName & " preserves explicit-time count", checks, failures
    Set sld = TQTestNewSlide(caseName)
    TQTestCheck fh, RenderTimeline(sld, ev, n, "Hours", "Gray", False, False, False, False, summary), _
                caseName & " renders successfully", checks, failures
    Print #fh, "RENDER | " & caseName & " | " & Replace(summary, vbCrLf, " | ")
    CollectEntryGroups sld.Shapes, entries
    TQTestCheck fh, entries.Count = n, caseName & " draws every entry", checks, failures
    For Each entry In entries
        Set ln = LeadingLineOf(entry)
        If ln Is Nothing Then
            noLeaders = noLeaders + 1
            TQTestCheck fh, ShapeTagVal(entry, "TLNoLeader") = "1", caseName & " leaderless entry is marked", checks, failures
        Else
            leaderCount = leaderCount + 1
        End If
    Next entry
    TQTestCheck fh, noLeaders = expectedUntimed, caseName & " only untimed rows lack leaders", checks, failures
    TQTestCheck fh, leaderCount = n - expectedUntimed, caseName & " all timed rows have leaders", checks, failures
    Set bar = FindDateBar(sld)
    Dim labels As New Collection, label As Variant
    CollectBarLabels bar, labels
    TQTestCheck fh, labels.Count = 2, caseName & " draws only two bar shapes", checks, failures
    For Each label In labels
        TQTestCheck fh, InStr(CStr(label), "12:00") = 0, caseName & " has no midnight bar shape", checks, failures
    Next label
    Dim barCells As Long
    TQTestBarRectangles bar, caseName, fh, checks, failures, barCells
    TQTestCheck fh, barCells = 2, caseName & " has two rectangular hour cells", checks, failures
    TQTestEntryBounds sld, caseName, fh, checks, failures
    TQTestLeaderGeometry sld, caseName & " initial", fh, checks, failures
    sld.Export outputDir & "\" & caseName & ".png", "PNG", 1600, 900
    For Each entry In entries
        Set ln = LeadingLineOf(entry)
        If Not ln Is Nothing Then entry.Left = entry.Left + 7
    Next entry
    ReflowTimeline sld, bar, "Hours"
    TQTestLeaderGeometry sld, caseName & " reflow", fh, checks, failures, True
    TQTestCheck fh, LoadTimelineState(sld, restored, restoredN, t, color, gaps, multi, wipe, weighted), _
                caseName & " reloads saved state", checks, failures
    same = (restoredN = n)
    If same Then
        For i = 1 To n
            If restored(i).HasTime <> ev(i).HasTime Then same = False
            If Abs(CDbl(restored(i).RawDate) - CDbl(ev(i).RawDate)) > 0.0000001 Then same = False
            If Abs(CDbl(restored(i).PlacementDate) - CDbl(ev(i).PlacementDate)) > 0.0000001 Then same = False
        Next i
    End If
    TQTestCheck fh, same, caseName & " roundtrip preserves true dates and layout anchors", checks, failures
    Dim descriptionsSame As Boolean, batesSame As Boolean, labelsSame As Boolean
    descriptionsSame = (restoredN = n): batesSame = descriptionsSame: labelsSame = descriptionsSame
    If restoredN = n Then
        For i = 1 To n
            If StrComp(restored(i).Desc, ev(i).Desc, vbBinaryCompare) <> 0 Then descriptionsSame = False
            If StrComp(restored(i).Bates, ev(i).Bates, vbBinaryCompare) <> 0 Then batesSame = False
            If StrComp(restored(i).DateLabel, ev(i).DateLabel, vbBinaryCompare) <> 0 Then labelsSame = False
        Next i
    End If
    TQTestCheck fh, descriptionsSame, caseName & " roundtrip preserves exact description text and case", checks, failures
    TQTestCheck fh, batesSame, caseName & " roundtrip preserves exact Bates text and case", checks, failures
    TQTestCheck fh, labelsSame, caseName & " roundtrip preserves exact date labels", checks, failures
    TQTestCheck fh, StrComp(t, "Hours", vbBinaryCompare) = 0, caseName & " reloads canonical Hours unit", checks, failures
    sld.Export outputDir & "\" & caseName & "-reflow.png", "PNG", 1600, 900
End Sub

Private Sub TQTestPagedRender(ByVal path As String, ByVal caseName As String, ByVal expectedCount As Long, _
                              ByVal expectedUntimed As Long, ByVal outputDir As String, ByVal fh As Integer, _
                              ByRef checks As Long, ByRef failures As Long)
    Dim ev() As TLEvent, n As Long, errors As String, summary As String, firstSlide As slide, sld As slide
    Dim entries As Collection, entry As Shape, ln As Shape, bar As Shape, owner As String
    Dim pages As Long, totalEntries As Long, totalCells As Long, missingLeaders As Long, labels As Collection, label As Variant
    n = ReadEvents(path, ev, errors)
    TQTestCheck fh, n = expectedCount And Len(errors) = 0, caseName & " imports expected rows", checks, failures
    If n = 0 Then Exit Sub
    Set firstSlide = TQTestNewSlide(caseName)
    owner = CStr(firstSlide.SlideID)
    ' One hour per page exercises actual continuation creation even on a wide slide.
    TQTestCheck fh, RenderTimeline(firstSlide, ev, n, "Hours", "Gray", False, True, False, False, summary, 0, 1), _
                caseName & " renders with multiple slides allowed", checks, failures
    Print #fh, "RENDER | " & caseName & " | " & Replace(summary, vbCrLf, " | ")
    For Each sld In TQTestTargetPresentation().Slides
        If sld.SlideID = firstSlide.SlideID Or sld.Tags(SLIDE_TAG) = owner Then
            pages = pages + 1
            Set entries = New Collection
            CollectEntryGroups sld.Shapes, entries
            totalEntries = totalEntries + entries.Count
            For Each entry In entries
                Set ln = LeadingLineOf(entry)
                If ln Is Nothing Then missingLeaders = missingLeaders + 1
            Next entry
            Set bar = FindDateBar(sld)
            TQTestCheck fh, Not bar Is Nothing, caseName & " page has a datebar", checks, failures
            TQTestCheck fh, CountDateBars(sld) = 1, caseName & " page has exactly one datebar group", checks, failures
            Set labels = New Collection
            CollectBarLabels bar, labels
            TQTestCheck fh, labels.Count = 1, caseName & " page has one hour cell", checks, failures
            For Each label In labels
                TQTestCheck fh, InStr(CStr(label), "12:00") = 0, caseName & " continuation has no midnight cell", checks, failures
            Next label
            TQTestBarRectangles bar, caseName, fh, checks, failures, totalCells
            TQTestEntryBounds sld, caseName, fh, checks, failures
            TQTestLeaderGeometry sld, caseName & " page " & pages, fh, checks, failures
            sld.Export outputDir & "\" & caseName & "-" & pages & ".png", "PNG", 1600, 900
        End If
    Next sld
    TQTestCheck fh, pages = 2, caseName & " creates two owned pages", checks, failures
    TQTestCheck fh, totalCells = 2, caseName & " keeps exactly two hour cells across pages", checks, failures
    TQTestCheck fh, totalEntries = expectedCount, caseName & " keeps every entry across pages", checks, failures
    TQTestCheck fh, missingLeaders = expectedUntimed, caseName & " preserves untimed leader suppression across pages", checks, failures
End Sub

Private Sub TQTestDateOnly(ByVal outputDir As String, ByVal fh As Integer, ByRef checks As Long, ByRef failures As Long)
    Dim ev(1 To 3) As TLEvent, i As Long, sld As slide, summary As String
    Dim entries As New Collection, entry As Shape, ln As Shape, bar As Shape
    Dim lefts As Object, widths As Object, d As Date, key As String, targetX As Single
    For i = 1 To 3
        ev(i).RawDate = DateSerial(2099, 1, i)
        ev(i).DateLabel = Format$(ev(i).RawDate, "mmm d")
        ev(i).Desc = "Date-only event " & i
        ev(i).Prec = 3
        ev(i).OrigIndex = i
    Next i
    PrepareEventPositions ev, 3
    Set sld = TQTestNewSlide("date-only")
    TQTestCheck fh, RenderTimeline(sld, ev, 3, "Days", "Gray", False, False, False, False, summary), _
                "date-only renders successfully", checks, failures
    Set bar = FindDateBar(sld)
    Set lefts = CreateObject("Scripting.Dictionary")
    Set widths = CreateObject("Scripting.Dictionary")
    CollectBandCells bar, lefts, widths
    CollectEntryGroups sld.Shapes, entries
    TQTestCheck fh, entries.Count = 3, "date-only renders three entries", checks, failures
    For Each entry In entries
        Set ln = LeadingLineOf(entry)
        TQTestCheck fh, Not ln Is Nothing, "date-only entry has a leader", checks, failures
        If Not ln Is Nothing Then
            d = CDate(CDbl(ShapeTagVal(entry, "TLFullDate")))
            key = DateKeyOf(UnitStartOf(d, "Days"))
            targetX = CSng(lefts(key)) + CSng(widths(key)) / 2
            TQTestCheck fh, Abs(ln.Left - targetX) < 0.2, "date-only leader lands at day center", checks, failures
        End If
    Next entry
    TQTestLeaderGeometry sld, "date-only initial", fh, checks, failures, True
    ReflowTimeline sld, bar, "Days"
    For Each entry In entries
        Set ln = LeadingLineOf(entry)
        If Not ln Is Nothing Then
            d = CDate(CDbl(ShapeTagVal(entry, "TLFullDate")))
            key = DateKeyOf(UnitStartOf(d, "Days"))
            targetX = CSng(lefts(key)) + CSng(widths(key)) / 2
            TQTestCheck fh, Abs(ln.Left - targetX) < 0.2, "date-only reflow keeps leader centered", checks, failures
        End If
    Next entry
    TQTestLeaderGeometry sld, "date-only reflow", fh, checks, failures, True
    sld.Export outputDir & "\date-only.png", "PNG", 1600, 900
End Sub

Private Sub TQTestLeaderEdits(ByVal outputDir As String, ByVal fh As Integer, ByRef checks As Long, ByRef failures As Long)
    Dim sld As slide, ln As Shape, aBox As Shape, bBox As Shape, aLeft As Single, bLeft As Single, untimedLeft As Single
    Set sld = TQTestNewSlide("leader-edits")
    Dim cols(1 To 1) As Date, widths(1 To 1) As Single, lefts(1 To 1) As Single
    Dim a As Shape, b As Shape, untimed As Shape, h As Single, d As Date, bar As Shape, targetA As Single, targetB As Single
    cols(1) = DateSerial(2099, 1, 1) + TimeSerial(22, 0, 0): widths(1) = 960
    DrawBar sld, cols, 1, widths, lefts, 1, "Hours", "Gray"
    d = DateSerial(2099, 1, 1) + TimeSerial(22, 10, 0)
    targetA = COL_PAD + UnitFracOf(d, "Hours") * (960 - 2 * COL_PAD)
    Set a = CreateTimelineEntry(sld, "Jan 1 10:10 PM", "Line nudge stays within this card", 100, 140, 180, _
                                190, BAND_TOP + BAND_HEIGHT, True, 1, h, d, "Hours", "", True, True)
    a.Tags.Add "TLENTRY", "1": a.Tags.Add "TLFullDate", CStr(CDbl(d))
    d = DateSerial(2099, 1, 1) + TimeSerial(22, 50, 0)
    targetB = COL_PAD + UnitFracOf(d, "Hours") * (960 - 2 * COL_PAD)
    Set b = CreateTimelineEntry(sld, "Jan 1 10:50 PM", "Line nudge clamps at this card edge", 600, 280, 180, _
                                690, BAND_TOP + BAND_HEIGHT, True, 1, h, d, "Hours", "", True, True)
    b.Tags.Add "TLENTRY", "1": b.Tags.Add "TLFullDate", CStr(CDbl(d))
    Set untimed = CreateTimelineEntry(sld, "Jan 1", "Untimed entry stays in place", 400, 410, 180, _
                                      490, BAND_TOP + BAND_HEIGHT, True, 1, h, DateSerial(2099, 1, 1), "Hours", "", False, False)
    untimed.Tags.Add "TLENTRY", "1": untimed.Tags.Add "TLFullDate", CStr(CDbl(DateSerial(2099, 1, 1)))
    untimedLeft = untimed.Left
    Set aBox = DateBoxOf(a): Set bBox = DateBoxOf(b)
    aLeft = aBox.Left: bLeft = bBox.Left
    TQTestLeaderGeometry sld, "leader edit initial", fh, checks, failures
    ' The test-only fixture sends DateSnapCore to this slide and answers its prompts.
    sld.Tags.Add "TQTestPromptAnswer", "NO"
    DateSnapCore False
    Set ln = LeadingLineOf(a)
    TQTestCheck fh, Abs(ln.Left - targetA) < 0.2, "line nudge reaches a date within card bounds", checks, failures
    Set ln = LeadingLineOf(b)
    TQTestCheck fh, Abs(ln.Left - bBox.Left - bBox.Width) < 0.2, "line nudge clamps at card edge when group move declined", checks, failures
    TQTestCheck fh, Abs(aBox.Left - aLeft) < 0.2 And Abs(bBox.Left - bLeft) < 0.2, "line nudge leaves both cards in place", checks, failures
    TQTestLeaderGeometry sld, "line nudge clamped", fh, checks, failures
    sld.Tags.Add "TQTestPromptAnswer", "YES"
    DateSnapCore False
    TQTestCheck fh, bBox.Left > bLeft + 1, "line nudge optional group move shifts the unreachable card", checks, failures
    TQTestLeaderGeometry sld, "line nudge group-move fallback", fh, checks, failures, True
    a.Left = a.Left + 15: b.Left = b.Left - 15
    DateSnapCore True
    TQTestLeaderGeometry sld, "Date Snap group move", fh, checks, failures, True
    a.Left = a.Left + 15: b.Left = b.Left - 15
    Set bar = FindDateBar(sld)
    ReflowTimeline sld, bar, "Hours"
    TQTestLeaderGeometry sld, "leader edit reflow", fh, checks, failures, True
    TQTestCheck fh, Abs(untimed.Left - untimedLeft) < 0.2, "Date Snap and reflow leave untimed entries in place", checks, failures
    Set ln = LeadingLineOf(untimed)
    TQTestCheck fh, ln Is Nothing, "Date Snap and reflow do not add untimed leaders", checks, failures
    sld.Export outputDir & "\leader-edits.png", "PNG", 1600, 900
End Sub

Private Function TQTestMessageBox(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                  Optional ByVal title As String = "") As VbMsgBoxResult
    If (buttons And vbYesNo) = vbYesNo Then
        If TQTestTargetPresentation().Slides("leader-edits").Tags("TQTestPromptAnswer") = "YES" Then
            TQTestMessageBox = vbYes
        Else
            TQTestMessageBox = vbNo
        End If
    Else
        TQTestMessageBox = vbOK
    End If
End Function

Private Sub TQTestColumnLeaders(ByVal outputDir As String, ByVal fh As Integer, ByRef checks As Long, ByRef failures As Long)
    Dim ev(1 To 6) As TLEvent, i As Long, mode As Long, laneCount As Long, sld As slide, entry As Shape, ln As Shape, db As Shape
    Dim cols(1 To 1) As Date, widths(1 To 1) As Single, lefts(1 To 1) As Single, entries As Collection
    Dim count As Long, animSeq As Long, expectedX As Single, entryDate As Date, eventIndex As Long
    cols(1) = DateSerial(2099, 1, 1) + TimeSerial(22, 0, 0): widths(1) = 960
    For i = 1 To 6
        ev(i).RawDate = cols(1) + TimeSerial(0, 5 + (i - 1) * 10, 0)
        ev(i).DateLabel = Format$(ev(i).RawDate, "mmm d h:nn AM/PM")
        ev(i).Desc = "Entry " & i
        ev(i).Prec = 4: ev(i).HasTime = True: ev(i).OrigIndex = i
    Next i
    PrepareEventPositions ev, 6
    PrepareEventUnits ev, 6, "Hours"
    For mode = 1 To 2
        If mode = 1 Then laneCount = 1 Else laneCount = 3
        Set sld = TQTestNewSlide("column-leaders-" & laneCount)
        DrawBar sld, cols, 1, widths, lefts, 1, "Hours", "Gray"
        count = DrawColumn(sld, ev, 6, cols(1), 0, 960, BAND_TOP + BAND_HEIGHT, 0.7, False, animSeq, laneCount, "Hours")
        TQTestCheck fh, count = 6, laneCount & "-lane column renders all six entries", checks, failures
        Set entries = New Collection
        CollectEntryGroups sld.Shapes, entries
        For Each entry In entries
            Set ln = LeadingLineOf(entry): Set db = DateBoxOf(entry)
            If Not ln Is Nothing And Not db Is Nothing Then
                If StoredEntryDate(entry, entryDate) Then
                    eventIndex = DateDiff("n", cols(1), entryDate)
                    eventIndex = (eventIndex - 5) \ 10
                    If laneCount = 1 Then
                        expectedX = COL_PAD + UnitFracOf(entryDate, "Hours") * (960 - 2 * COL_PAD)
                    Else
                        expectedX = COL_PAD + (eventIndex \ 2 + 0.5) * ((960 - 2 * COL_PAD) / 3)
                        TQTestCheck fh, Abs(ln.Left - db.Left - db.Width / 2) < 0.2, "multi-lane leader remains centered within its card", checks, failures
                    End If
                    TQTestCheck fh, Abs(ln.Left - expectedX) < 0.2, laneCount & "-lane column retains its original horizontal anchor", checks, failures
                Else
                    TQTestCheck fh, False, "column leader has a stored date", checks, failures
                End If
            Else
                TQTestCheck fh, False, "column entry has a leader and date box", checks, failures
            End If
        Next entry
        TQTestLeaderGeometry sld, laneCount & "-lane column", fh, checks, failures, (laneCount = 1)
        sld.Export outputDir & "\column-leaders-" & laneCount & ".png", "PNG", 1600, 900
    Next mode
End Sub

Private Sub TQTestBatesRebuild(ByVal path As String, ByVal outputDir As String, ByVal fh As Integer, _
                              ByRef checks As Long, ByRef failures As Long)
    Dim ev() As TLEvent, restored() As TLEvent, n As Long, rn As Long, errors As String, summary As String, sld As slide
    Dim t As String, color As String, gaps As Boolean, multi As Boolean, wipe As Boolean, weighted As Boolean
    Dim i As Long, expected As Long, shown As Long, entries As Collection, entry As Shape
    n = ReadEvents(path, ev, errors)
    If n = 0 Then TQTestCheck fh, False, "Bates rebuild imports entries", checks, failures: Exit Sub
    For i = 1 To n: If Len(ev(i).Bates) > 0 Then expected = expected + 1
    Next i
    Set sld = TQTestNewSlide("bates-visible-rebuild")
    sld.Tags.Add "TLBatesVisible", "1"
    TQTestCheck fh, RenderTimeline(sld, ev, n, "Hours", "Gray", False, False, False, False, summary), _
                "Bates-visible timeline renders with measured footers", checks, failures
    TQTestEntryBounds sld, "Bates-visible initial", fh, checks, failures
    TQTestLeaderGeometry sld, "Bates-visible initial", fh, checks, failures
    TQTestCheck fh, LoadTimelineState(sld, restored, rn, t, color, gaps, multi, wipe, weighted), _
                "Bates-visible timeline reloads state", checks, failures
    TQTestCheck fh, RenderTimeline(sld, restored, rn, t, color, gaps, multi, wipe, weighted, summary), _
                "Bates-visible timeline rerenders from saved state", checks, failures
    TQTestCheck fh, TimelineBatesAreVisible(sld), "Bates-visible rebuild preserves visibility setting", checks, failures
    Set entries = New Collection
    CollectEntryGroups sld.Shapes, entries
    For Each entry In entries
        If TimelineEntryBatesVisible(entry) Then shown = shown + 1
    Next entry
    TQTestCheck fh, entries.Count = n, "Bates-visible rebuild keeps every entry", checks, failures
    TQTestCheck fh, shown = expected And expected > 0, "Bates-visible rebuild restores every supplied footer", checks, failures
    TQTestEntryBounds sld, "Bates-visible rebuild", fh, checks, failures
    TQTestLeaderGeometry sld, "Bates-visible rebuild", fh, checks, failures
    sld.Export outputDir & "\bates-visible-rebuild.png", "PNG", 1600, 900
End Sub

Private Sub TQTestBates(ByVal outputDir As String, ByVal fh As Integer, ByRef checks As Long, ByRef failures As Long)
    Dim sld As slide, entry As Shape, box As Shape, h As Single, body As String, bates As String
    Dim footer As TextRange, i As Long, originalBody As String
    body = "Dr. Example" & vbCr & "The patient reports symptoms. The note records examination and treatment."
    bates = "aBc-000123"
    Set sld = TQTestNewSlide("bates-format")
    Set entry = CreateTimelineEntry(sld, "Jan 1", body, 72, 110, 320, 110, 70, True, 1, h, DateSerial(2099, 1, 1), "Days")
    Set box = EntryBoxOf(entry)
    box.TextFrame.TextRange.Characters(1, 11).Font.Bold = msoTrue
    box.TextFrame.TextRange.Characters(1, 11).Font.Size = 16
    box.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment = ppAlignCenter
    originalBody = box.TextFrame.TextRange.Text
    SetTimelineEntryBates entry, bates, True
    TQTestCheck fh, box.TextFrame.TextRange.Text = originalBody & vbCr & bates, _
                "Bates appends one exact-case paragraph", checks, failures
    TQTestCheck fh, TimelineEntryBodyText(entry) = originalBody, "Bates is excluded from body extraction", checks, failures
    Set footer = box.TextFrame.TextRange.Characters(Len(originalBody) + 2, Len(bates))
    TQTestCheck fh, footer.ParagraphFormat.Alignment = ppAlignRight, "Bates paragraph is right aligned", checks, failures
    TQTestCheck fh, footer.Font.Italic = msoTrue, "Bates is italic", checks, failures
    TQTestCheck fh, footer.Font.Bold = msoFalse, "Bates is not bold", checks, failures
    TQTestCheck fh, footer.Font.Color.RGB = RGB(128, 128, 128), "Bates is mid gray", checks, failures
    TQTestCheck fh, Abs(footer.Font.Size - 7.2) < 0.1, "Bates is 60 percent of 12 pt body", checks, failures
    For i = 1 To 3
        ApplyTimelineEntryBates entry, True
        TQTestCheck fh, box.TextFrame.TextRange.Text = originalBody & vbCr & bates, _
                    "Bates repeat show does not duplicate suffix", checks, failures
        ApplyTimelineEntryBates entry, False
        TQTestCheck fh, box.TextFrame.TextRange.Text = originalBody, "Bates hide restores exact body", checks, failures
        TQTestCheck fh, box.TextFrame.TextRange.Characters(1, 11).Font.Bold = msoTrue, _
                    "Bates toggle preserves bold title", checks, failures
        TQTestCheck fh, box.TextFrame.TextRange.Characters(1, 11).Font.Size = 16, _
                    "Bates toggle preserves title size", checks, failures
        TQTestCheck fh, box.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Alignment = ppAlignCenter, _
                    "Bates toggle preserves title alignment", checks, failures
    Next i
    ApplyTimelineEntryBates entry, True
    sld.Export outputDir & "\bates-format.png", "PNG", 1600, 900
End Sub

Public Sub TQTimelineRenderRegression(ByVal correctedPath As String, ByVal originalPath As String, _
                                     ByVal reportPath As String, ByVal outputDir As String)
    Dim fh As Integer, checks As Long, failures As Long, deck As Presentation, i As Long
    Dim fraction As Double, label As String, accepted As Boolean
    On Error GoTo Fatal
    fh = FreeFile
    Open reportPath For Output As #fh
    Print #fh, "TIMELINE REGRESSION | " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set deck = TQTestTargetPresentation()
    For i = deck.Slides.Count To 1 Step -1: deck.Slides(i).Delete: Next i
    deck.PageSetup.SlideWidth = 960
    deck.PageSetup.SlideHeight = 540
    TQTestImportRender correctedPath, "corrected-hours", 30, 0, outputDir, fh, checks, failures
    TQTestImportRender originalPath, "original-mixed", 37, 7, outputDir, fh, checks, failures
    TQTestPagedRender correctedPath, "corrected-paged", 30, 0, outputDir, fh, checks, failures
    TQTestPagedRender originalPath, "original-paged", 37, 7, outputDir, fh, checks, failures
    TQTestDateOnly outputDir, fh, checks, failures
    TQTestColumnLeaders outputDir, fh, checks, failures
    TQTestLeaderEdits outputDir, fh, checks, failures
    TQTestBates outputDir, fh, checks, failures
    TQTestBatesRebuild correctedPath, outputDir, fh, checks, failures
    accepted = TryParseTimeCell("10:30 PM - 11:15 PM", fraction, label)
    TQTestCheck fh, accepted, "time range parses", checks, failures
    TQTestCheck fh, Abs(fraction - CDbl(TimeSerial(22, 30, 0))) < 0.0000001, _
                "time range anchors to start", checks, failures
    TQTestCheck fh, InStr(label, "11:15") > 0, "time range keeps ending time in label", checks, failures
    accepted = TryParseTimeCell("11:30 PM - 12:15 AM", fraction, label)
    TQTestCheck fh, accepted And InStr(label, "12:15") > 0, "overnight time range parses and retains end", checks, failures
    accepted = TryParseTimeCell("00:00", fraction, label)
    TQTestCheck fh, accepted And Abs(fraction) < 0.0000001, "explicit midnight parses as valid time", checks, failures
    Dim originalTimeText As String, swappedTimeText As String
    originalTimeText = "Jan 1, 2099 10:47:35 PM - 10:56:28 PM"
    swappedTimeText = SwapTimeFormat.SwapFormat(originalTimeText)
    TQTestCheck fh, swappedTimeText = "Jan 1, 2099 22:47:35 - 22:56:28", _
                "SwapFormat preserves date and seconds while toggling both range endpoints", checks, failures
    TQTestCheck fh, SwapTimeFormat.SwapFormat(swappedTimeText) = originalTimeText, _
                "SwapFormat restores date and both full range endpoints on second toggle", checks, failures
    TQTestCheck fh, SwapTimeFormat.SwapFormat("Jan 1, 2099 23:50:05 - 00:10:07") = "Jan 1, 2099 11:50:05 PM - 12:10:07 AM", _
                "SwapFormat handles overnight range endpoints with seconds", checks, failures
    Print #fh, "RESULT | " & IIf(failures = 0, "PASS", "FAIL") & " | checks=" & checks & " | failures=" & failures
    Close #fh
    deck.Save
    Exit Sub
Fatal:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If fh > 0 Then
        Print #fh, "ERROR | " & errorNumber & " | " & errorText
        Print #fh, "RESULT | ERROR | checks=" & checks & " | failures=" & failures
        Close #fh
    End If
End Sub
'@

$pptApplication = $null
$testPresentation = $null
$sourcePresentation = $null
$ownsApplication = $false
try {
    try { $pptApplication = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
    catch { $pptApplication = New-Object -ComObject PowerPoint.Application; $ownsApplication = $true }
    foreach ($openDeck in $pptApplication.Presentations) {
        if ($openDeck.Name -ieq 'TimelineRegression.pptm') {
            throw 'A previous TimelineRegression.pptm is still open. Close that test deck first.'
        }
    }
    $previousSecurity = $pptApplication.AutomationSecurity
    Write-Host ('Existing PowerPoint AutomationSecurity: ' + $previousSecurity)
    try {
        $pptApplication.AutomationSecurity = 3
        $sourcePresentation = $pptApplication.Presentations.Open($sourceCopyPath, -1, 0, 0)
    }
    finally { $pptApplication.AutomationSecurity = $previousSecurity }
    $sourceProject = $sourcePresentation.VBProject
    if ($null -eq $sourceProject) { throw 'Source copy VBProject is null; trusted VBOM access is required.' }
    if ([int]$sourceProject.Protection -ne 0) { throw 'Source copy VBA project is protected.' }
    $referenceSpecs = @()
    foreach ($reference in $sourceProject.References) {
        if ($reference.IsBroken) { throw ('Source project has a broken reference: ' + $reference.Name) }
        $referenceSpecs += [pscustomobject]@{
            Name = [string]$reference.Name
            Guid = [string]$reference.GUID
            Major = [int]$reference.Major
            Minor = [int]$reference.Minor
        }
    }
    $exportedComponents = @()
    for ($index = 1; $index -le $sourceProject.VBComponents.Count; $index++) {
        $sourceComponent = $sourceProject.VBComponents.Item($index)
        switch ([int]$sourceComponent.Type) {
            1 { $extension = '.bas' }
            2 { $extension = '.cls' }
            3 { $extension = '.frm' }
            default { throw ('Unsupported original component type: ' + $sourceComponent.Name + ' (' + $sourceComponent.Type + ')') }
        }
        $exportPath = Join-Path $componentDirectory ($sourceComponent.Name + $extension)
        $sourceComponent.Export($exportPath)
        $exportedComponents += $exportPath
    }
    $sourcePresentation.Close()
    $sourcePresentation = $null

    # A newly created project runs explicitly injected tests without reopening a
    # document whose macros were disabled. Trust Center settings stay unchanged.
    $testPresentation = $pptApplication.Presentations.Add(0)
    $project = $testPresentation.VBProject
    if ($null -eq $project) { throw 'New test presentation VBProject is null.' }
    $components = $project.VBComponents
    foreach ($referenceSpec in $referenceSpecs) {
        $present = $false
        foreach ($reference in $project.References) {
            if ([string]$reference.GUID -ieq $referenceSpec.Guid) { $present = $true; break }
        }
        if (-not $present) {
            [void]$project.References.AddFromGuid($referenceSpec.Guid, $referenceSpec.Major, $referenceSpec.Minor)
        }
    }
    foreach ($exportPath in $exportedComponents) { [void]$components.Import($exportPath) }
    Write-Host ('Preserved ' + $exportedComponents.Count + ' original components and ' + $referenceSpecs.Count + ' references')
    $modulePaths = @(& git -c "safe.directory=$repoRoot" -C $repoRoot ls-files --cached --others --exclude-standard -- 'Modules/*.bas')
    if ($LASTEXITCODE -ne 0 -or $modulePaths.Count -eq 0) { throw 'Cannot enumerate Modules/*.bas.' }
    foreach ($relativePath in ($modulePaths | Sort-Object -Unique)) {
        $sourcePath = Join-Path $repoRoot $relativePath
        $sourceText = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::Default)
        $nameMatch = [regex]::Match($sourceText, '(?m)^Attribute VB_Name = "([^"]+)"\s*$')
        if (-not $nameMatch.Success) { throw "Module name missing in $sourcePath" }
        $moduleName = $nameMatch.Groups[1].Value
        for ($index = $components.Count; $index -ge 1; $index--) {
            $component = $components.Item($index)
            if ($component.Name -ieq $moduleName) {
                if ([int]$component.Type -ne 1) { throw "Refusing to replace nonstandard component $moduleName" }
                $components.Remove($component)
            }
        }
        [void]$components.Import($sourcePath)
    }
    $timelineCode = $components.Item('TimelineCreator').CodeModule
    $productionText = $timelineCode.Lines(1, $timelineCode.CountOfLines)
    $fixtureText = [regex]::Replace($productionText, '(?<![\w.])ActivePresentation\b', 'TQTestTargetPresentation()')
    # Run the real Date Snap logic against an isolated fixture with scripted dialog answers.
    # These substitutions exist only in the temporary regression presentation.
    $snapMatch = [regex]::Match($fixtureText, '(?ms)^Private Sub DateSnapCore\(.*?^End Sub')
    if (-not $snapMatch.Success) { throw 'Cannot locate DateSnapCore for isolated regression coverage.' }
    $snapFixture = $snapMatch.Value.Replace('Set sld = ActiveTargetSlide()', 'Set sld = TQTestTargetPresentation().Slides("leader-edits")')
    $snapFixture = [regex]::Replace($snapFixture, '\bMsgBox\b', 'TQTestMessageBox')
    $fixtureText = $fixtureText.Substring(0, $snapMatch.Index) + $snapFixture + $fixtureText.Substring($snapMatch.Index + $snapMatch.Length)
    $timelineCode.DeleteLines(1, $timelineCode.CountOfLines)
    $timelineCode.AddFromString($fixtureText + "`r`n" + $testCode)
    $correctedLiteral = $CorrectedWorkbook.Replace('"', '""')
    $originalLiteral = $OriginalWorkbook.Replace('"', '""')
    $reportLiteral = $reportPath.Replace('"', '""')
    $outputLiteral = $outputDirectory.Replace('"', '""')
    $smokeLiteral = (Join-Path $outputDirectory 'smoke.txt').Replace('"', '""')
    $runnerCode = @"
Option Explicit

Public Sub TQRunRenderToFile()
    Dim fh As Integer, errorNumber As Long, errorText As String
    On Error GoTo Failed
    fh = FreeFile
    Open "$reportLiteral" For Output As #fh
    Print #fh, "STARTED | no-argument caller before TimelineCreator call"
    Close #fh
    TimelineCreator.TQTimelineRenderRegression "$correctedLiteral", "$originalLiteral", "$reportLiteral", "$outputLiteral"
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    fh = FreeFile
    Open "$reportLiteral" For Append As #fh
    Print #fh, "CALLER ERROR | " & errorNumber & " | " & errorText
    Close #fh
End Sub
"@
    $runnerComponent = $components.Add(1)
    $runnerComponent.Name = 'TQRenderHarness'
    $runnerComponent.CodeModule.AddFromString($runnerCode)
    $smokeCode = @"
Option Explicit
Public Sub TQRenderSmoke()
    Dim fh As Integer
    fh = FreeFile
    Open "$smokeLiteral" For Output As #fh
    Print #fh, "PASS | independent no-argument macro executed"
    Close #fh
End Sub
"@
    $smokeComponent = $components.Add(1)
    $smokeComponent.Name = 'TQRenderProbe'
    $smokeComponent.CodeModule.AddFromString($smokeCode)
    $inventory = @()
    for ($index = 1; $index -le $components.Count; $index++) {
        $component = $components.Item($index)
        $inventory += ('{0} | type={1} | codeLines={2}' -f $component.Name, $component.Type, $component.CodeModule.CountOfLines)
    }
    $inventory | Set-Content -LiteralPath (Join-Path $outputDirectory 'component-inventory.txt') -Encoding UTF8
    [IO.File]::WriteAllText((Join-Path $outputDirectory 'TQRenderHarness-code.txt'), $runnerCode)
    [IO.File]::WriteAllText((Join-Path $outputDirectory 'TimelineCreator-test-code.txt'), ($fixtureText + "`r`n" + $testCode))
    $testPresentation.SaveAs($testDeckPath, 25) # ppSaveAsOpenXMLPresentationMacroEnabled
    $macro = $testPresentation.Name + '!TQRenderHarness.TQRunRenderToFile'
    Write-Host ('Running ' + $macro)
    # PowerPoint Run takes ref object[]; PowerShell's direct COM binder cannot
    # pass it correctly. Match the installed interop signature through C#.
    if (-not ('TimelinePowerPointRenderRunner' -as [type])) {
        $interop = [Reflection.Assembly]::Load('Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
        Add-Type -ReferencedAssemblies $interop.Location -TypeDefinition @'
public static class TimelinePowerPointRenderRunner {
    public static object Run(object application, string macroName) {
        object[] arguments = new object[0];
        return ((Microsoft.Office.Interop.PowerPoint._Application)application).Run(macroName, ref arguments);
    }
}
'@
    }
    [void][TimelinePowerPointRenderRunner]::Run($pptApplication, ($testPresentation.Name + '!TQRenderProbe.TQRenderSmoke'))
    Write-Host 'Independent no-argument smoke macro succeeded.'
    [void][TimelinePowerPointRenderRunner]::Run($pptApplication, $macro)
    $testPresentation.Save()
}
catch {
    $originalError = $_
    $diagnostics = @('PowerPoint invocation failed: ' + $_.Exception.Message)
    if ($null -ne $pptApplication) {
        try {
            $activePane = $pptApplication.VBE.ActiveCodePane
            if ($null -ne $activePane) {
                $activeCode = $activePane.CodeModule
                $diagnostics += ('Active VBE component: ' + $activeCode.Name)
                [int]$startLine = 0; [int]$startColumn = 0; [int]$endLine = 0; [int]$endColumn = 0
                $activePane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
                $diagnostics += ('Selection: line {0}:{1} through {2}:{3}' -f $startLine, $startColumn, $endLine, $endColumn)
                if ($startLine -gt 0) {
                    $firstLine = [Math]::Max(1, $startLine - 3)
                    $lineCount = [Math]::Min(10, $activeCode.CountOfLines - $firstLine + 1)
                    $diagnostics += $activeCode.Lines($firstLine, $lineCount)
                }
            }
        }
        catch { $diagnostics += ('VBE selection unavailable: ' + $_.Exception.Message) }
    }
    $diagnostics | Set-Content -LiteralPath (Join-Path $outputDirectory 'invocation-diagnostics.txt') -Encoding UTF8
    Write-Output ($diagnostics -join "`r`n")
    throw $originalError
}
finally {
    if ($null -ne $sourcePresentation) {
        try { $sourcePresentation.Close() }
        catch { Write-Warning ('Could not close read-only source copy: ' + $_.Exception.Message) }
    }
    if ($null -ne $testPresentation) {
        try { $testPresentation.Saved = -1; $testPresentation.Close() }
        catch { Write-Warning ('Could not close test deck: ' + $_.Exception.Message) }
    }
    if ($ownsApplication -and $null -ne $pptApplication) {
        try { if ($pptApplication.Presentations.Count -eq 0) { $pptApplication.Quit() } }
        catch { Write-Warning ('Could not close script-owned PowerPoint: ' + $_.Exception.Message) }
    }
    Write-Host ('Test artifacts: ' + $outputDirectory)
}
if (-not (Test-Path -LiteralPath $reportPath)) {
    throw 'VBA did not create the test report. Inspect compile errors and existing macro security; no trust settings were changed.'
}
$report = Get-Content -LiteralPath $reportPath -Raw
Write-Output $report
if ($report -notmatch '(?m)^RESULT \| PASS \|') { throw "Timeline regressions failed. See $reportPath" }
