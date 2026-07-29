Attribute VB_Name = "TitleManager"
Option Explicit

' ============================================================================
' TitleManager - the Timeline Title Manager engine (ribbon: Timeline Edits).
' ----------------------------------------------------------------------------
' A "title" is the bold first line of a timeline entry's white Entry Box (the
' provider/facility line that Parse Content produces). This tool scans every
' timeline slide in the deck, groups entries by that title, and recolors each
' entry's navy DATE BOX to the title's assigned color - with the date-box text
' auto-switched white/black by brightness so it stays readable.
'
' The registry + palettes live in TitleColors.bas; the UI in frmTitleManager.
' Detection helpers here are self-contained copies (the TimelineCreator ones are
' Private) so this module stands alone.
' ============================================================================

' --- ribbon entry -----------------------------------------------------------
Public Sub OpenTitleManager(control As IRibbonControl)
    On Error GoTo Fail
    If Not AnyTimelineSlide() Then
        MsgBox "No timeline slides were found in this presentation." & vbCrLf & vbCrLf & _
               "Make a timeline (Make Timeline / Make DateBar) first.", vbInformation, "Title Manager"
        Exit Sub
    End If

    Dim recs As Collection
    Set recs = DistinctTitles()
    If recs.count = 0 Then
        MsgBox "No titled entries were found." & vbCrLf & vbCrLf & _
               "A 'title' is an entry whose first line is bold - the provider/facility line " & _
               "that Parse Content creates. Run Parse Content on your entries first, then reopen " & _
               "the Title Manager.", vbInformation, "Title Manager"
        Exit Sub
    End If

    ' register every scanned title so its state persists even before it is colored
    Dim reg As Object, rec As Variant
    Set reg = TitleColors.LoadTitleReg()
    TitleColors.EnsureTitleSettings reg
    For Each rec In recs
        TitleColors.EnsureTitle reg, CStr(rec(0))
    Next rec
    TitleColors.SaveTitleReg reg

    frmTitleManager.Show
    Exit Sub
Fail:
    MsgBox "Title Manager failed to open:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Title Manager"
End Sub

' ============================================================================
' SCAN
' ============================================================================
' Every titled entry across all timeline slides, as Array(slideIndex, groupName,
' titleKey, rawTitle, dateText, descText).
Public Function GetTitledEntries() As Collection
    Dim col As Collection, sld As slide, acc As Collection, grpV As Variant, grp As Shape
    Dim eb As Shape, db As Shape, raw As String, key As String, dt As String, ds As String
    Set col = New Collection
    For Each sld In ActivePresentation.Slides
        If IsTimelineSlide(sld) Then
            Set acc = New Collection
            CollectEntryGroups sld.Shapes, acc
            For Each grpV In acc
                Set grp = grpV
                Set eb = FindTaggedDescendant(grp, "Entry Box")
                If Not eb Is Nothing Then
                    If IsTitledEntryBox(eb) Then
                        raw = TitleOfEntryBox(eb)
                        key = TitleColors.SanitizeTitleName(raw)
                        If key <> "" Then
                            Set db = FindTaggedDescendant(grp, "Date Box")
                            dt = ""
                            If Not db Is Nothing Then dt = FlattenText(SafeText(db))
                            ds = BodyTextOf(eb)
                            col.Add Array(sld.SlideIndex, grp.name, key, raw, dt, ds)
                        End If
                    End If
                End If
            Next grpV
        End If
    Next sld
    Set GetTitledEntries = col
End Function

' Distinct titles as Array(titleKey, rawTitle, count), ordered count-desc then A-Z
' (so recurring titles get the prime palette colors during Auto-Assign).
Public Function DistinctTitles() As Collection
    Dim entries As Collection, e As Variant, d As Object, raws As Object
    Dim key As String, res As Collection
    Set entries = GetTitledEntries()
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set raws = CreateObject("Scripting.Dictionary")
    raws.CompareMode = vbTextCompare
    For Each e In entries
        key = CStr(e(2))
        If d.Exists(key) Then
            d(key) = CLng(d(key)) + 1
        Else
            d(key) = 1
            raws(key) = CStr(e(3))
        End If
    Next e

    ' build + sort (insertion sort by count desc, then key asc)
    Dim keys() As String, cnts() As Long, n As Long, k As Variant, i As Long, j As Long
    n = d.count
    Set res = New Collection
    If n = 0 Then Set DistinctTitles = res: Exit Function
    ReDim keys(0 To n - 1)
    ReDim cnts(0 To n - 1)
    i = 0
    For Each k In d.Keys
        keys(i) = CStr(k)
        cnts(i) = CLng(d(k))
        i = i + 1
    Next k
    Dim tk As String, tc As Long
    For i = 1 To n - 1
        tk = keys(i): tc = cnts(i): j = i - 1
        Do While j >= 0
            If cnts(j) > tc Then Exit Do
            If cnts(j) = tc And StrComp(keys(j), tk, vbTextCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j): cnts(j + 1) = cnts(j): j = j - 1
        Loop
        keys(j + 1) = tk: cnts(j + 1) = tc
    Next i
    For i = 0 To n - 1
        res.Add Array(keys(i), CStr(raws(keys(i))), cnts(i))
    Next i
    Set DistinctTitles = res
End Function

' The entries of one title, as Array(slideIndex, dateText, descText), for the UI.
Public Function EntriesOfTitle(ByVal key As String) As Collection
    Dim entries As Collection, e As Variant, res As Collection
    Set res = New Collection
    Set entries = GetTitledEntries()
    For Each e In entries
        If StrComp(CStr(e(2)), key, vbTextCompare) = 0 Then
            res.Add Array(e(0), e(4), e(5))
        End If
    Next e
    Set EntriesOfTitle = res
End Function

' ============================================================================
' APPLY  (single source of truth: recolor every titled entry's date box)
' ============================================================================
Public Sub ApplyAllColors(reg As Object)
    Dim sld As slide, acc As Collection, grpV As Variant, grp As Shape
    Dim eb As Shape, db As Shape, key As String, c As Long
    For Each sld In ActivePresentation.Slides
        If IsTimelineSlide(sld) Then
            Set acc = New Collection
            CollectEntryGroups sld.Shapes, acc
            For Each grpV In acc
                Set grp = grpV
                Set eb = FindTaggedDescendant(grp, "Entry Box")
                If Not eb Is Nothing Then
                    If IsTitledEntryBox(eb) Then
                        Set db = FindTaggedDescendant(grp, "Date Box")
                        If Not db Is Nothing Then
                            key = TitleColors.SanitizeTitleName(TitleOfEntryBox(eb))
                            If TitleColors.HasColor(reg, key) Then
                                c = TitleColors.GetTitleColor(reg, key)
                            Else
                                c = TitleColors.DefaultNavy()
                            End If
                            SetDateBoxColor db, c
                        End If
                    End If
                End If
            Next grpV
        End If
    Next sld
End Sub

' Fill = color; text ink = white on dark colors, black on light (perceptual luma).
Public Sub SetDateBoxColor(ByVal db As Shape, ByVal c As Long)
    Dim ink As Long
    On Error Resume Next
    db.Fill.Visible = msoTrue
    db.Fill.Solid
    db.Fill.ForeColor.RGB = c
    If Brightness(c) < 145 Then ink = RGB(255, 255, 255) Else ink = RGB(0, 0, 0)
    db.TextFrame.TextRange.Font.color.RGB = ink
    On Error GoTo 0
End Sub

Private Function Brightness(ByVal c As Long) As Double
    Dim r As Long, g As Long, b As Long
    r = c And &HFF
    g = (c \ &H100) And &HFF
    b = (c \ &H10000) And &HFF
    Brightness = 0.299 * r + 0.587 * g + 0.114 * b
End Function

' ============================================================================
' AUTO-ASSIGN  (random, cool colors first, preserving Manual titles)
' ============================================================================
Public Sub AutoAssign(reg As Object)
    Dim recs As Collection, rec As Variant, key As String, cnt As Long
    Dim pIdx As Long, uniqueSingles As Boolean, order() As Long, recurIdx As Long, col As Long
    Set recs = DistinctTitles()
    pIdx = TitleColors.PaletteIndexOf(TitleColors.GetPalette(reg))
    uniqueSingles = TitleColors.GetSingletonUnique(reg)

    Randomize
    order = ShuffledPaletteOrder(pIdx)     ' cool block (shuffled) then warm block (shuffled)
    recurIdx = 0

    For Each rec In recs
        key = CStr(rec(0))
        cnt = CLng(rec(2))
        If Not TitleColors.GetTitleManual(reg, key) Then      ' preserve hand-set titles
            If cnt >= 2 Or uniqueSingles Then
                If recurIdx < TitleColors.PALETTE_SIZE Then
                    col = TitleColors.TitlePaletteColor(pIdx, order(recurIdx))
                Else
                    col = RandomCoolColor()                   ' palette exhausted -> accessible cool color
                End If
                TitleColors.SetTitleColor reg, key, col
                TitleColors.SetTitleManual reg, key, False
                recurIdx = recurIdx + 1
            Else
                TitleColors.SetTitleColor reg, key, TitleColors.SingletonColor()
                TitleColors.SetTitleManual reg, key, False
            End If
        End If
    Next rec
End Sub

' 0..11 with the palette's cool indices (shuffled) first, then the warm tail (shuffled).
Private Function ShuffledPaletteOrder(ByVal pIdx As Long) As Long()
    Dim cc As Long, i As Long, res() As Long
    ReDim res(0 To TitleColors.PALETTE_SIZE - 1)
    cc = TitleColors.CoolCount(pIdx)
    If cc < 1 Then cc = 1
    If cc > TitleColors.PALETTE_SIZE Then cc = TitleColors.PALETTE_SIZE
    Dim cool() As Long, warm() As Long
    ReDim cool(0 To cc - 1)
    For i = 0 To cc - 1: cool(i) = i: Next i
    ShuffleLongs cool
    Dim wn As Long
    wn = TitleColors.PALETTE_SIZE - cc
    Dim w As Long
    For i = 0 To cc - 1: res(i) = cool(i): Next i
    If wn > 0 Then
        ReDim warm(0 To wn - 1)
        For i = 0 To wn - 1: warm(i) = cc + i: Next i
        ShuffleLongs warm
        For i = 0 To wn - 1: res(cc + i) = warm(i): Next i
    End If
    ShuffledPaletteOrder = res
End Function

' Fisher-Yates on a 0-based Long array (assumes Randomize already called).
Private Sub ShuffleLongs(ByRef a() As Long)
    Dim i As Long, j As Long, tmp As Long, hi As Long
    hi = UBound(a)
    For i = hi To 1 Step -1
        j = Int(Rnd * (i + 1))
        tmp = a(i): a(i) = a(j): a(j) = tmp
    Next i
End Sub

' A random cool color (hue 150-285 deg), moderate saturation/value.
Private Function RandomCoolColor() As Long
    Dim h As Double, s As Double, v As Double
    h = 150 + Rnd * 135
    s = 0.45 + Rnd * 0.35
    v = 0.4 + Rnd * 0.45
    RandomCoolColor = HSVtoRGB(h, s, v)
End Function

Private Function HSVtoRGB(ByVal h As Double, ByVal s As Double, ByVal v As Double) As Long
    Dim c As Double, x As Double, m As Double, r As Double, g As Double, b As Double
    Dim hp As Double
    Do While h < 0: h = h + 360: Loop
    Do While h >= 360: h = h - 360: Loop
    hp = h / 60#
    c = v * s
    x = c * (1 - Abs((hp - 2 * Int(hp / 2)) - 1))
    m = v - c
    Select Case Int(hp)
        Case 0: r = c: g = x: b = 0
        Case 1: r = x: g = c: b = 0
        Case 2: r = 0: g = c: b = x
        Case 3: r = 0: g = x: b = c
        Case 4: r = x: g = 0: b = c
        Case Else: r = c: g = 0: b = x
    End Select
    HSVtoRGB = RGB(CLng((r + m) * 255), CLng((g + m) * 255), CLng((b + m) * 255))
End Function

' ============================================================================
' DETECTION HELPERS (self-contained)
' ============================================================================
Public Function AnyTimelineSlide() As Boolean
    Dim sld As slide
    For Each sld In ActivePresentation.Slides
        If IsTimelineSlide(sld) Then AnyTimelineSlide = True: Exit Function
    Next sld
End Function

' A slide is a timeline if it carries imported-timeline data, is a paginated
' importer page, has a datebar, or contains at least one entry group.
Public Function IsTimelineSlide(ByVal sld As slide) As Boolean
    On Error Resume Next
    If Len(sld.Tags("TLDATA_N")) > 0 Then IsTimelineSlide = True: Exit Function
    If Len(sld.Tags("TLIMPORTER")) > 0 Then IsTimelineSlide = True: Exit Function   ' value is now the owning slide's ID
    On Error GoTo 0
    If Not FindDateBar(sld) Is Nothing Then IsTimelineSlide = True: Exit Function   ' FindDateBar is Public in TimelineCreator
    Dim shp As Shape
    For Each shp In sld.Shapes
        If HasEntryGroupInside(shp) Then IsTimelineSlide = True: Exit Function
    Next shp
End Function

Private Function HasEntryGroupInside(ByVal shp As Shape) As Boolean
    Dim it As Shape
    If IsEntryGroup(shp) Then HasEntryGroupInside = True: Exit Function
    If shp.Type = msoGroup Then
        For Each it In shp.GroupItems
            If HasEntryGroupInside(it) Then HasEntryGroupInside = True: Exit Function
        Next it
    End If
End Function

' Collect every entry GROUP under a shapes container (does not descend into an
' entry group once found - its children aren't separate entries).
Private Sub CollectEntryGroups(ByVal shapesColl As Object, ByVal acc As Collection)
    Dim shp As Shape
    For Each shp In shapesColl
        If IsDateBar(shp) Then
            ' skip: a datebar is never an entry and never contains one
        ElseIf IsEntryGroup(shp) Then
            acc.Add shp
        ElseIf shp.Type = msoGroup Then
            CollectEntryGroups shp.GroupItems, acc
        End If
    Next shp
End Sub

' An entry = a group carrying TLENTRY, having a leader-line child, or containing a tagged
' Date/Entry Box. Mirrors TimelineCreator's broadened IsEntryGroup so both agree about what
' counts as an entry (a renamed leader or hand-built entry stays visible).
Private Function IsEntryGroup(ByVal shp As Shape) As Boolean
    Dim it As Shape, tg As String
    If shp.Type <> msoGroup Then Exit Function
    If IsDateBar(shp) Then Exit Function              ' a bar is never an entry (it may hold rules/lines)
    tg = ""
    On Error Resume Next
    tg = shp.Tags("TLENTRY")
    On Error GoTo 0
    If tg = "1" Then IsEntryGroup = True: Exit Function
    For Each it In shp.GroupItems
        If it.name Like "LeadingLine*" Then IsEntryGroup = True: Exit Function
        If it.Type = msoLine Then
            If it.Width < 4 Then IsEntryGroup = True: Exit Function   ' vertical drop line only
        End If
    Next it
    ' depth-bounded so a wrapper that merely CONTAINS entries (boxes at depth >=3) isn't itself
    ' taken for an entry - CollectEntryGroups then descends into it to find the real entries.
    If Not FindTaggedChild(shp, "Date Box", 2) Is Nothing Then IsEntryGroup = True: Exit Function
    If Not FindTaggedChild(shp, "Entry Box", 2) Is Nothing Then IsEntryGroup = True
End Function

' Like FindTaggedDescendant but bounded to maxDepth levels below shp.
Private Function FindTaggedChild(ByVal shp As Shape, ByVal tagVal As String, ByVal maxDepth As Long) As Shape
    Dim it As Shape, found As Shape, gs As String
    gs = ""
    On Error Resume Next
    gs = shp.Tags("GroupStyle")
    On Error GoTo 0
    If gs = tagVal Then Set FindTaggedChild = shp: Exit Function
    If maxDepth <= 0 Then Exit Function
    If shp.Type = msoGroup Then
        For Each it In shp.GroupItems
            Set found = FindTaggedChild(it, tagVal, maxDepth - 1)
            If Not found Is Nothing Then Set FindTaggedChild = found: Exit Function
        Next it
    End If
End Function

' Find a descendant carrying a given GroupStyle tag ("Entry Box" / "Date Box").
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

' A titled entry box = first paragraph is fully bold and NOT a bullet.
' (Centering is incidental - long titles wrap and left-align - so it is ignored.)
Public Function IsTitledEntryBox(ByVal eb As Shape) As Boolean
    Dim tr As Object, p1 As Object, t As String
    On Error Resume Next
    If Not eb.HasTextFrame Then Exit Function
    If Not eb.TextFrame.HasText Then Exit Function
    Set tr = eb.TextFrame.TextRange
    If tr.Paragraphs.count < 1 Then Exit Function
    Set p1 = tr.Paragraphs(1)
    t = Trim$(Replace(p1.text, vbCr, ""))
    If t = "" Then Exit Function
    If p1.Font.Bold <> msoTrue Then Exit Function              ' msoTriStateMixed(-2)/msoFalse -> not a clean title
    If p1.ParagraphFormat.Bullet.Visible <> msoFalse Then Exit Function
    IsTitledEntryBox = True
    On Error GoTo 0
End Function

Public Function TitleOfEntryBox(ByVal eb As Shape) As String
    On Error Resume Next
    TitleOfEntryBox = Trim$(Replace(eb.TextFrame.TextRange.Paragraphs(1).text, vbCr, ""))
    On Error GoTo 0
End Function

' Body text (paragraphs 2..N) of a titled entry, flattened, for the UI list.
Private Function BodyTextOf(ByVal eb As Shape) As String
    Dim tr As Object, i As Long, s As String, ln As String
    On Error Resume Next
    Set tr = eb.TextFrame.TextRange
    For i = 2 To tr.Paragraphs.count
        ln = Trim$(Replace(tr.Paragraphs(i).text, vbCr, ""))
        If ln <> "" Then s = s & IIf(s = "", "", " / ") & ln
    Next i
    On Error GoTo 0
    BodyTextOf = s
End Function

Private Function SafeText(ByVal shp As Shape) As String
    On Error Resume Next
    SafeText = shp.TextFrame.TextRange.text
    On Error GoTo 0
End Function

Private Function FlattenText(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    FlattenText = Trim$(s)
End Function

' ============================================================================
' FORM HELPERS  (thin wrappers the userform calls; the form owns reg + save)
' ============================================================================
Public Sub ManualAssign(reg As Object, ByVal key As String, ByVal color As Long)
    TitleColors.SetTitleColor reg, key, color
    TitleColors.SetTitleManual reg, key, True
End Sub

Public Sub ReleaseToAuto(reg As Object, ByVal key As String)
    TitleColors.SetTitleManual reg, key, False    ' rejoins the auto pool; recolored on next Auto-Assign
End Sub

Public Sub ClearTitle(reg As Object, ByVal key As String)
    TitleColors.ClearTitleColor reg, key
End Sub
