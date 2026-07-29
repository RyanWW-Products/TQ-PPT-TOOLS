Attribute VB_Name = "ParseContentTools"
' ============================================================================
'  PARSE CONTENT  (Timeline Edits)
' ----------------------------------------------------------------------------
'  Rewrites the description text of every "Entry Box" on the active slide:
'    1. 1-2 displayed lines -> centered, 3+ -> left aligned.
'    2. Sentence case (each sentence starts capitalized; case of names/acronyms
'       is left alone).
'    3. The medical provider / facility is hoisted to a BOLD title line
'       ("Dr. Mayer", "ORMC"), the rest becomes the body.
'    4. One sentence -> no trailing period; two or more -> each ends in a period.
'    5. Multiple sentences -> broken into bullet lines.
'    6. Light grammar cleanup, but anything containing a quotation is left
'       untouched (only its alignment is adjusted).
'
'  ParseContent          - the one-click, fully-local VBA pass (best effort on
'                          provider extraction; awkward leftovers are left as-is).
'  ParseContentAIReword  - a smart clipboard button for true rewording:
'                            * clipboard is NOT a result  -> copies a Copilot /
'                              Claude prompt of every entry (keyed by a PCID tag)
'                            * clipboard IS a result       -> parses it and
'                              re-applies through the same renderer.
'
'  Both paths funnel through ApplyParsed so the on-slide result is identical.
' ============================================================================

Private Const SENTINEL As String = "TQ-REWORD"
Private Const BULLET_MARK As String = "- "     ' body-line prefix meaning "bulleted"

' ---------------------------------------------------------------------------
' Button 1: local, one-click parse.
' ---------------------------------------------------------------------------
Public Sub ParseContent(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveSlide()
    If sld Is Nothing Then
        MsgBox "Open a presentation and select a slide with timeline entries first.", vbExclamation, "Parse Content"
        Exit Sub
    End If

    Dim boxes As Collection
    Set boxes = New Collection
    CollectEntryBoxes sld.Shapes, boxes
    If boxes.count = 0 Then
        MsgBox "No timeline entry boxes were found on this slide.", vbExclamation, "Parse Content"
        Exit Sub
    End If

    Dim shp As Shape, raw As String, i As Long
    Dim changed As Long, skippedQuote As Long
    Dim title As String, body As Collection
    For i = 1 To boxes.count
        Set shp = boxes(i)
        raw = shp.TextFrame.TextRange.text
        If HasQuote(raw) Then
            AlignByLines shp                 ' leave the words alone, just fix alignment
            skippedQuote = skippedQuote + 1
        Else
            ProcessEntryText raw, title, body
            ApplyParsed shp, title, body
            changed = changed + 1
        End If
    Next i

    Dim msg As String
    msg = changed & " entr" & IIf(changed = 1, "y", "ies") & " parsed and reformatted."
    If skippedQuote > 0 Then _
        msg = msg & vbCrLf & skippedQuote & " left untouched (contains a quotation) - only realigned."
    MsgBox msg, vbInformation, "Parse Content"
    Exit Sub
Fail:
    MsgBox "Parse Content failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Parse Content"
End Sub

' ---------------------------------------------------------------------------
' Button 2: AI reword round-trip on one button.
'   - clipboard holds a TQ-REWORD result -> apply it.
'   - otherwise                          -> build + copy the prompt.
' The prompt itself contains the literal "TQ-REWORD" (in the format spec) and
' the phrase "OUTPUT FORMAT"; a real result never does, so we disambiguate on
' the absence of "OUTPUT FORMAT".
' ---------------------------------------------------------------------------
Public Sub ParseContentAIReword(control As IRibbonControl)
    On Error GoTo Fail
    Dim sld As slide
    Set sld = ActiveSlide()
    If sld Is Nothing Then
        MsgBox "Open a presentation and select a slide with timeline entries first.", vbExclamation, "Parse Content (AI Reword)"
        Exit Sub
    End If

    Dim clip As String
    clip = GetClipboardText()

    If InStr(1, clip, SENTINEL, vbTextCompare) > 0 And InStr(1, clip, "OUTPUT FORMAT", vbTextCompare) = 0 Then
        ' --- apply a reworded result ---
        Dim applied As Long, missing As Long
        ApplyRewordResult sld, clip, applied, missing
        Dim rm As String
        rm = applied & " entr" & IIf(applied = 1, "y", "ies") & " updated from the reworded result."
        If missing > 0 Then _
            rm = rm & vbCrLf & missing & " id(s) in the result had no matching entry on this slide (skipped)."
        MsgBox rm, vbInformation, "Parse Content (AI Reword)"
        Exit Sub
    End If

    ' --- build + copy the prompt ---
    Dim boxes As Collection
    Set boxes = New Collection
    CollectEntryBoxes sld.Shapes, boxes
    If boxes.count = 0 Then
        MsgBox "No timeline entry boxes were found on this slide.", vbExclamation, "Parse Content (AI Reword)"
        Exit Sub
    End If

    Dim prompt As String
    prompt = BuildRewordPrompt(boxes)      ' also stamps each box with a PCID tag
    SetClipboardText prompt

    MsgBox "Reword prompt for " & boxes.count & " entr" & IIf(boxes.count = 1, "y", "ies") & _
           " copied to the clipboard." & vbCrLf & vbCrLf & _
           "1. Paste it into Microsoft Copilot or Claude and run it." & vbCrLf & _
           "2. Copy the reply it gives you." & vbCrLf & _
           "3. Click Parse Content (AI Reword) again to apply it to this slide.", _
           vbInformation, "Parse Content (AI Reword)"
    Exit Sub
Fail:
    MsgBox "Parse Content (AI Reword) failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Parse Content (AI Reword)"
End Sub

' ===========================================================================
'  HEURISTIC ENGINE  (used by the one-click path)
' ===========================================================================
' Turns one raw entry string into a bold title + a body Collection of
' paragraph strings (a BULLET_MARK prefix means "render this as a bullet").
Private Sub ProcessEntryText(ByVal raw As String, ByRef outTitle As String, ByRef outBody As Collection)
    Dim flat As String
    flat = BasicGrammar(FlattenBreaks(raw))

    Dim provider As String, rest As String
    ExtractProvider flat, provider, rest
    outTitle = provider

    Set outBody = New Collection
    rest = Trim$(rest)

    Dim sents As Collection
    Set sents = SplitSentences(rest)

    Dim i As Long, s As String, multi As Boolean
    multi = (sents.count >= 2)
    For i = 1 To sents.count
        s = SentenceCaseOne(Trim$(sents(i)))
        If s = "" Then GoTo NextS
        If multi Then
            If Not EndsWithTerminal(s) Then s = s & "."
            outBody.Add BULLET_MARK & s          ' 2+ sentences -> bullets, each ends in a period
        Else
            s = StripTrailingPeriod(s)           ' single sentence -> no trailing period
            outBody.Add s
        End If
NextS:
    Next i
End Sub

' Best-effort provider/facility extraction. Conservative on purpose: it only
' drops pure prepositions/conjunctions around the provider, never verbs, so the
' one-click path never mangles the sentence badly (the AI path does real rewrites).
Private Sub ExtractProvider(ByVal s As String, ByRef provider As String, ByRef remainder As String)
    provider = "": remainder = s

    Dim toks() As String
    toks = Split(Trim$(s), " ")
    If UBound(toks) < 0 Then Exit Sub

    Dim t As Long, hon As Long, nameEnd As Long
    hon = -1

    ' --- doctor pattern: Dr / Dr. / Doctor followed by a capitalized name ---
    For t = 0 To UBound(toks)
        Dim bare As String
        bare = LCase$(StripEdgePunct(toks(t)))
        If bare = "dr" Or bare = "doctor" Then
            If t < UBound(toks) Then
                If IsCapWord(toks(t + 1)) Then
                    hon = t
                    nameEnd = t + 1
                    If (t + 2) <= UBound(toks) Then
                        If IsCapWord(toks(t + 2)) And Not IsConnector(toks(t + 2)) Then nameEnd = t + 2
                    End If
                    Exit For
                End If
            End If
        End If
    Next t

    If hon >= 0 Then
        Dim nm As String, k As Long
        For k = hon + 1 To nameEnd
            nm = nm & IIf(nm = "", "", " ") & TitleCaseWord(StripEdgePunct(toks(k)))
        Next k
        provider = "Dr. " & nm
        remainder = RebuildRemainder(toks, hon, nameEnd)
        Exit Sub
    End If

    ' --- facility acronym: an ALL-CAPS token immediately after at / to / from ---
    For t = 1 To UBound(toks)
        If IsLocConnector(toks(t - 1)) And IsAcronym(toks(t)) Then
            provider = StripEdgePunct(toks(t))
            remainder = RebuildRemainder(toks, t - 1, t)   ' also drops the at/to/from connector
            Exit Sub
        End If
    Next t
End Sub

' Join every token except indices firstDrop..lastDrop, then also drop a pure
' preposition/conjunction sitting immediately before firstDrop and a who/and
' sitting immediately after lastDrop, so the seam reads cleanly.
Private Function RebuildRemainder(ByRef toks() As String, ByVal firstDrop As Long, ByVal lastDrop As Long) As String
    Dim lo As Long, hi As Long
    lo = firstDrop: hi = lastDrop
    If lo - 1 >= 0 Then
        If IsPreDrop(toks(lo - 1)) Then lo = lo - 1
    End If
    If hi + 1 <= UBound(toks) Then
        If IsPostDrop(toks(hi + 1)) Then hi = hi + 1
    End If

    Dim res As String, t As Long
    For t = 0 To UBound(toks)
        If t < lo Or t > hi Then res = res & IIf(res = "", "", " ") & toks(t)
    Next t
    RebuildRemainder = CollapseSpaces(Trim$(res))
End Function

' ===========================================================================
'  RENDERER  (shared by both paths)
' ===========================================================================
' title: "" = none. body: Collection of paragraph strings; a BULLET_MARK prefix
' marks a bullet. Sets text, bolds the title, applies bullets, then aligns
' (1-2 lines centered / 3+ left).
Private Sub ApplyParsed(ByVal entryBox As Shape, ByVal title As String, ByVal body As Collection)
    Dim total As Long, hasTitle As Boolean
    hasTitle = (Len(title) > 0)
    total = IIf(hasTitle, 1, 0) + body.count
    If total = 0 Then Exit Sub

    Dim paras() As String, isBul() As Boolean
    ReDim paras(1 To total)
    ReDim isBul(1 To total)

    Dim idx As Long, v As Variant, ln As String
    idx = 0
    If hasTitle Then
        idx = idx + 1
        paras(idx) = title
        isBul(idx) = False
    End If
    For Each v In body
        idx = idx + 1
        ln = CStr(v)
        If Left$(ln, Len(BULLET_MARK)) = BULLET_MARK Then
            isBul(idx) = True
            paras(idx) = Mid$(ln, Len(BULLET_MARK) + 1)
        Else
            isBul(idx) = False
            paras(idx) = ln
        End If
    Next v

    Dim full As String, i As Long
    For i = 1 To total
        full = full & IIf(i = 1, "", vbCr) & paras(i)
    Next i

    Dim tr As Object
    Set tr = entryBox.TextFrame.TextRange
    tr.text = full

    Dim p As Object, cnt As Long
    cnt = tr.Paragraphs.count
    For i = 1 To cnt
        If i > total Then Exit For
        Set p = tr.Paragraphs(i)
        p.Font.Bold = IIf(hasTitle And i = 1, msoTrue, msoFalse)
        If isBul(i) Then
            p.ParagraphFormat.Bullet.Visible = msoTrue
            p.ParagraphFormat.Bullet.Type = ppBulletUnnumbered
            p.ParagraphFormat.Bullet.Character = 8226
        Else
            p.ParagraphFormat.Bullet.Visible = msoFalse
        End If
    Next i

    AlignByLines entryBox
End Sub

' 1-2 displayed (wrapped) lines -> center, 3+ -> left.
Private Sub AlignByLines(ByVal entryBox As Shape)
    Dim lc As Long
    On Error Resume Next
    lc = entryBox.TextFrame.TextRange.Lines.count
    On Error GoTo 0
    If lc <= 0 Then Exit Sub
    If lc <= 2 Then
        entryBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    Else
        entryBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    End If
End Sub

' ===========================================================================
'  AI REWORD  -  prompt build + result apply
' ===========================================================================
Private Function BuildRewordPrompt(ByVal boxes As Collection) As String
    Dim sb As String, i As Long, shp As Shape
    sb = "You are reformatting timeline entry text for a legal exhibit slide. " & _
         "Below are " & boxes.count & " entries, each marked with @<id> followed by its current text. " & _
         "Rewrite EACH entry per the rules, then return the result in the EXACT output format at the bottom." & vbCrLf & vbCrLf

    sb = sb & "RULES:" & vbCrLf
    sb = sb & "1. Sentence case: every sentence starts with a capital letter. Do NOT change the case of names, providers, or acronyms (e.g. ORMC, ER, CT, MRI)." & vbCrLf
    sb = sb & "2. Separate the medical provider or facility from the action and put it as a short TITLE (e.g. ""Dr. Mayer"", ""ORMC""). Reword the remaining sentence so it still reads correctly with the provider removed." & vbCrLf
    sb = sb & "   Example: ""follow-up visit with Dr. Mayer""  ->  T: Dr. Mayer / B: Follow-up visit" & vbCrLf
    sb = sb & "   Example: ""Px visits Dr. Johnson who prescribes pain killers""  ->  T: Dr. Johnson / B: Px visits and is prescribed pain killers" & vbCrLf
    sb = sb & "   If there is no clear provider, leave the title blank." & vbCrLf
    sb = sb & "3. If the entry is ONE sentence it must NOT end in a period. If it is TWO OR MORE sentences, each must end in a period." & vbCrLf
    sb = sb & "4. If the entry has multiple sentences, break them into separate bullet lines (prefix each body line with ""- ""). A single sentence gets no bullet." & vbCrLf
    sb = sb & "5. Fix only the most basic grammar/spelling. Do NOT alter anything inside quotation marks - reproduce quoted text VERBATIM." & vbCrLf
    sb = sb & "6. Keep the meaning identical. Do not invent facts, dates, names, or numbers." & vbCrLf & vbCrLf

    sb = sb & "OUTPUT FORMAT (return ONLY this, no commentary, no code fences):" & vbCrLf
    sb = sb & SENTINEL & vbCrLf
    sb = sb & "@<id>" & vbCrLf
    sb = sb & "T: <title, or leave blank>" & vbCrLf
    sb = sb & "B: <body line 1>" & vbCrLf
    sb = sb & "B: <body line 2>   (repeat B: per line; bullet lines start with ""- "")" & vbCrLf
    sb = sb & "...repeat the @<id> / T: / B: block for every entry..." & vbCrLf & vbCrLf

    sb = sb & "ENTRIES:" & vbCrLf
    For i = 1 To boxes.count
        Set shp = boxes(i)
        shp.Tags.Add "PCID", CStr(i)     ' Tags.Add overwrites an existing tag of the same name
        sb = sb & "@" & i & vbCrLf
        sb = sb & BasicGrammar(FlattenBreaks(shp.TextFrame.TextRange.text)) & vbCrLf
    Next i

    BuildRewordPrompt = sb
End Function

' Parse a TQ-REWORD block and apply each @id to the entry box tagged PCID=id.
Private Sub ApplyRewordResult(ByVal sld As slide, ByVal clip As String, ByRef applied As Long, ByRef missing As Long)
    applied = 0: missing = 0

    ' index PCID-tagged boxes on this slide
    Dim boxes As Collection
    Set boxes = New Collection
    CollectEntryBoxes sld.Shapes, boxes

    Dim lines() As String
    lines = Split(NormalizeNewlines(clip), vbLf)

    Dim i As Long, ln As String
    Dim curId As String, curTitle As String, curBody As Collection
    Dim haveBlock As Boolean
    haveBlock = False
    curId = "": curTitle = "": Set curBody = New Collection

    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        If Left$(LTrim$(ln), 1) = "@" Then
            If haveBlock Then FlushRewordBlock boxes, curId, curTitle, curBody, applied, missing
            curId = Trim$(Mid$(LTrim$(ln), 2))
            curTitle = "": Set curBody = New Collection
            haveBlock = True
        ElseIf UCase$(Left$(LTrim$(ln), 2)) = "T:" Then
            curTitle = Trim$(Mid$(LTrim$(ln), 3))
        ElseIf UCase$(Left$(LTrim$(ln), 2)) = "B:" Then
            Dim bl As String
            bl = Trim$(Mid$(LTrim$(ln), 3))
            If bl <> "" Then curBody.Add bl
        End If
    Next i
    If haveBlock Then FlushRewordBlock boxes, curId, curTitle, curBody, applied, missing
End Sub

Private Sub FlushRewordBlock(ByVal boxes As Collection, ByVal id As String, ByVal title As String, _
                             ByVal body As Collection, ByRef applied As Long, ByRef missing As Long)
    Dim shp As Shape
    Set shp = FindByPCID(boxes, id)
    If shp Is Nothing Then
        missing = missing + 1
        Exit Sub
    End If
    ApplyParsed shp, title, body
    applied = applied + 1
End Sub

Private Function FindByPCID(ByVal boxes As Collection, ByVal id As String) As Shape
    Dim shp As Shape, i As Long
    For i = 1 To boxes.count
        Set shp = boxes(i)
        If shp.Tags("PCID") = id Then
            Set FindByPCID = shp
            Exit Function
        End If
    Next i
End Function

' ===========================================================================
'  COLLECTION / TEXT HELPERS
' ===========================================================================
Private Sub CollectEntryBoxes(ByVal shapesColl As Object, ByVal acc As Collection)
    Dim shp As Shape
    For Each shp In shapesColl
        If shp.Type = msoGroup Then
            CollectEntryBoxes shp.GroupItems, acc
        ElseIf IsEntryBoxShape(shp) Then
            acc.Add shp
        End If
    Next shp
End Sub

Private Function IsEntryBoxShape(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If shp.Tags("GroupStyle") <> "Entry Box" Then Exit Function
    If Not shp.HasTextFrame Then Exit Function
    If Not shp.TextFrame.HasText Then Exit Function
    IsEntryBoxShape = True
End Function

Private Function ActiveSlide() As slide
    On Error Resume Next
    Set ActiveSlide = ActiveWindow.View.slide
    On Error GoTo 0
End Function

Private Function HasQuote(ByVal s As String) As Boolean
    HasQuote = (InStr(s, Chr$(34)) > 0) Or (InStr(s, Chr$(147)) > 0) Or _
               (InStr(s, Chr$(148)) > 0) Or (InStr(s, Chr$(145)) > 0) Or (InStr(s, Chr$(146)) > 0)
End Function

' Replace paragraph/line breaks and tabs with single spaces.
Private Function FlattenBreaks(ByVal s As String) As String
    s = Replace$(s, vbCrLf, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, Chr$(11), " ")     ' vertical tab (soft return)
    s = Replace$(s, vbTab, " ")
    FlattenBreaks = CollapseSpaces(Trim$(s))
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    CollapseSpaces = s
End Function

' Very light grammar cleanup - deliberately minimal so nothing gets harmed.
Private Function BasicGrammar(ByVal s As String) As String
    s = Replace$(s, " ,", ",")
    s = Replace$(s, " .", ".")
    s = Replace$(s, " ;", ";")
    s = Replace$(s, " :", ":")
    s = Replace$(s, " )", ")")
    s = Replace$(s, "( ", "(")
    ' standalone lowercase "i" -> "I"
    s = " " & s & " "
    s = Replace$(s, " i ", " I ")
    s = Replace$(s, " i'", " I'")
    s = Trim$(s)
    BasicGrammar = CollapseSpaces(s)
End Function

Private Function NormalizeNewlines(ByVal s As String) As String
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    NormalizeNewlines = s
End Function

' ---- sentence splitting -----------------------------------------------------
Private Function SplitSentences(ByVal s As String) As Collection
    Dim res As New Collection
    Dim i As Long, ch As String, buf As String
    buf = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        buf = buf & ch
        If ch = "." Or ch = "!" Or ch = "?" Then
            If IsSentenceEnd(s, i) Then
                If Trim$(buf) <> "" Then res.Add Trim$(buf)
                buf = ""
            End If
        End If
    Next i
    If Trim$(buf) <> "" Then res.Add Trim$(buf)
    Set SplitSentences = res
End Function

Private Function IsSentenceEnd(ByVal s As String, ByVal pos As Long) As Boolean
    Dim nxt As String
    If pos < Len(s) Then nxt = Mid$(s, pos + 1, 1) Else nxt = ""
    ' a "." glued to a letter/digit is a decimal, URL, or abbreviation - not an end
    If nxt <> "" And nxt <> " " Then
        If nxt Like "[0-9A-Za-z]" Then
            IsSentenceEnd = False
            Exit Function
        End If
    End If
    ' inspect the word ending just before the punctuation
    Dim j As Long, c As String, w As String
    j = pos - 1
    Do While j >= 1
        c = Mid$(s, j, 1)
        If c Like "[A-Za-z]" Then
            w = c & w
        Else
            Exit Do
        End If
        j = j - 1
    Loop
    If Len(w) = 1 Then          ' a lone initial ("J.")
        IsSentenceEnd = False
        Exit Function
    End If
    If IsAbbrev(w) Then
        IsSentenceEnd = False
        Exit Function
    End If
    IsSentenceEnd = True
End Function

Private Function IsAbbrev(ByVal w As String) As Boolean
    Dim lw As String
    lw = LCase$(w)
    Select Case lw
        Case "dr", "mr", "mrs", "ms", "prof", "st", "ave", "rd", "blvd", "mt", _
             "vs", "etc", "inc", "ltd", "co", "corp", "dept", "univ", "no", "fig", _
             "approx", "jr", "sr", "am", "pm", "min", "hr", "hrs", "oz", "ph", "esq"
            IsAbbrev = True
    End Select
End Function

' ---- casing / punctuation ---------------------------------------------------
Private Function SentenceCaseOne(ByVal s As String) As String
    Dim i As Long, ch As String
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            SentenceCaseOne = Left$(s, i - 1) & UCase$(ch) & Mid$(s, i + 1)
            Exit Function
        ElseIf ch Like "[0-9]" Then
            SentenceCaseOne = s        ' starts with a number - leave it
            Exit Function
        End If
    Next i
    SentenceCaseOne = s
End Function

Private Function EndsWithTerminal(ByVal s As String) As Boolean
    Dim c As String
    s = RTrim$(s)
    If Len(s) = 0 Then Exit Function
    c = Right$(s, 1)
    EndsWithTerminal = (c = "." Or c = "!" Or c = "?")
End Function

Private Function StripTrailingPeriod(ByVal s As String) As String
    s = RTrim$(s)
    Do While Right$(s, 1) = "."
        s = Left$(s, Len(s) - 1)
        s = RTrim$(s)
    Loop
    StripTrailingPeriod = s
End Function

' ---- token predicates -------------------------------------------------------
Private Function StripEdgePunct(ByVal w As String) As String
    Do While Len(w) > 0
        If Left$(w, 1) Like "[!-/:-@]" Then w = Mid$(w, 2) Else Exit Do
    Loop
    Do While Len(w) > 0
        If Right$(w, 1) Like "[!-/:-@]" Then w = Left$(w, Len(w) - 1) Else Exit Do
    Loop
    StripEdgePunct = w
End Function

Private Function IsCapWord(ByVal w As String) As Boolean
    w = StripEdgePunct(w)
    If Len(w) = 0 Then Exit Function
    IsCapWord = (Left$(w, 1) Like "[A-Z]")
End Function

' 2-6 letters, all uppercase (an acronym / facility code like ORMC, ER, ICU).
Private Function IsAcronym(ByVal w As String) As Boolean
    w = StripEdgePunct(w)
    If Len(w) < 2 Or Len(w) > 6 Then Exit Function
    Dim i As Long, c As String
    For i = 1 To Len(w)
        c = Mid$(w, i, 1)
        If Not (c Like "[A-Z]") Then Exit Function
    Next i
    IsAcronym = True
End Function

Private Function TitleCaseWord(ByVal w As String) As String
    If Len(w) = 0 Then Exit Function
    If w = UCase$(w) And Len(w) <= 6 Then     ' keep acronyms as-is
        TitleCaseWord = w
    Else
        TitleCaseWord = UCase$(Left$(w, 1)) & Mid$(w, 2)
    End If
End Function

Private Function IsConnector(ByVal w As String) As Boolean
    Select Case LCase$(StripEdgePunct(w))
        Case "with", "by", "who", "and", "at", "to", "from", "the", "a", "an", _
             "seen", "sees", "see", "saw", "visits", "visit", "visited", "for", "of"
            IsConnector = True
    End Select
End Function

' preposition/conjunction ok to drop immediately BEFORE the provider
Private Function IsPreDrop(ByVal w As String) As Boolean
    Select Case LCase$(StripEdgePunct(w))
        Case "with", "by", "and", "at", "to", "from"
            IsPreDrop = True
    End Select
End Function

' word ok to drop immediately AFTER the provider
Private Function IsPostDrop(ByVal w As String) As Boolean
    Select Case LCase$(StripEdgePunct(w))
        Case "who", "and"
            IsPostDrop = True
    End Select
End Function

Private Function IsLocConnector(ByVal w As String) As Boolean
    Select Case LCase$(StripEdgePunct(w))
        Case "at", "to", "from"
            IsLocConnector = True
    End Select
End Function

' ===========================================================================
'  CLIPBOARD  (late-bound MSForms.DataObject - no form needed)
' ===========================================================================
Private Function GetClipboardText() As String
    On Error Resume Next
    Dim dobj As Object
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dobj.GetFromClipboard
    GetClipboardText = dobj.GetText
    On Error GoTo 0
End Function

Private Sub SetClipboardText(ByVal s As String)
    On Error Resume Next
    Dim dobj As Object
    Set dobj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dobj.SetText s
    dobj.PutInClipboard
    On Error GoTo 0
End Sub
