Attribute VB_Name = "EZHighlightsTests"
Option Explicit
Private report As Integer, checks As Long, failures As Long

Public Sub RunAll()
    Dim deck As Presentation, i As Long
    Set deck = Application.Presentations("EZHighlightsRegression.pptm")
    report = FreeFile
    Open deck.Path & "\report.txt" For Output As #report
    For i = 1 To 9
        RunCase deck, i
    Next
    RunGroupCase deck, False
    RunGroupCase deck, True
    RunMultiChildCase deck
    RunReplayCase deck
    RunAddReplayCase deck
    RunGroupedReplayButtonCase deck

    Print #report, "RESULT | checks=" & checks & " failures=" & failures
    Close #report
    deck.Save
End Sub

Private Sub RunAddReplayCase(ByVal deck As Presentation)
    Dim sld As Slide, source As Shape, highlight As Shape, neighbor As Shape, faded As Shape
    Dim effect As Effect, window As DocumentWindow, count As Long, originalId As Long
    On Error GoTo failed
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    sld.Name = "Add Replay on fresh highlight"
    Set neighbor = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 35, 55, 650, 55)
    neighbor.TextFrame.TextRange.Text = "BLACK text under YELLOW native ink"
    neighbor.TextFrame.TextRange.Font.Size = 24
    neighbor.TextFrame.TextRange.Font.Color.RGB = vbBlack
    If deck.Windows.Count = 0 Then Set window = deck.NewWindow Else Set window = deck.Windows(1)
    window.Activate: window.ViewType = ppViewNormal: window.View.GotoSlide sld.SlideIndex
    Set source = sld.Shapes.AddShape(msoShapeRectangle, 35, 60, 500, 45)
    EZHighlightsDrawn sld, source
    Set highlight = sld.Shapes(2)
    originalId = highlight.Id
    Check sld.TimeLine.MainSequence.Count = 0, "newly drawn highlight needs no preliminary animation"
    highlight.Select
    AddHighlightReplayClick Nothing
    Check sld.TimeLine.MainSequence.Count = 1, "Add Replay creates one entrance directly on a new highlight"
    Set effect = sld.TimeLine.MainSequence(1)
    Check effect.Shape.Id = originalId And highlight.Type = msoGroup, "Add Replay keeps the intact highlight group"
    Check effect.Exit = msoFalse And effect.Behaviors(2).Type = msoAnimTypeProperty, "Add Replay uses a native replay entrance"
    Check effect.Behaviors(2).PropertyEffect.Points(1).Value = 0 And effect.Behaviors(2).PropertyEffect.Points(2).Value = 1, "Replay draws from zero to full ink"
    Check Abs(effect.Timing.Duration - 2) < 0.01 And effect.Timing.TriggerDelayTime = 0, "new Replay defaults to two seconds without delay"
    Check effect.Timing.TriggerType = msoAnimTriggerOnPageClick, "new Replay defaults to On Click"
    Check Application.CommandBars.GetPressedMso("AnimationCustom"), "Add Replay opens the Animation Pane"
    Check sld.Shapes.Count = 2, "Add Replay removes its staging shape"
    effect.Timing.Duration = 1.2: effect.Timing.TriggerDelayTime = 0.3
    highlight.Select: neighbor.Select msoFalse
    AddHighlightReplayClick Nothing
    Check sld.TimeLine.MainSequence.Count = 1, "repeated Add Replay ignores ordinary neighbors and avoids duplicates"
    Set effect = sld.TimeLine.MainSequence(1)
    Check Abs(effect.Timing.Duration - 1.2) < 0.01 And Abs(effect.Timing.TriggerDelayTime - 0.3) < 0.01, "repeated Add Replay retains edited timing"
    effect.Timing.Duration = 2: effect.Timing.TriggerDelayTime = 0

    Set source = sld.Shapes.AddShape(msoShapeRectangle, 80, 180, 250, 30)
    Set faded = EZHighlightsConvert(sld, source)
    Set effect = sld.TimeLine.MainSequence.AddEffect(faded, msoAnimEffectFade)
    effect.Timing.Duration = 1.1: effect.Timing.TriggerDelayTime = 0.25
    window.Activate: window.View.GotoSlide sld.SlideIndex
    faded.Select
    AddHighlightReplayClick Nothing
    Check sld.TimeLine.MainSequence.Count = 4, "Add Replay repairs existing Fade without adding a second entrance"
    For count = 2 To 4
        Set effect = sld.TimeLine.MainSequence(count)
        Check Abs(effect.Timing.Duration - 1.1) < 0.01 And Abs(effect.Timing.TriggerDelayTime - 0.25) < 0.01, "Add Replay retains preexisting Fade timing"
    Next
    Check sld.Shapes.Count = 3 And highlight.Id = originalId, "Fade repair leaves the first highlight and shapes intact"
    Exit Sub
failed:
    Check False, "Add Replay button runtime " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunGroupedReplayButtonCase(ByVal deck As Presentation)
    Dim sld As Slide, a As Shape, b As Shape, neighbor As Shape, inner As Shape, root As Shape, part As Shape
    Dim inkA As Shape, inkB As Shape, editableB As Shape, window As DocumentWindow, effect As Effect
    Dim originalId As Long, leafCount As Long, left As Single, top As Single, width As Single, height As Single
    On Error GoTo failed
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    Set a = sld.Shapes.AddShape(msoShapeRectangle, 80, 80, 150, 40)
    Set b = sld.Shapes.AddShape(msoShapeOval, 260, 80, 140, 40)
    Set inner = sld.Shapes.Range(Array(a.Name, b.Name)).Group
    inner.Rotation = 18
    Set neighbor = sld.Shapes.AddShape(msoShapeRectangle, 450, 200, 70, 40)
    neighbor.Name = "unanimated neighbor"
    Set root = sld.Shapes.Range(Array(inner.Name, neighbor.Name)).Group
    root.Rotation = 12: root.Flip msoFlipHorizontal
    If deck.Windows.Count = 0 Then Set window = deck.NewWindow Else Set window = deck.Windows(1)
    window.Activate: window.ViewType = ppViewNormal: window.View.GotoSlide sld.SlideIndex
    Set a = root.GroupItems(1): Set b = root.GroupItems(2)
    a.Select: b.Select msoFalse
    EZHighlightsClick Nothing
    Set root = sld.Shapes(1)
    For Each part In root.GroupItems
        If part.Tags("EZHighlight") = "1" Then
            If inkA Is Nothing Then Set inkA = part Else Set inkB = part
        End If
        If part.Tags("EZHighlightMember") = "1" Then Set editableB = part
    Next
    Check Not inkA Is Nothing And Not inkB Is Nothing, "nested group contains two highlights ready for Replay"
    originalId = root.Id: leafCount = root.GroupItems.Count
    left = root.Left: top = root.Top: width = root.Width: height = root.Height
    inkA.Select
    AddHighlightReplayClick Nothing
    Check sld.TimeLine.MainSequence.Count = 1 And sld.TimeLine.MainSequence(1).Shape.Id = inkA.Id, "Replay targets only the selected grouped ink"
    Set effect = sld.TimeLine.MainSequence(1)
    effect.Timing.Duration = 1.3: effect.Timing.TriggerDelayTime = 0.2
    window.Activate: window.View.GotoSlide sld.SlideIndex
    editableB.Select
    AddHighlightReplayClick Nothing
    Check sld.TimeLine.MainSequence.Count = 2 And sld.TimeLine.MainSequence(2).Shape.Id = inkB.Id, "selecting a retained member targets only its matching grouped ink"
    window.Activate: window.View.GotoSlide sld.SlideIndex
    root.Select
    AddHighlightReplayClick Nothing
    Check sld.TimeLine.MainSequence.Count = 2, "whole-group Replay skips already animated highlights and ordinary neighbors"
    Check Abs(sld.TimeLine.MainSequence(1).Timing.Duration - 1.3) < 0.01 And Abs(sld.TimeLine.MainSequence(1).Timing.TriggerDelayTime - 0.2) < 0.01, "group Replay preserves existing order and timing"
    Check root.Id = originalId And root.GroupItems.Count = leafCount And sld.Shapes.Count = 1, "group Replay retains hierarchy without staging shapes"
    Check Abs(root.Left - left) < 0.01 And Abs(root.Top - top) < 0.01 And Abs(root.Width - width) < 0.01 And Abs(root.Height - height) < 0.01, "group Replay retains transformed geometry"
    Exit Sub
failed:
    Check False, "group Add Replay button runtime " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunReplayCase(ByVal deck As Presentation)
    Dim sld As Slide, source As Shape, highlight As Shape, neighbor As Shape, trigger As Shape
    Dim effect As Effect, sequence As Sequence, part As Shape, beforeCount As Long, inkCount As Long, textCount As Long
    Dim afterCount As Long, savedId As Long, duration As Single, delay As Single
    On Error GoTo failed
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    Set neighbor = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 35, 60, 600, 55)
    neighbor.TextFrame.TextRange.Text = "BLACK text under YELLOW native ink"
    neighbor.TextFrame.TextRange.Font.Size = 24
    Set source = sld.Shapes.AddShape(msoShapeRectangle, 35, 60, 500, 45)
    source.TextFrame.TextRange.Text = "Editable caption"
    Set effect = sld.TimeLine.MainSequence.AddEffect(source, msoAnimEffectFade)
    effect.Timing.Duration = 2.4: effect.Timing.TriggerDelayTime = 0.35
    Set effect = sld.TimeLine.MainSequence.AddEffect(source, msoAnimEffectFade, , msoAnimTriggerAfterPrevious)
    effect.Exit = msoTrue
    effect.Timing.Duration = 1.6: effect.Timing.TriggerDelayTime = 0.2
    Set highlight = EZHighlightsConvert(sld, source)
    savedId = highlight.Id: beforeCount = sld.Shapes.Count
    EZHighlightReplay.RepairHighlightAnimations sld, highlight
    Check highlight.Id = savedId And sld.Shapes.Count = beforeCount, "Replay repair retains the group and removes staging shapes"
    For Each effect In sld.TimeLine.MainSequence
        Set part = effect.Shape
        If effect.Exit Then duration = 1.6: delay = 0.2 Else duration = 2.4: delay = 0.35
        Check Abs(effect.Timing.Duration - duration) < 0.01, "Replay retains effect duration"
        Check Abs(effect.Timing.TriggerDelayTime - delay) < 0.01, "Replay retains effect delay"
        Check effect.Timing.RepeatCount <= 1, "Replay does not add animation repeats"
        If part.Type = msoInk Or part.Type = msoInkComment Then
            inkCount = inkCount + 1
            Check effect.Behaviors.Count = 2, "ink uses native draw-progress animation"
        Else
            textCount = textCount + 1
            Check effect.EffectType = msoAnimEffectFade, "editable caption retains Fade"
        End If
    Next
    Check inkCount = 4 And textCount = 2, "both ink strokes and editable caption have entry and exit effects"
    Check sld.TimeLine.MainSequence(1).Timing.TriggerType = msoAnimTriggerOnPageClick, "Replay retains entrance click"
    Check sld.TimeLine.MainSequence(4).Timing.TriggerType = msoAnimTriggerAfterPrevious, "Rewind retains after-previous start"
    afterCount = sld.TimeLine.MainSequence.Count
    EZHighlightReplay.RepairHighlightAnimations sld, highlight
    Check sld.TimeLine.MainSequence.Count = afterCount, "repeating Replay repair adds no animations"
    Set trigger = sld.Shapes.AddShape(msoShapeOval, 590, 250, 50, 50)
    Set sequence = sld.TimeLine.InteractiveSequences.Add
    Set effect = sequence.AddEffect(highlight, msoAnimEffectFade, , msoAnimTriggerOnShapeClick)
    effect.Timing.TriggerShape = trigger
    effect.Timing.Duration = 3.2: effect.Timing.TriggerDelayTime = 0.6
    EZHighlightReplay.RepairHighlightAnimations sld, highlight
    Check sld.TimeLine.InteractiveSequences.Count = 1, "Replay keeps one interactive sequence"
    Set sequence = sld.TimeLine.InteractiveSequences(1)
    Check sequence.Count = 3, "interactive Replay and caption run together"
    For Each effect In sequence
        Check effect.Timing.TriggerShape.Id = trigger.Id, "interactive Replay retains click target"
        Check Abs(effect.Timing.Duration - 3.2) < 0.01 And Abs(effect.Timing.TriggerDelayTime - 0.6) < 0.01, "interactive Replay retains timing"
    Next
    Check sld.Shapes.Count = beforeCount + 1, "interactive repair removes staging shape"
    Exit Sub
failed:
    Check False, "Replay case runtime " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunMultiChildCase(ByVal deck As Presentation)
    Dim sld As Slide, a As Shape, b As Shape, c As Shape, root As Shape, child As Shape
    Dim window As DocumentWindow, effect As Effect, ink As Shape, editable As Shape
    Dim originalId As Long, highlights As Long, originalShapeCount As Long
    On Error GoTo failed
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    Set a = sld.Shapes.AddShape(msoShapeRectangle, 80, 80, 120, 35)
    Set b = sld.Shapes.AddShape(msoShapeOval, 240, 80, 100, 35)
    Set c = sld.Shapes.AddShape(msoShapeRectangle, 400, 80, 70, 35)
    c.Name = "untouched": c.Fill.ForeColor.RGB = vbGreen
    Set root = sld.Shapes.Range(Array(a.Name, b.Name, c.Name)).Group
    root.Name = "multiple selected children"
    Set a = root.GroupItems(1): Set b = root.GroupItems(2)
    a.Name = "same child name": b.Name = "same child name"
    a.TextFrame.TextRange.Text = "Keep animated text"
    Set effect = sld.TimeLine.MainSequence.AddEffect(a, msoAnimEffectWipe)
    Set effect = sld.TimeLine.MainSequence.AddEffect(a, msoAnimEffectFade, , msoAnimTriggerAfterPrevious)
    effect.Exit = msoTrue
    Set effect = sld.TimeLine.MainSequence.AddEffect(b, msoAnimEffectFade, , msoAnimTriggerAfterPrevious)
    If deck.Windows.Count = 0 Then Set window = deck.NewWindow Else Set window = deck.Windows(1)
    window.Activate: window.ViewType = ppViewNormal: window.View.GotoSlide sld.SlideIndex
    a.Select: b.Select msoFalse
    Check window.Selection.HasChildShapeRange And window.Selection.ChildShapeRange.Count = 2, "multiple child selection"
    EZHighlightsClick Nothing
    Set root = sld.Shapes(1)
    For Each child In root.GroupItems
        If child.Tags("EZHighlight") = "1" Then highlights = highlights + 1: Set ink = child
        If child.Tags("EZHighlightMember") = "1" Then Set editable = child
    Next
    Check highlights = 2, "two selected children converted despite duplicate names"
    Check root.GroupItems("untouched").Fill.ForeColor.RGB = vbGreen, "multi-selection leaves unselected neighbor unchanged"
    Check sld.Shapes.Count = 1, "multi-selection leaves one containing group"
    Check sld.TimeLine.MainSequence.Count = 5, "multiple effects retain synchronized editable text"
    Check sld.TimeLine.MainSequence(1).EffectType = msoAnimEffectWipe And sld.TimeLine.MainSequence(3).Behaviors(1).Type = msoAnimTypeProperty, "multiple child animation order with safe ink exit"
    Check sld.TimeLine.MainSequence(3).Exit And sld.TimeLine.MainSequence(4).Exit, "ink and editable text retain exit effects"
    originalId = root.Id: originalShapeCount = root.GroupItems.Count
    ink.Select
    EZHighlightsClick Nothing
    Check sld.Shapes(1).Id = originalId And sld.Shapes(1).GroupItems.Count = originalShapeCount, "repeat click on grouped ink is unchanged"
    Set root = sld.Shapes(1)
    For Each child In root.GroupItems
        If child.Tags("EZHighlightMember") = "1" Then Set editable = child
    Next
    window.Activate: window.View.GotoSlide sld.SlideIndex
    editable.Select
    EZHighlightsClick Nothing
    Check sld.Shapes(1).Id = originalId And sld.Shapes(1).GroupItems.Count = originalShapeCount, "repeat click on retained editable child is unchanged"
    Exit Sub
failed:
    Check False, "multiple child case runtime " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunGroupCase(ByVal deck As Presentation, ByVal nested As Boolean)
    Dim sld As Slide, a As Shape, b As Shape, inner As Shape, root As Shape, result As Shape
    Dim c As Shape, effect As Effect, sequence As Sequence, window As DocumentWindow, ink As Shape, child As Shape
    Dim rootId As Long, tag As String, originalLeft As Single, originalTop As Single
    Dim originalWidth As Single, originalHeight As Single, originalRotation As Single
    Dim bx As Single, by As Single, bw As Single, bh As Single, br As Single
    On Error GoTo failed
    tag = "group nested=" & nested
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    Set a = sld.Shapes.AddShape(msoShapeChevron, 80, 100, 190, 60)
    a.Name = "selected child": a.Tags.Add "ChildTag", "keep"
    a.TextFrame.TextRange.Text = "Selected"
    a.Fill.ForeColor.RGB = vbRed
    Set b = sld.Shapes.AddShape(msoShapeOval, 300, 200, 80, 40)
    b.Name = "neighbor": b.Fill.ForeColor.RGB = vbBlue
    Set inner = sld.Shapes.Range(Array(a.Name, b.Name)).Group
    inner.Name = "inner group"
    Set root = inner
    If nested Then
        inner.Rotation = 23
        inner.Flip msoFlipHorizontal
        Set c = sld.Shapes.AddShape(msoShapeRectangle, 400, 300, 50, 40)
        c.Name = "outer neighbor": c.Fill.ForeColor.RGB = vbGreen
        Set root = sld.Shapes.Range(Array(inner.Name, c.Name)).Group
        root.Rotation = 31
        root.Flip msoFlipVertical
        root.LockAspectRatio = msoFalse
        root.Width = root.Width * 1.3
        root.Height = root.Height * 0.8
    End If
    root.Name = "outer group": root.Tags.Add "GroupTag", "keep"
    root.AlternativeText = "Group description"
    rootId = root.Id
    originalLeft = root.Left: originalTop = root.Top
    originalWidth = root.Width: originalHeight = root.Height: originalRotation = root.Rotation
    Set a = root.GroupItems("selected child"): Set b = root.GroupItems("neighbor")
    bx = b.Left: by = b.Top: bw = b.Width: bh = b.Height: br = b.Rotation
    Set effect = sld.TimeLine.MainSequence.AddEffect(root, msoAnimEffectAppear)
    Set effect = sld.TimeLine.MainSequence.AddEffect(a, msoAnimEffectWipe, , msoAnimTriggerAfterPrevious)
    effect.Timing.Duration = 1.2
    Set effect = sld.TimeLine.MainSequence.AddEffect(b, msoAnimEffectFade, , msoAnimTriggerWithPrevious)
    Set sequence = sld.TimeLine.InteractiveSequences.Add
    Set effect = sequence.AddEffect(b, msoAnimEffectAppear, , msoAnimTriggerOnShapeClick)
    effect.Exit = msoTrue
    effect.Timing.TriggerShape = a
    For Each effect In sld.TimeLine.MainSequence
        Print #report, "BEFORE " & tag & " | " & effect.Shape.Name & " | " & effect.EffectType & " | " & effect.Timing.Duration
    Next
    If deck.Windows.Count = 0 Then Set window = deck.NewWindow Else Set window = deck.Windows(1)
    window.Activate
    window.ViewType = ppViewNormal
    window.View.GotoSlide sld.SlideIndex
    a.Select
    Check window.Selection.HasChildShapeRange, tag & " child selection reached ribbon handler"
    EZHighlightsClick Nothing
    Set result = sld.Shapes(1)
    For Each effect In sld.TimeLine.MainSequence
        Print #report, "AFTER " & tag & " | " & effect.Shape.Name & " | " & effect.EffectType & " | " & effect.Timing.Duration
    Next
    Check result.Id <> rootId, tag & " group replacement completed"
    Check sld.Shapes.Count = 1, tag & " no staging shapes left"
    Check result.Type = msoGroup, tag & " outer group retained"
    Check result.Name = "outer group" And result.Tags("GroupTag") = "keep", tag & " group name and tags"
    Check result.AlternativeText = "Group description", tag & " group alt text"
    Check Abs(result.Left - originalLeft) < 0.1 And Abs(result.Top - originalTop) < 0.1, tag & " group position"
    Check Abs(result.Width - originalWidth) < 0.1 And Abs(result.Height - originalHeight) < 0.1, tag & " group size"
    Check Abs(result.Rotation - originalRotation) < 0.1, tag & " group rotation"
    If nested Then
        Check result.VerticalFlip = msoTrue, tag & " outer flip"
        Check result.GroupItems("outer neighbor").Fill.ForeColor.RGB = vbGreen, tag & " outer sibling preserved"
    End If
    For Each child In result.GroupItems
        If child.Tags("EZHighlight") = "1" Then Set ink = child
        If child.Tags("ChildTag") = "keep" Then Set a = child
    Next
    Set b = result.GroupItems("neighbor")
    Check Not ink Is Nothing, tag & " selected child has native highlight"
    Check a.Tags("ChildTag") = "keep" And a.Fill.Visible = msoFalse, tag & " editable child metadata"
    Check b.Type = msoAutoShape And b.Fill.ForeColor.RGB = vbBlue, tag & " neighbor unchanged"
    Check b.Name = "neighbor", tag & " neighbor name"
    Check Abs(b.Left - bx) < 0.1 And Abs(b.Top - by) < 0.1, tag & " neighbor position preserved"
    Check Abs(b.Width - bw) < 0.1 And Abs(b.Height - bh) < 0.1 And Abs(b.Rotation - br) < 0.1, tag & " neighbor transform preserved"
    Check sld.TimeLine.MainSequence.Count = 4, tag & " animations retained with synchronized text"
    Check sld.TimeLine.MainSequence(1).Shape.Id = result.Id, tag & " outer animation retargeted"
    Check sld.TimeLine.MainSequence(2).Shape.Id = ink.Id, tag & " child animation retargeted"
    Check Abs(sld.TimeLine.MainSequence(2).Timing.Duration - 1.2) < 0.001, tag & " child animation timing"
    Check sld.TimeLine.MainSequence(3).Shape.Id = a.Id And sld.TimeLine.MainSequence(3).Timing.TriggerType = msoAnimTriggerWithPrevious, tag & " editable text animates with ink"
    Check sld.TimeLine.MainSequence(4).Shape.Id = b.Id, tag & " sibling animation retargeted"
    Check sld.TimeLine.InteractiveSequences(1)(1).Timing.TriggerShape.Id = ink.Id, tag & " child click trigger retained"
    sld.Export deck.Path & "\" & IIf(nested, "group-nested", "group-simple") & ".png", "PNG", 1440, 810
    Exit Sub
failed:
    Check False, tag & " runtime " & Err.Number & ": " & Err.Description
    On Error Resume Next
End Sub

Public Sub PrepareMouseDraw()
    Dim deck As Presentation, window As DocumentWindow, sld As Slide, f As Integer
    On Error GoTo failed
    Set deck = Application.Presentations("EZHighlightsRegression.pptm")
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    Set window = deck.NewWindow
    Application.Visible = msoTrue
    window.Activate
    window.ViewType = ppViewNormal
    window.View.GotoSlide sld.SlideIndex
    window.Selection.Unselect
    DoEvents
    Exit Sub
failed:
    TestMessage Err.Description
End Sub

Public Sub ArmMouseDraw()
    ActiveWindow.Selection.Unselect
    EZHighlightsClick Nothing
End Sub

Public Sub AddReplayToMouseHighlight()
    AddHighlightReplayClick Nothing
End Sub

Public Sub PlainRectangleMode()
    Application.CommandBars.ExecuteMso "ShapeRectangle"
End Sub

Public Sub TestMessage(ByVal message As String, Optional ByVal style As Long, Optional ByVal title As String)
    Dim f As Integer
    f = FreeFile
    Open Application.Presentations("EZHighlightsRegression.pptm").Path & "\ui-errors.txt" For Append As #f
    Print #f, message
    Close #f
End Sub

Private Sub Check(ByVal okay As Boolean, ByVal message As String)
    checks = checks + 1
    If okay Then
        Print #report, "PASS | " & message
    Else
        failures = failures + 1
        Print #report, "FAIL | " & message
    End If
End Sub

Private Sub RunCase(ByVal deck As Presentation, ByVal test As Long)
    Dim sld As Slide, source As Shape, converted As Shape, ink As Shape, editable As Shape
    Dim effect As Effect, other As Shape, sequence As Sequence, target As Shape
    Dim shapeType As Long, i As Long, count As Long, again As Shape
    Dim originalId As Long, originalName As String, child As Shape, builder As FreeformBuilder
    On Error GoTo failed
    Set sld = deck.Slides.Add(deck.Slides.Count + 1, ppLayoutBlank)
    sld.Name = "EZ case " & test
    sld.FollowMasterBackground = msoFalse
    sld.Background.Fill.ForeColor.RGB = vbWhite
    Set other = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 55, 700, 55)
    other.TextFrame.TextRange.Text = "BLACK text under YELLOW native ink"
    other.TextFrame.TextRange.Font.Size = 24
    other.TextFrame.TextRange.Font.Color.RGB = vbBlack
    Select Case test
        Case 1, 6, 7, 8, 9: shapeType = msoShapeRectangle
        Case 2: shapeType = msoShapeOval
        Case 3: shapeType = msoShapeDonut
        Case 4: shapeType = msoShapeChevron
        Case 5: shapeType = msoShapeRoundedRectangle
    End Select
    Set source = sld.Shapes.AddShape(shapeType, 35, 60, 500, 45)
    source.Name = "Original " & test
    source.Tags.Add "KeepMe", "original tag"
    source.AlternativeText = "Accessible description"
    source.Title = "Accessible title"
    source.Fill.ForeColor.RGB = vbRed
    source.Line.Weight = 3
    If test = 4 Then source.Rotation = 27
    If test = 5 Then source.Adjustments(1) = 0.35
    If test = 6 Then
        source.TextFrame.TextRange.Text = "Keep editable text"
        source.TextFrame.TextRange.Font.Italic = msoTrue
        source.TextFrame.TextRange.Font.Size = 18
    End If
    If test = 7 Then
        source.Delete
        Set builder = sld.Shapes.BuildFreeform(msoEditingCorner, 35, 60)
        builder.AddNodes msoSegmentLine, msoEditingCorner, 535, 60
        builder.AddNodes msoSegmentLine, msoEditingCorner, 435, 105
        builder.AddNodes msoSegmentLine, msoEditingCorner, 35, 90
        builder.AddNodes msoSegmentLine, msoEditingCorner, 35, 60
        Set source = builder.ConvertToShape
        source.Tags.Add "KeepMe", "original tag"
        source.AlternativeText = "Accessible description"
        source.Title = "Accessible title"
    End If
    If test = 8 Then
        source.Width = 0.5: source.Height = 0.5
    End If
    If test = 9 Then
        source.Flip msoFlipHorizontal
        source.Visible = msoFalse
    End If
    originalId = source.Id: originalName = source.Name
    Set effect = sld.TimeLine.MainSequence.AddEffect(source, msoAnimEffectWipe, , msoAnimTriggerAfterPrevious)
    effect.EffectParameters.Direction = msoAnimDirectionLeft
    effect.Timing.Duration = 1.75
    effect.Timing.TriggerDelayTime = 0.35
    Set target = sld.Shapes.AddShape(msoShapeOval, 650, 150, 50, 50)
    Set sequence = sld.TimeLine.InteractiveSequences.Add
    Set effect = sequence.AddEffect(target, msoAnimEffectAppear, , msoAnimTriggerOnShapeClick)
    effect.Timing.TriggerShape = source
    Set effect = sequence.AddEffect(source, msoAnimEffectFade, , msoAnimTriggerWithPrevious)
    effect.Exit = msoTrue
    Set converted = EZHighlightsConvert(sld, source)
    Check converted.Type = msoGroup, test & " editable highlight group"
    Check converted.Name = originalName, test & " name preserved"
    Check converted.Tags("KeepMe") = "original tag", test & " tags preserved"
    Check converted.AlternativeText = "Accessible description", test & " alt text preserved"
    Check converted.Title = "Accessible title", test & " title preserved"
    Check converted.ZOrderPosition = 2, test & " stacking position preserved"
    Check sld.Shapes.Count = 3, test & " no stray staging shapes"
    For Each child In converted.GroupItems
        If child.Type = msoInk Or child.Type = msoInkComment Then
            Set ink = child
        Else
            Set editable = child
        End If
    Next
    Check Not ink Is Nothing, test & " native ink embedded"
    Check Not editable Is Nothing, test & " editable geometry retained"
    Check editable.Fill.Visible = msoFalse And editable.Line.Visible = msoFalse, test & " original fill and outline hidden"
    ' Native ink bounds describe its centerline; brush thickness extends beyond
    ' that line. The retained editable shape controls the group's full bounds.
    Check Abs(converted.Width - editable.Width) < 0.1 Or test = 4, test & " full highlight width retained"
    Check Abs(ink.Rotation - editable.Rotation) < 0.1, test & " rotation preserved"
    If test = 6 Then
        Check editable.TextFrame.TextRange.Text = "Keep editable text", "text content retained"
        Check editable.TextFrame.TextRange.Font.Italic = msoTrue, "text formatting retained"
    End If
    If test = 5 Then Check Abs(editable.Adjustments(1) - 0.35) < 0.001, "adjusted geometry retained"
    If test = 9 Then Check converted.Visible = msoFalse, "visibility retained"
    Check sld.TimeLine.MainSequence.Count = 1, test & " no duplicate animations"
    Set effect = sld.TimeLine.MainSequence(1)
    Check effect.Shape.Id = converted.Id, test & " animation retargeted"
    Check effect.EffectType = msoAnimEffectWipe, test & " effect type preserved"
    Check Abs(effect.Timing.Duration - 1.75) < 0.001, test & " duration preserved"
    Check Abs(effect.Timing.TriggerDelayTime - 0.35) < 0.001, test & " delay preserved"
    Check effect.EffectParameters.Direction = msoAnimDirectionLeft, test & " direction preserved"
    Check effect.Timing.TriggerType = msoAnimTriggerAfterPrevious, test & " sequence timing preserved"
    Check sequence.Count = 2, test & " interactive animation count preserved"
    Check sequence(1).Timing.TriggerShape.Id = converted.Id, test & " click trigger retargeted"
    Check sequence(2).Shape.Id = converted.Id And sequence(2).Exit, test & " interactive exit preserved"
    Set again = EZHighlightsConvert(sld, converted)
    Check again.Id = converted.Id And sld.Shapes.Count = 3, test & " repeat click is idempotent"
    sld.Export deck.Path & "\case-" & test & ".png", "PNG", 1440, 810
    Exit Sub
failed:
    Check False, "case " & test & " runtime " & Err.Number & ": " & Err.Description
End Sub
