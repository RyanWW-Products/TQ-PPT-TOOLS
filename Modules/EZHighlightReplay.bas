Attribute VB_Name = "EZHighlightReplay"
Option Explicit

Private mAddingReplay As Boolean

' Add a native entrance directly; no preliminary Fade or ungrouping is needed.
Public Sub AddHighlightReplayClick(ByVal control As IRibbonControl)
    Dim window As DocumentWindow, sld As Slide, selected As Selection, chosen As ShapeRange
    Dim originals As New Collection, targets As New Collection, seen As Object
    Dim target As Shape, template As Presentation, path As String, message As String, i As Long
    On Error GoTo failed
    If mAddingReplay Then Exit Sub
    EZHighlightsCancel
    If Application.Windows.Count = 0 Then Exit Sub
    Set window = ActiveWindow
    If window.ViewType <> ppViewNormal And window.ViewType <> ppViewSlide Then
        MsgBox "Open a slide in Normal view and select an EZ Highlight first.", vbInformation, "Add Replay"
        Exit Sub
    End If
    Set sld = window.View.Slide
    Set selected = window.Selection
    Set seen = CreateObject("Scripting.Dictionary")
    If selected.Type = ppSelectionShapes Or selected.Type = ppSelectionText Then
        If selected.HasChildShapeRange Then Set chosen = selected.ChildShapeRange Else Set chosen = selected.ShapeRange
        For Each target In chosen
            originals.Add target
            CollectReplayTargets target, targets, seen
        Next
    End If
    If targets.Count = 0 Then
        MsgBox "Select an EZ Highlight first, then click Add Replay.", vbInformation, "Add Replay"
        Exit Sub
    End If
    mAddingReplay = True
    Application.StartNewUndoEntry
    For Each target In targets
        ' An existing Fade should not remain in front of the new safe entrance.
        RepairHighlightAnimations sld, target
        If Not HasReplayEntrance(sld, target) Then
            If template Is Nothing Then Set template = OpenReplayTemplate(path)
            AppendReplay sld, target, template
        End If
    Next
    window.Activate
    For i = 1 To originals.Count
        originals(i).Select (i = 1)
    Next
    ' The pane exposes the new Replay's Start, Duration and Delay controls even
    ' though PowerPoint still hides the gallery preset for grouped selections.
    On Error Resume Next
    If Not Application.CommandBars.GetPressedMso("AnimationCustom") Then Application.CommandBars.ExecuteMso "AnimationCustom"
    On Error GoTo failed
    GoTo cleanup
failed:
    message = Err.Description
cleanup:
    On Error Resume Next
    If Not template Is Nothing Then template.Close
    If Len(path) > 0 Then Kill path
    mAddingReplay = False
    On Error GoTo 0
    If Len(message) > 0 Then MsgBox "Replay could not be added. " & message, vbExclamation, "Add Replay"
End Sub

Private Sub CollectReplayTargets(ByVal selected As Shape, ByVal targets As Collection, ByVal seen As Object)
    Dim parent As Shape, part As Shape, precedingInk As Shape
    If selected.Type = msoGroup Then
        If selected.Tags("EZHighlight") = "1" Then
            RememberReplayTarget selected, targets, seen
        Else
            ' GroupItems exposes flattened leaves; never animate ordinary
            ' neighbors just because they share an outer group with highlights.
            For Each part In selected.GroupItems
                If part.Tags("EZHighlight") = "1" And (part.Type = msoInk Or part.Type = msoInkComment) Then
                    RememberReplayTarget part, targets, seen
                End If
            Next
        End If
        Exit Sub
    End If
    On Error Resume Next
    Set parent = selected.ParentGroup
    On Error GoTo 0
    If Not parent Is Nothing Then
        If parent.Tags("EZHighlight") = "1" Then
            RememberReplayTarget parent, targets, seen
            Exit Sub
        End If
    End If
    If selected.Tags("EZHighlight") = "1" And (selected.Type = msoInk Or selected.Type = msoInkComment) Then
        RememberReplayTarget selected, targets, seen
    ElseIf selected.Tags("EZHighlightMember") = "1" And Not parent Is Nothing Then
        ' Existing highlights keep their editable member immediately after the
        ' ink in the flattened paint order, including nested/duplicated groups.
        ' Stop at any intervening non-ink member instead of selecting a neighbor.
        For Each part In parent.GroupItems
            If part.Id = selected.Id Then
                If Not precedingInk Is Nothing Then RememberReplayTarget precedingInk, targets, seen
                Exit For
            End If
            Set precedingInk = Nothing
            If part.Tags("EZHighlight") = "1" And (part.Type = msoInk Or part.Type = msoInkComment) Then Set precedingInk = part
        Next
    End If
End Sub

Private Sub RememberReplayTarget(ByVal target As Shape, ByVal targets As Collection, ByVal seen As Object)
    If seen.Exists(CStr(target.Id)) Then Exit Sub
    seen.Add CStr(target.Id), True
    targets.Add target
End Sub

Private Function HasReplayEntrance(ByVal sld As Slide, ByVal target As Shape) As Boolean
    Dim ids As Object, part As Shape, sequence As Sequence
    Set ids = CreateObject("Scripting.Dictionary")
    ids(CStr(target.Id)) = True
    If target.Type = msoGroup Then
        For Each part In target.GroupItems
            ids(CStr(part.Id)) = True
        Next
    Else
        On Error Resume Next
        Set part = target.ParentGroup
        On Error GoTo 0
        If Not part Is Nothing Then ids(CStr(part.Id)) = True
    End If
    If SequenceHasReplay(sld.TimeLine.MainSequence, ids) Then HasReplayEntrance = True: Exit Function
    For Each sequence In sld.TimeLine.InteractiveSequences
        If SequenceHasReplay(sequence, ids) Then HasReplayEntrance = True: Exit Function
    Next
End Function

Private Function SequenceHasReplay(ByVal sequence As Sequence, ByVal ids As Object) As Boolean
    Dim effect As Effect, kind As Long, native As Boolean
    For Each effect In sequence
        If ids.Exists(CStr(effect.Shape.Id)) Then
            If effect.Exit = msoFalse Then
                On Error Resume Next
                Err.Clear
                kind = effect.EffectType
                native = (Err.Number <> 0)
                On Error GoTo 0
                If native And effect.Behaviors.Count = 2 Then
                    If effect.Behaviors(1).Type = msoAnimTypeSet And effect.Behaviors(2).Type = msoAnimTypeProperty Then
                        With effect.Behaviors(2).PropertyEffect.Points
                            If .Count = 2 Then
                                If .Item(1).Value = 0 And .Item(2).Value = 1 Then SequenceHasReplay = True: Exit Function
                            End If
                        End With
                    End If
                End If
            End If
        End If
    Next
End Function

Private Sub AppendReplay(ByVal sld As Slide, ByVal target As Shape, ByVal template As Presentation)
    Dim seed As Shape, seedEffect As Effect, copied As Effect, candidate As Effect
    Dim sequence As Sequence, e As Long, message As String
    On Error GoTo failed
    Set sequence = sld.TimeLine.MainSequence
    template.Slides(1).Shapes(1).Copy
    Set seed = sld.Shapes.Paste()(1)
    For Each candidate In sequence
        If candidate.Shape.Id = seed.Id Then Set seedEffect = candidate: Exit For
    Next
    If seedEffect Is Nothing Then Err.Raise vbObjectError + 2720, , "PowerPoint did not copy the ink animation."
    Set copied = sequence.Clone(seedEffect)
    copied.Shape = target
    copied.Timing.TriggerType = msoAnimTriggerOnPageClick
    copied.Timing.Duration = 2
    copied.Timing.TriggerDelayTime = 0
    copied.MoveTo sequence.Count
    seed.Delete
    Exit Sub
failed:
    e = Err.Number: message = Err.Description
    On Error Resume Next
    If Not copied Is Nothing Then copied.Delete
    If Not seed Is Nothing Then seed.Delete
    On Error GoTo 0
    Err.Raise e, "Add Replay", message
End Sub

' Fade composites Windows Ink through black. Native Replay/Rewind animates
' drawProgress instead. PowerPoint hides those presets for grouped selections
' and does not expose them in MsoAnimEffect, so copy its native preset template.
' Everything is embedded in the slide; playback never needs this add-in.
Public Sub RepairHighlightAnimations(ByVal sld As Slide, ByVal selected As Shape)
    Dim wanted As Object, part As Shape, sequence As Sequence, effect As Effect
    Dim pending As New Collection, template As Presentation, path As String
    Dim i As Long, e As Long, message As String
    Set wanted = CreateObject("Scripting.Dictionary")
    wanted(CStr(selected.Id)) = True
    If selected.Type = msoGroup Then
        For Each part In selected.GroupItems
            wanted(CStr(part.Id)) = True
        Next
    Else
        On Error Resume Next
        Set part = selected.ParentGroup
        On Error GoTo 0
        If Not part Is Nothing Then wanted(CStr(part.Id)) = True
    End If
    On Error GoTo failed
    CollectFades sld.TimeLine.MainSequence, wanted, pending
    For Each sequence In sld.TimeLine.InteractiveSequences
        CollectFades sequence, wanted, pending
    Next
    If pending.Count = 0 Then Exit Sub
    Set template = OpenReplayTemplate(path)
    ' Snapshot the original effects before pasting template effects changes
    ' sequence counts. Work backwards to retain the existing click order.
    For i = pending.Count To 1 Step -1
        Set effect = pending(i)
        ReplaceFade sld, effect, template
    Next
    GoTo cleanup
failed:
    e = Err.Number: message = Err.Description
cleanup:
    On Error Resume Next
    If Not template Is Nothing Then template.Close
    If Len(path) > 0 Then Kill path
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "EZ Highlights animation", message
End Sub

Private Sub CollectFades(ByVal sequence As Sequence, ByVal wanted As Object, ByVal pending As Collection)
    Dim effect As Effect, kind As Long, readable As Boolean
    For Each effect In sequence
        If wanted.Exists(CStr(effect.Shape.Id)) Then
            If HasHighlightInk(effect.Shape) Then
                ' Reading EffectType raises E_FAIL for native Replay/Rewind.
                On Error Resume Next
                Err.Clear
                kind = effect.EffectType
                readable = (Err.Number = 0)
                On Error GoTo 0
                If readable And kind = msoAnimEffectFade Then pending.Add effect
            End If
        End If
    Next
End Sub

Private Function HasHighlightInk(ByVal shape As Shape) As Boolean
    Dim part As Shape, parent As Shape
    If shape.Type = msoGroup Then
        For Each part In shape.GroupItems
            If part.Type = msoInk Or part.Type = msoInkComment Then
                If shape.Tags("EZHighlight") = "1" Or part.Tags("EZHighlight") = "1" Then
                    HasHighlightInk = True
                    Exit Function
                End If
            End If
        Next
    ElseIf shape.Type = msoInk Or shape.Type = msoInkComment Then
        If shape.Tags("EZHighlight") = "1" Then
            HasHighlightInk = True
        Else
            On Error Resume Next
            Set parent = shape.ParentGroup
            If Not parent Is Nothing Then HasHighlightInk = (parent.Tags("EZHighlight") = "1")
            On Error GoTo 0
        End If
    End If
End Function

Private Function OpenReplayTemplate(ByRef path As String) As Presentation
    Dim xml As Object, node As Object, stream As Object, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".pptx")
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("data")
    node.DataType = "bin.base64"
    node.Text = InkReplayTemplateBase64()
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write node.nodeTypedValue
    stream.SaveToFile path, 2
    stream.Close
    Set OpenReplayTemplate = Application.Presentations.Open(path, msoTrue, msoFalse, msoFalse)
End Function

Private Sub ReplaceFade(ByVal sld As Slide, ByVal original As Effect, ByVal template As Presentation)
    Dim sequence As Sequence, seed As Shape, seedEffect As Effect, copied As Effect
    Dim target As Shape, targets As New Collection, made As New Collection
    Dim isExit As Boolean, i As Long, first As Boolean, e As Long, message As String
    Dim indexes() As Long, j As Long, swap As Long, stage As String
    On Error GoTo failed
    stage = "reading animation"
    Set sequence = original.Parent
    Set target = original.Shape
    If target.Type = msoGroup Then
        For Each target In original.Shape.GroupItems
            targets.Add target
        Next
    Else
        targets.Add target
    End If
    isExit = (original.Exit = msoTrue)
    If isExit Then i = 2 Else i = 1
    stage = "copying native animation"
    template.Slides(i).Shapes(1).Copy
    Set seed = sld.Shapes.Paste()(1)
    For Each copied In sld.TimeLine.MainSequence
        If copied.Shape.Id = seed.Id Then Set seedEffect = copied: Exit For
    Next
    If seedEffect Is Nothing Then Err.Raise vbObjectError + 2720, , "PowerPoint did not copy the ink animation."
    stage = "retaining click target"
    ' MoveBefore also moves between main and interactive sequences, retaining
    ' the destination's trigger. TriggerShape's setter fails for native Replay.
    seedEffect.MoveBefore original
    Set sequence = original.Parent
    first = True
    For Each target In targets
        stage = "cloning native animation"
        If target.Type = msoInk Or target.Type = msoInkComment Then
            Set copied = sequence.Clone(seedEffect)
        Else
            ' Editable text and other ordinary group members keep their Fade.
            Set copied = sequence.Clone(original)
        End If
        made.Add copied
        copied.Shape = target
        stage = "retaining animation timing"
        CopyTiming original.Timing, copied.Timing, first
        copied.MoveBefore original
        first = False
    Next
    original.Delete
    seed.Delete
    Exit Sub
failed:
    e = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    ' Capture all indices before deletion invalidates Effect references.
    If made.Count > 0 Then
        ReDim indexes(1 To made.Count)
        For i = 1 To made.Count
            indexes(i) = made(i).Index
        Next
        For i = 1 To made.Count
            For j = i + 1 To made.Count
                If indexes(j) > indexes(i) Then swap = indexes(i): indexes(i) = indexes(j): indexes(j) = swap
            Next
        Next
        For i = 1 To made.Count
            If indexes(i) > 0 Then sequence(indexes(i)).Delete
        Next
    End If
    If Not seed Is Nothing Then seed.Delete
    On Error GoTo 0
    Err.Raise e, "EZ Highlights animation", message
End Sub

Private Sub CopyTiming(ByVal source As Timing, ByVal target As Timing, ByVal first As Boolean)
    With target
        .Duration = source.Duration
        .Speed = source.Speed
        .TriggerDelayTime = source.TriggerDelayTime
        If source.RepeatCount > 0 Then .RepeatCount = source.RepeatCount
        If source.RepeatDuration > 0 Then .RepeatDuration = source.RepeatDuration
        .AutoReverse = source.AutoReverse
        .SmoothStart = source.SmoothStart
        .SmoothEnd = source.SmoothEnd
        .Accelerate = source.Accelerate
        .Decelerate = source.Decelerate
        .RewindAtEnd = source.RewindAtEnd
        .Restart = source.Restart
        .BounceEnd = source.BounceEnd
        .BounceEndIntensity = source.BounceEndIntensity
        If first Then .TriggerType = source.TriggerType Else .TriggerType = msoAnimTriggerWithPrevious
    End With
End Sub
