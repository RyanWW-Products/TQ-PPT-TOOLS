Attribute VB_Name = "EZHighlights"
Option Explicit

' Native Windows Ink, embedded by PowerPoint in the saved presentation.
' The editable source shape remains in the group with no fill or outline.
Private mDrawEvents As EZHighlightEvents
Private mBusy As Boolean

Public Sub EZHighlightsClick(ByVal control As IRibbonControl)
    Dim selected As Selection, sld As Slide, source As Shape, result As Shape
    Dim items As New Collection, i As Long, message As String, chosen As ShapeRange, isChild As Boolean, stage As String
    On Error GoTo failed
    If mBusy Then Exit Sub
    EZHighlightsCancel
    If Application.Windows.Count = 0 Then Exit Sub
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Open a slide in Normal view to use EZ Highlights.", vbInformation, "EZ Highlights"
        Exit Sub
    End If
    Set sld = ActiveWindow.View.Slide
    Set selected = ActiveWindow.Selection
    stage = "reading selection"
    If selected.Type = ppSelectionShapes Or selected.Type = ppSelectionText Then
        isChild = selected.HasChildShapeRange
        If isChild Then
            Set chosen = selected.ChildShapeRange
        Else
            Set chosen = selected.ShapeRange
        End If
        For i = 1 To chosen.Count
            Set source = chosen(i)
            EZHighlightsValidate source
            items.Add source
        Next
        Application.StartNewUndoEntry
        mBusy = True
        stage = "converting selection"
        If isChild Then
            Set result = ConvertGroupChildren(sld, items)
            EZHighlightReplay.RepairHighlightAnimations sld, result
        Else
            For Each source In items
                Set result = EZHighlightsConvert(sld, source)
                EZHighlightReplay.RepairHighlightAnimations sld, result
            Next
        End If
        stage = "selecting result"
        result.Select
        mBusy = False
    Else
        If ActiveWindow.ViewType = ppViewNormal Then ActiveWindow.Panes(2).Activate
        DoEvents
        Set mDrawEvents = New EZHighlightEvents
        mDrawEvents.Arm Application, ActiveWindow, sld
        Application.CommandBars.ExecuteMso "ShapeRectangle"
    End If
    Exit Sub
failed:
    message = stage & ": " & Err.Description
    mBusy = False
    EZHighlightsCancel
    MsgBox message, vbExclamation, "EZ Highlights"
End Sub

' PowerPoint exposes only the leaf shapes of nested groups through GroupItems.
' Work on a duplicate, ungroup it one level at a time, then restore each wrapper.
' Map animation targets by leaf ID after regrouping invalidates Shape references.
Private Function ConvertGroupChildren(ByVal sld As Slide, ByVal items As Collection) As Shape
    Dim wanted As Object, roots As New Collection, seen As Object
    Dim source As Shape, root As Shape, result As Shape
    Set wanted = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    For Each source In items
        Set root = source.ParentGroup
        If source.Tags("EZHighlight") <> "1" And source.Tags("EZHighlightMember") <> "1" And _
           root.Tags("EZHighlight") <> "1" Then wanted(CStr(source.Id)) = True
        If Not seen.Exists(CStr(root.Id)) Then
            roots.Add root
            seen.Add CStr(root.Id), True
        End If
    Next
    For Each root In roots
        If wanted.Count > 0 Then
            Set result = RebuildRoot(sld, root, wanted)
        Else
            Set result = root
        End If
    Next
    Set ConvertGroupChildren = result
End Function

Private Function RebuildRoot(ByVal sld As Slide, ByVal original As Shape, ByVal wanted As Object) As Shape
    Dim work As Shape, result As Shape, source As Shape, target As Shape, editable As Shape
    Dim originals As New Collection, cleanup As New Collection, mapped As Object, selected As Object
    Dim i As Long, originalZ As Long, e As Long, message As String, stage As String, ids As Variant, stagedId As Variant
    Set mapped = CreateObject("Scripting.Dictionary")
    Set selected = CreateObject("Scripting.Dictionary")
    On Error GoTo failed
    originalZ = original.ZOrderPosition
    stage = "copying group"
    Set work = original.Duplicate()(1)
    cleanup.Add work.Id
    RemoveDuplicateAnimations sld, work
    For i = 1 To original.GroupItems.Count
        Set source = original.GroupItems(i)
        originals.Add source
        work.GroupItems(i).Tags.Add "EZHighlightSource", CStr(source.Id)
        RemoveDuplicateAnimations sld, work.GroupItems(i)
    Next
    work.Left = original.Left: work.Top = original.Top
    stage = "rebuilding group"
    Set result = RebuildBranch(sld, work, wanted, mapped, selected, cleanup)
    stage = "preserving group animations"
    For Each source In originals
        ids = mapped(CStr(source.Id))
        Set target = FindLeaf(result, CLng(ids(0)))
        If selected.Exists(CStr(source.Id)) And CLng(ids(1)) <> 0 Then
            Set editable = FindLeaf(result, CLng(ids(1)))
            ' The visible ink and retained editable text are separate leaves in
            ' the group. Their effects run together, without an extra click.
            CopyTextAnimations sld, source, editable
        End If
        TransferAllAnimations sld, source, target
    Next
    TransferAllAnimations sld, original, result
    Do While result.ZOrderPosition > originalZ + 1
        result.ZOrder msoSendBackward
    Loop
    original.Delete
    Set RebuildRoot = result
    Exit Function
failed:
    e = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    If Not result Is Nothing Then
        For Each source In originals
            ids = mapped(CStr(source.Id))
            Set target = FindLeaf(result, CLng(ids(0)))
            TransferAllAnimations sld, target, source
        Next
        TransferAllAnimations sld, result, original
    End If
    For i = sld.Shapes.Count To 1 Step -1
        For Each stagedId In cleanup
            If sld.Shapes(i).Id = CLng(stagedId) Then
                sld.Shapes(i).Delete
                Exit For
            End If
        Next
    Next
    On Error GoTo 0
    If e = 0 Then e = vbObjectError + 2704
    Err.Raise e, "EZ Highlights", message
End Function

Private Function RebuildBranch(ByVal sld As Slide, ByVal work As Shape, ByVal wanted As Object, _
                               ByVal mapped As Object, ByVal selected As Object, ByVal cleanup As Collection) As Shape
    Dim parts As ShapeRange, members As New Collection, result As Shape, child As Shape, names() As Variant
    Dim style As Shape, ink As Shape, editable As Shape, originalId As String, originalName As String
    Dim x As Single, y As Single, w As Single, h As Single, angle As Single
    Dim flipH As Boolean, flipV As Boolean, i As Long, stage As String, e As Long, message As String
    On Error GoTo failed
    stage = "reading group member"
    originalName = work.Name
    If work.Type = msoGroup And work.Tags("EZHighlight") = "1" Then
        For Each child In work.GroupItems
            originalId = child.Tags("EZHighlightSource")
            child.Tags.Delete "EZHighlightSource"
            mapped.Add originalId, Array(child.Id, 0)
        Next
        Set RebuildBranch = work
        Exit Function
    End If
    If work.Type <> msoGroup Then
        originalId = work.Tags("EZHighlightSource")
        work.Tags.Delete "EZHighlightSource"
        If wanted.Exists(originalId) Then
            Set result = EZHighlightsConvert(sld, work)
            cleanup.Add result.Id
            For Each child In result.GroupItems
                If child.Type = msoInk Or child.Type = msoInkComment Then Set ink = child Else Set editable = child
            Next
            ink.Tags.Add "EZHighlight", "1"
            ink.Name = originalName & " highlight"
            mapped.Add originalId, Array(ink.Id, editable.Id)
            selected.Add originalId, True
        Else
            Set result = work
            mapped.Add originalId, Array(work.Id, 0)
        End If
        Set RebuildBranch = result
        Exit Function
    End If
    If work.ThreeD.Visible = msoTrue Then Err.Raise vbObjectError + 2703, , "Turn off the containing group's 3-D effect before highlighting a child shape."
    stage = "copying group formatting"
    Set style = work.Duplicate()(1)
    cleanup.Add style.Id
    x = work.Left: y = work.Top: w = work.Width: h = work.Height
    angle = work.Rotation
    flipH = (work.HorizontalFlip = msoTrue): flipV = (work.VerticalFlip = msoTrue)
    work.Rotation = 0
    If flipH Then work.Flip msoFlipHorizontal
    If flipV Then work.Flip msoFlipVertical
    stage = "ungrouping working copy"
    Set parts = work.Ungroup
    ReDim names(1 To parts.Count)
    For i = 1 To parts.Count
        cleanup.Add parts(i).Id
        members.Add parts(i)
    Next
    stage = "converting members"
    For i = 1 To members.Count
        Set child = RebuildBranch(sld, members(i), wanted, mapped, selected, cleanup)
        names(i) = child.Id
    Next
    For i = 1 To UBound(names)
        names(i) = SlideShapeIndex(sld, CLng(names(i)))
    Next
    stage = "regrouping members"
    Set result = sld.Shapes.Range(names).Group
    cleanup.Add result.Id
    stage = "restoring group transform"
    result.LockAspectRatio = msoFalse
    result.Width = w: result.Height = h
    result.Left = x: result.Top = y
    If flipH Then result.Flip msoFlipHorizontal
    If flipV Then result.Flip msoFlipVertical
    result.Rotation = angle
    stage = "restoring group formatting"
    CopyGroupProperties style, result
    result.Name = originalName
    style.Delete
    Set RebuildBranch = result
    Exit Function
failed:
    e = Err.Number: message = stage & ": " & Err.Description
    Err.Raise e, "EZ Highlights", message
End Function

Private Function SlideShapeIndex(ByVal sld As Slide, ByVal shapeId As Long) As Long
    Dim i As Long
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).Id = shapeId Then SlideShapeIndex = i: Exit Function
    Next
    Err.Raise vbObjectError + 2705, , "A grouped shape could not be located while highlighting."
End Function

Private Function FindLeaf(ByVal root As Shape, ByVal shapeId As Long) As Shape
    Dim child As Shape
    For Each child In root.GroupItems
        If child.Id = shapeId Then Set FindLeaf = child: Exit Function
    Next
    Err.Raise vbObjectError + 2705, , "A grouped shape could not be located after highlighting."
End Function

Private Sub CopyTextAnimations(ByVal sld As Slide, ByVal source As Shape, ByVal editable As Shape)
    Dim sequence As Sequence
    If Not editable.HasTextFrame Then Exit Sub
    If Not editable.TextFrame.HasText Then Exit Sub
    CopyTextSequence sld.TimeLine.MainSequence, source, editable
    For Each sequence In sld.TimeLine.InteractiveSequences
        CopyTextSequence sequence, source, editable
    Next
End Sub

Private Sub CopyTextSequence(ByVal sequence As Sequence, ByVal source As Shape, ByVal editable As Shape)
    Dim i As Long, effect As Effect, copied As Effect
    For i = sequence.Count To 1 Step -1
        Set effect = sequence(i)
        If effect.Shape.Id = source.Id Then
            Set copied = sequence.Clone(effect)
            copied.Shape = editable
            copied.Exit = effect.Exit
            copied.Timing.Duration = effect.Timing.Duration
            copied.Timing.TriggerDelayTime = effect.Timing.TriggerDelayTime
            copied.Timing.TriggerType = msoAnimTriggerWithPrevious
            copied.MoveTo i + 1
        End If
    Next
End Sub

Private Sub CopyGroupProperties(ByVal source As Shape, ByVal target As Shape)
    Dim i As Long
    target.Name = source.Name
    target.Title = source.Title
    target.AlternativeText = source.AlternativeText
    If source.Visible = msoFalse Then target.Visible = msoFalse
    target.LockAspectRatio = (source.LockAspectRatio = msoTrue)
    For i = 1 To source.Tags.Count
        target.Tags.Add source.Tags.Name(i), source.Tags.Value(i)
    Next
    With target.Shadow
        If source.Shadow.Visible = msoTrue Then
            .Type = source.Shadow.Type
            .ForeColor.RGB = source.Shadow.ForeColor.RGB
            .Transparency = source.Shadow.Transparency
            .Size = source.Shadow.Size
            .Blur = source.Shadow.Blur
            .OffsetX = source.Shadow.OffsetX: .OffsetY = source.Shadow.OffsetY
            .RotateWithShape = source.Shadow.RotateWithShape
        End If
        If source.Shadow.Visible = msoTrue Then .Visible = msoTrue
    End With
    With target.Glow
        If source.Glow.Radius > 0 Then
            .Color.RGB = source.Glow.Color.RGB
            .Transparency = source.Glow.Transparency
        End If
        If source.Glow.Radius > 0 Then .Radius = source.Glow.Radius
    End With
    If source.SoftEdge.Radius > 0 Then target.SoftEdge.Radius = source.SoftEdge.Radius
    With target.Reflection
        If source.Reflection.Type <> msoReflectionTypeNone Then
            .Type = source.Reflection.Type
            .Transparency = source.Reflection.Transparency
            .Size = source.Reflection.Size
            .Blur = source.Reflection.Blur
            .Offset = source.Reflection.Offset
        End If
    End With
End Sub

Public Sub EZHighlightsCancel()
    If Not mDrawEvents Is Nothing Then mDrawEvents.Disarm
    Set mDrawEvents = Nothing
End Sub

Public Sub EZHighlightsDrawn(ByVal sld As Slide, ByVal source As Shape)
    Dim result As Shape, message As String
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo failed
    Application.StartNewUndoEntry
    Set result = EZHighlightsConvert(sld, source)
    result.Select
    mBusy = False
    Exit Sub
failed:
    message = Err.Description
    mBusy = False
    MsgBox "The rectangle was kept, but could not be highlighted. " & message, vbExclamation, "EZ Highlights"
End Sub

Public Sub EZHighlightsValidate(ByVal source As Shape)
    If source.Tags("EZHighlight") = "1" Then Exit Sub
    Select Case source.Type
        Case msoAutoShape, msoFreeform, msoTextBox
        Case Else
            Err.Raise vbObjectError + 2700, "EZ Highlights", _
                "Select a filled shape or text box, including individual shapes inside a group. Pictures, charts, lines and whole groups cannot be converted."
    End Select
    If source.Width < 0.5 Or source.Height < 0.5 Then
        Err.Raise vbObjectError + 2701, "EZ Highlights", "The shape must be at least half a point wide and high."
    End If
    If source.ThreeD.Visible Then
        Err.Raise vbObjectError + 2702, "EZ Highlights", "Turn off the shape's 3-D effect before converting it."
    End If
End Sub

' Public for the application event sink and isolated regression harness.
' Build completely before touching the source; reverse animation retargets on failure.
Public Function EZHighlightsConvert(ByVal sld As Slide, ByVal source As Shape) As Shape
    Dim ink As Shape, editable As Shape, combined As Shape
    Dim sourceName As String, sourceZ As Long, i As Long, e As Long, message As String
    Dim stage As String
    EZHighlightsValidate source
    If source.Tags("EZHighlight") = "1" Then
        Set EZHighlightsConvert = source
        Exit Function
    End If
    sourceName = source.Name: sourceZ = source.ZOrderPosition
    On Error GoTo rollback
    Set ink = EZHighlightInk.CreateHighlightInk(sld, source)
    Set editable = source.Duplicate()(1)
    RemoveDuplicateAnimations sld, editable
    editable.Left = source.Left: editable.Top = source.Top
    editable.Fill.Visible = msoFalse
    editable.Line.Visible = msoFalse
    editable.Tags.Add "EZHighlightMember", "1"
    ' The original geometry, adjustments, text, actions, tags and formatting live here.
    editable.Name = "EZ Highlights editable shape " & editable.Id
    editable.ZOrder msoBringToFront
    Set combined = sld.Shapes.Range(Array(ink.Name, editable.Name)).Group
    combined.AlternativeText = source.AlternativeText
    combined.Title = source.Title
    For i = 1 To source.Tags.Count
        combined.Tags.Add source.Tags.Name(i), source.Tags.Value(i)
    Next
    combined.Tags.Add "EZHighlight", "1"
    combined.LockAspectRatio = source.LockAspectRatio
    combined.Visible = source.Visible
    ' PowerPoint deletes effects still internally attached to a removed source,
    ' even after Effect.Shape is assigned. Clone keeps the full behavior/timing
    ' while establishing an independent animation before the source is removed.
    stage = "preserving animations"
    TransferAllAnimations sld, source, combined
    stage = "stacking"
    Do While combined.ZOrderPosition > sourceZ + 1
        combined.ZOrder msoSendBackward
    Loop
    stage = "source removal"
    source.Delete
    combined.Name = sourceName
    Set EZHighlightsConvert = combined
    Exit Function
rollback:
    e = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    If Not combined Is Nothing Then TransferAllAnimations sld, combined, source
    If Not combined Is Nothing Then
        combined.Delete
    Else
        If Not editable Is Nothing Then editable.Delete
        If Not ink Is Nothing Then ink.Delete
    End If
    On Error GoTo 0
    Err.Raise e, "EZ Highlights", message
End Function

Private Sub TransferSequence(ByVal sequence As Sequence, ByVal source As Shape, ByVal target As Shape)
    Dim i As Long, copiedIndex As Long, sourceIndex As Long, effect As Effect, copied As Effect
    For i = sequence.Count To 1 Step -1
        Set effect = sequence(i)
        If effect.Shape.Id = source.Id Then
            Set copied = sequence.Clone(effect)
            copied.Shape = target
            copied.Exit = effect.Exit
            copied.Timing.Duration = effect.Timing.Duration
            copied.Timing.TriggerDelayTime = effect.Timing.TriggerDelayTime
            ' Assigning an effect to a grouped child can move the clone away
            ' from the end of the sequence. Reacquire it at its actual index.
            copiedIndex = copied.Index
            sourceIndex = effect.Index
            sequence(sourceIndex).Delete
            If copiedIndex > sourceIndex Then copiedIndex = copiedIndex - 1
            sequence(copiedIndex).MoveTo i
        End If
    Next
End Sub

Private Sub TransferAllAnimations(ByVal sld As Slide, ByVal source As Shape, ByVal target As Shape)
    Dim i As Long, j As Long, count As Long, sequence As Sequence, trigger As Shape
    TransferSequence sld.TimeLine.MainSequence, source, target
    For i = sld.TimeLine.InteractiveSequences.Count To 1 Step -1
        Set sequence = sld.TimeLine.InteractiveSequences(i)
        TransferSequence sequence, source, target
        count = sequence.Count
        If count > 0 Then
            Set trigger = sequence(1).Timing.TriggerShape
            If trigger.Id = source.Id Then
                ' TriggerShape moves the effect to the trigger's sequence. Moving
                ' the first remaining effect repeatedly preserves their order.
                For j = 1 To count
                    sequence(1).Timing.TriggerShape = target
                Next
            End If
        End If
    Next
End Sub

Private Sub RemoveDuplicateAnimations(ByVal sld As Slide, ByVal duplicate As Shape)
    Dim sequence As Sequence
    RemoveFromSequence sld.TimeLine.MainSequence, duplicate.Id
    For Each sequence In sld.TimeLine.InteractiveSequences
        RemoveFromSequence sequence, duplicate.Id
    Next
End Sub

Private Sub RemoveFromSequence(ByVal sequence As Sequence, ByVal shapeId As Long)
    Dim i As Long
    For i = sequence.Count To 1 Step -1
        If sequence(i).Shape.Id = shapeId Then sequence(i).Delete
    Next
End Sub
