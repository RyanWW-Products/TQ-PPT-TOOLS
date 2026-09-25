Attribute VB_Name = "EZHighlights"
Option Explicit

' Native Windows Ink, embedded by PowerPoint in the saved presentation.
' The editable source shape remains in the group with no fill or outline.
Private mDrawEvents As EZHighlightEvents
Private mBusy As Boolean

Public Sub EZHighlightsClick(ByVal control As IRibbonControl)
    Dim selected As Selection, sld As Slide, source As Shape, result As Shape
    Dim items As New Collection, i As Long, message As String
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
    If selected.Type = ppSelectionShapes Or selected.Type = ppSelectionText Then
        If selected.HasChildShapeRange Then
            MsgBox "Select the whole shape outside its group first.", vbInformation, "EZ Highlights"
            Exit Sub
        End If
        For i = 1 To selected.ShapeRange.Count
            Set source = selected.ShapeRange(i)
            EZHighlightsValidate source
            items.Add source
        Next
        Application.StartNewUndoEntry
        mBusy = True
        For Each source In items
            Set result = EZHighlightsConvert(sld, source)
        Next
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
    message = Err.Description
    mBusy = False
    EZHighlightsCancel
    MsgBox message, vbExclamation, "EZ Highlights"
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
                "Select a filled shape or text box. Pictures, charts, lines and existing groups cannot be converted."
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
    Dim i As Long, effect As Effect, copied As Effect
    For i = sequence.Count To 1 Step -1
        Set effect = sequence(i)
        If effect.Shape.Id = source.Id Then
            Set copied = sequence.Clone(effect)
            copied.Shape = target
            copied.Exit = effect.Exit
            copied.Timing.Duration = effect.Timing.Duration
            copied.Timing.TriggerDelayTime = effect.Timing.TriggerDelayTime
            sequence(i).Delete
            sequence(sequence.Count).MoveTo i
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
