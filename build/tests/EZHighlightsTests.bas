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

    Print #report, "RESULT | checks=" & checks & " failures=" & failures
    Close #report
    deck.Save
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
