Attribute VB_Name = "EZHighlightInk"
Option Explicit

' Windows Ink uses HIMETRIC units (1/100 mm). MaskPen is the native
' highlighter raster operation: yellow stays saturated and black stays black.
Private Const INK_PER_POINT As Double = 2540# / 72#
Private Type GdiStartup
    Version As Long
    Callback As LongPtr
    SuppressThread As Long
    SuppressCodecs As Long
End Type
Private Type GdiRect
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type
Private Type GdiBitmapData
    Width As Long
    Height As Long
    Stride As Long
    PixelFormat As Long
    Scan0 As LongPtr
    Reserved As LongPtr
End Type
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputData As GdiStartup, ByVal outputData As LongPtr) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef value As Long) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef value As Long) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef area As GdiRect, ByVal flags As Long, ByVal format As Long, ByRef data As GdiBitmapData) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef data As GdiBitmapData) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal bitmap As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef target As Any, ByVal source As LongPtr, ByVal length As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Function CreateHighlightInk(ByVal sld As Slide, ByVal source As Shape) As Shape
    Dim ink As Object, result As Shape, packet(0 To 3) As Long, registration As Object
    Dim pasted As ShapeRange, part As Shape, e As Long, message As String
    On Error GoTo failed
    Set ink = CreateObject("msinkaut.InkObject")
    ' Transparent registration stroke gives the packet geometry a real width AND
    ' height. Without it PowerPoint mis-scales the PNG fallback for straight ink.
    packet(2) = CLng(source.Width * INK_PER_POINT)
    packet(3) = CLng(source.Height * INK_PER_POINT)
    Set registration = ink.CreateStroke(packet, Empty)
    With registration.DrawingAttributes
        .Width = 1: .Height = 1
        .Transparency = 255
        .IgnorePressure = True
        .FitToCurve = False
    End With
    If source.Type = msoTextBox Then
        AddBand ink, 0, 0, source.Width, source.Height
    ElseIf source.Type = msoAutoShape And source.AutoShapeType = msoShapeRectangle Then
        AddBand ink, 0, 0, source.Width, source.Height
    Else
        AddShapeMask ink, sld, source
    End If
    ' Copy native serialized ink, not a rendered bitmap; no third-party DLLs,
    ' linked images, slide macros or add-in are needed to display the saved ink.
    Set pasted = PasteInk(sld, ink)
    For Each part In pasted
        If part.Type <> msoInk And part.Type <> msoInkComment Then
            Err.Raise vbObjectError + 2710, , "PowerPoint did not paste native ink. The original shape was kept."
        End If
    Next
    ' PowerPoint pastes one ink shape per stroke. Keep ALL of them and resize
    ' their group, never a zero-height horizontal stroke on its own.
    Set result = pasted.Group
    result.LockAspectRatio = msoFalse
    result.Width = source.Width: result.Height = source.Height
    result.Left = source.Left: result.Top = source.Top
    result.Rotation = source.Rotation
    result.Name = "EZ Highlights ink " & CStr(result.Id)
    Set CreateHighlightInk = result
    Exit Function
failed:
    e = Err.Number: message = Err.Description
    On Error Resume Next
    If Not result Is Nothing Then
        result.Delete
    ElseIf Not pasted Is Nothing Then
        pasted.Delete
    End If
    On Error GoTo 0
    Err.Raise e, "EZ Highlights", message
End Function

Private Function PasteInk(ByVal sld As Slide, ByVal ink As Object) As ShapeRange
    Dim attempt As Long, e As Long, message As String
    For attempt = 1 To 10
        On Error Resume Next
        Err.Clear
        ink.ClipboardCopy Nothing, 7, 0
        e = Err.Number: message = Err.Description
        If e = 0 Then
            DoEvents
            Set PasteInk = sld.Shapes.Paste()
            e = Err.Number: message = Err.Description
        End If
        On Error GoTo 0
        If e = 0 Then Exit Function
        Sleep 40
        DoEvents
    Next
    Err.Raise e, "EZ Highlights", message
End Function

Private Sub AddBand(ByVal ink As Object, ByVal left As Double, ByVal top As Double, _
                    ByVal width As Double, ByVal height As Double)
    Dim packet(0 To 3) As Long, stroke As Object, penWidth As Double
    If width <= 0 Or height <= 0 Then Exit Sub
    penWidth = 1 ' One HIMETRIC unit, with square ends.
    packet(0) = CLng(left * INK_PER_POINT + penWidth / 2)
    packet(1) = CLng((top + height / 2) * INK_PER_POINT)
    packet(2) = CLng((left + width) * INK_PER_POINT - penWidth / 2)
    packet(3) = packet(1)
    Set stroke = ink.CreateStroke(packet, Empty)
    With stroke.DrawingAttributes
        .Color = vbYellow
        .PenTip = 1 ' IPT_Rectangle
        .Width = penWidth
        .Height = CSng(height * INK_PER_POINT)
        .Transparency = 0
        .RasterOperation = 9 ' IRO_MaskPen
        .IgnorePressure = True
        .FitToCurve = False
    End With
End Sub

' PowerPoint renders a temporary, unrotated copy to obtain the exact silhouette
' of adjusted autoshapes/freeforms (including holes). Scanlines become native
' vector ink strokes; the image itself is never inserted into the presentation.
Private Sub AddShapeMask(ByVal ink As Object, ByVal sld As Slide, ByVal source As Shape)
    Dim mask As Shape, path As String, fso As Object, exportScale As Double
    Dim token As LongPtr, bitmap As LongPtr, startup As GdiStartup
    Dim area As GdiRect, data As GdiBitmapData, locked As Boolean
    Dim pixels() As Byte, x As Long, y As Long, startX As Long, index As Long
    Dim previous As Object, current As Object, key As Variant, band As Variant
    Dim e As Long, message As String
    On Error GoTo failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".png")
    Set mask = source.Duplicate()(1)
    mask.Rotation = 0
    mask.Visible = msoTrue
    mask.Line.Visible = msoFalse
    mask.Shadow.Visible = msoFalse
    mask.Glow.Radius = 0
    mask.SoftEdge.Radius = 0
    mask.Fill.Solid
    mask.Fill.ForeColor.RGB = vbBlack
    mask.Fill.Transparency = 0
    If mask.HasTextFrame Then mask.TextFrame.TextRange.Text = ""
    exportScale = 3 ' 4 pixels per point after PowerPoint's 96-DPI conversion.
    If source.Width * 4 > 2048 Then exportScale = 1536# / source.Width
    If source.Height * exportScale * 4 / 3 > 2048 Then exportScale = 1536# / source.Height
    mask.Export path, ppShapeFormatPNG, CLng(sld.Parent.PageSetup.SlideWidth * exportScale), _
                CLng(sld.Parent.PageSetup.SlideHeight * exportScale), ppRelativeToSlide
    mask.Delete: Set mask = Nothing
    startup.Version = 1
    If GdiplusStartup(token, startup, 0) <> 0 Then Err.Raise vbObjectError + 2711, , "Windows graphics could not start."
    If GdipCreateBitmapFromFile(StrPtr(path), bitmap) <> 0 Then Err.Raise vbObjectError + 2712, , "The shape outline could not be read."
    GdipGetImageWidth bitmap, area.Width
    GdipGetImageHeight bitmap, area.Height
    If area.Width < 1 Or area.Height < 1 Then Err.Raise vbObjectError + 2713, , "The shape has no filled area."
    If GdipBitmapLockBits(bitmap, area, 1, &H26200A, data) <> 0 Then Err.Raise vbObjectError + 2714, , "The shape outline could not be locked."
    locked = True
    ReDim pixels(0 To area.Width * 4 - 1)
    Set previous = CreateObject("Scripting.Dictionary")
    For y = 0 To area.Height - 1
        CopyMemory pixels(0), data.Scan0 + CLngPtr(y) * data.Stride, area.Width * 4
        Set current = CreateObject("Scripting.Dictionary")
        x = 0
        Do While x < area.Width
            If pixels(x * 4 + 3) >= 128 Then
                startX = x
                Do While x < area.Width
                    If pixels(x * 4 + 3) < 128 Then Exit Do
                    x = x + 1
                Loop
                key = CStr(startX) & ":" & CStr(x)
                If previous.Exists(key) Then
                    band = previous(key): band(3) = y + 1
                    previous.Remove key
                Else
                    band = Array(startX, y, x, y + 1)
                End If
                current.Add key, band
            Else
                x = x + 1
            End If
        Loop
        For Each key In previous.Keys
            EmitBand ink, previous(key), source.Width / area.Width, source.Height / area.Height
        Next
        Set previous = current
    Next
    For Each key In previous.Keys
        EmitBand ink, previous(key), source.Width / area.Width, source.Height / area.Height
    Next
    If ink.Strokes.Count < 2 Then Err.Raise vbObjectError + 2715, , "The shape has no filled area."
    GoTo cleanup
failed:
    e = Err.Number: message = Err.Description
cleanup:
    On Error Resume Next
    If locked Then GdipBitmapUnlockBits bitmap, data
    If bitmap <> 0 Then GdipDisposeImage bitmap
    If token <> 0 Then GdiplusShutdown token
    If Not mask Is Nothing Then mask.Delete
    If Len(path) > 0 Then
        If fso.FileExists(path) Then fso.DeleteFile path
    End If
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "EZ Highlights", message
End Sub

Private Sub EmitBand(ByVal ink As Object, ByVal band As Variant, ByVal sx As Double, ByVal sy As Double)
    AddBand ink, CDbl(band(0)) * sx, CDbl(band(1)) * sy, _
            CDbl(band(2) - band(0)) * sx, CDbl(band(3) - band(1)) * sy
End Sub
