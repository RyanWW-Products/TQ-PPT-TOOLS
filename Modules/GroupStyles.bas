Attribute VB_Name = "GroupStyles"
Option Explicit

'====================================================
' (A) File I/O: Get/Load/Save Style Definitions
' Now using the presentation�s custom document properties.
'====================================================
Function LoadStyleDefinitions() As Object
    Dim defs As Object, curStyle As Object
    Dim xml As String
    Dim lines() As String, i As Long, line As String
    Set defs = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    xml = ActivePresentation.CustomDocumentProperties("StyleDefinitions").Value
    On Error GoTo 0
    
    If xml <> "" Then
        lines = Split(xml, vbCrLf)
        For i = 0 To UBound(lines)
            line = Trim(lines(i))
            If line <> "" And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
                If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                    Dim styleName As String
                    styleName = Mid(line, 2, Len(line) - 2)
                    Set curStyle = CreateObject("Scripting.Dictionary")
                    Set defs(styleName) = curStyle
                ElseIf InStr(line, "=") > 0 Then
                    Dim parts() As String
                    parts = Split(line, "=", 2)
                    If Not curStyle Is Nothing Then
                        curStyle(Trim(parts(0))) = Trim(parts(1))
                    End If
                End If
            End If
        Next i
    End If
    
    Set LoadStyleDefinitions = defs
End Function


Sub SaveStyleDefinitions(defs As Object)
    Dim ini As String
    Dim styleName As Variant, propName As Variant
    Dim styleDict As Object

    ini = ""
    For Each styleName In defs.Keys
        ini = ini & "[" & styleName & "]" & vbCrLf
        Set styleDict = defs(styleName)
        For Each propName In styleDict.Keys
            ' CleanValue strips CR/LF so a value can never fabricate a bogus key or [section]
            ini = ini & propName & "=" & CleanValue(CStr(styleDict(propName))) & vbCrLf
        Next propName
        ini = ini & vbCrLf
    Next styleName

    On Error Resume Next
    ActivePresentation.CustomDocumentProperties("StyleDefinitions").Value = ini
    If Err.Number <> 0 Then
        Err.Clear
        ActivePresentation.CustomDocumentProperties.Add Name:="StyleDefinitions", _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=ini
    End If
    On Error GoTo 0

    ' A custom document property can silently TRUNCATE a long value. Read it back and
    ' warn instead of corrupting the style DB invisibly.
    Dim check As String
    On Error Resume Next
    check = ActivePresentation.CustomDocumentProperties("StyleDefinitions").Value
    On Error GoTo 0
    ' Compare CR-stripped lengths so a CRLF->LF normalization on store isn't mistaken for truncation;
    ' only a genuinely SHORTER readback means data was lost.
    If Len(Replace(check, vbCr, "")) < Len(Replace(ini, vbCr, "")) Then
        MsgBox "The style definitions may not have fully saved (" & Len(check) & " of " & Len(ini) & _
               " characters stored)." & vbCrLf & vbCrLf & _
               "There are likely more/larger styles than this presentation's property storage allows.", _
               vbExclamation, "Group Styles"
    End If
End Sub

' Keep a serialized value on a single line so it cannot fabricate INI structure.
Private Function CleanValue(ByVal v As String) As String
    CleanValue = Replace(Replace(v, vbCr, " "), vbLf, " ")
End Function


'====================================================
' (B) Save Style Definitions
'----- Single Style --------------------------------------------------
Sub SaveSingleStyleDefinition(styleName As String, shp As Shape)
    Dim defs As Object, def As Object
    Set defs = LoadStyleDefinitions()
    Set def = CreateObject("Scripting.Dictionary")
    
    With shp.TextFrame.TextRange.Font
        def("FontName") = .Name
        def("FontSize") = .Size
        def("Bold") = .Bold
        def("Italic") = .Italic
        def("Underline") = .Underline
        def("FontColor") = .Color.RGB
    End With
    With shp.TextFrame.TextRange.ParagraphFormat
        def("Alignment") = .Alignment
        def("SpaceBefore") = .SpaceBefore
        def("SpaceAfter") = .SpaceAfter
    End With
    With shp.TextFrame
        def("MarginLeft") = .MarginLeft
        def("MarginRight") = .MarginRight
        def("MarginTop") = .MarginTop
        def("MarginBottom") = .MarginBottom
        def("VerticalAnchor") = .VerticalAnchor
        def("WordWrap") = .WordWrap
    End With
    On Error Resume Next
    def("AutoSize") = shp.TextFrame2.AutoSize
    On Error GoTo 0
    If shp.Fill.Visible Then
        def("FillColor") = shp.Fill.ForeColor.RGB
        def("FillTransparency") = shp.Fill.Transparency
    End If
    If shp.line.Visible Then
        def("LineColor") = shp.line.ForeColor.RGB
        def("LineWeight") = shp.line.Weight
    End If
    def("Dual") = False
    
    
    ' Capture the shape�s shadow properties
On Error Resume Next
def("TextShadow") = shp.TextFrame2.TextRange.Font.Shadow.Visible
If shp.TextFrame2.TextRange.Font.Shadow.Visible Then
    def("TextShadowColor") = shp.TextFrame2.TextRange.Font.Shadow.ForeColor.RGB
    def("TextShadowOffsetX") = shp.TextFrame2.TextRange.Font.Shadow.OffsetX
    def("TextShadowOffsetY") = shp.TextFrame2.TextRange.Font.Shadow.OffsetY
    def("TextShadowTransparency") = shp.TextFrame2.TextRange.Font.Shadow.Transparency
End If
On Error GoTo 0


    
    
    Set defs(styleName) = def
    SaveStyleDefinitions defs
    Debug.Print "Saved single style: " & styleName
End Sub

'----- Dual Style ----------------------------------------------------
Sub SaveDualStyleDefinition(styleName As String, shp As Shape)
    Dim defs As Object, def As Object
    Set defs = LoadStyleDefinitions()
    Set def = CreateObject("Scripting.Dictionary")
    
    Dim tr As TextRange
    Dim paraCount As Long
    Dim line1TR As TextRange, line2TR As TextRange
    Dim startOfLine2 As Long, lengthOfLine2 As Long
    Set tr = shp.TextFrame.TextRange
    paraCount = tr.Paragraphs.count
    If paraCount < 2 Then
        MsgBox "Shape does not contain 2 or more paragraphs. Use a real Enter key for a paragraph break.", vbExclamation
        Exit Sub
    End If
    
    Set line1TR = tr.Paragraphs(1)
    startOfLine2 = tr.Paragraphs(2).Start
    lengthOfLine2 = tr.Length - (startOfLine2 - 1)
    Set line2TR = tr.Characters(startOfLine2, lengthOfLine2)
    
    With line1TR.Font
        def("Line1_FontName") = .Name
        def("Line1_FontSize") = .Size
        def("Line1_Bold") = .Bold
        def("Line1_Italic") = .Italic
        def("Line1_Underline") = .Underline
        def("Line1_FontColor") = .Color.RGB
    End With
    With line1TR.ParagraphFormat
        def("Line1_Alignment") = .Alignment
        def("Line1_SpaceBefore") = .SpaceBefore
        def("Line1_SpaceAfter") = .SpaceAfter
    End With
    
    With line2TR.Font
        def("Line2_FontName") = .Name
        def("Line2_FontSize") = .Size
        def("Line2_Bold") = .Bold
        def("Line2_Italic") = .Italic
        def("Line2_Underline") = .Underline
        def("Line2_FontColor") = .Color.RGB
    End With
    With tr.Paragraphs(2).ParagraphFormat
        def("Line2_Alignment") = .Alignment
        def("Line2_SpaceBefore") = .SpaceBefore
        def("Line2_SpaceAfter") = .SpaceAfter
    End With
    
    With shp.TextFrame
        def("MarginLeft") = .MarginLeft
        def("MarginRight") = .MarginRight
        def("MarginTop") = .MarginTop
        def("MarginBottom") = .MarginBottom
        def("VerticalAnchor") = .VerticalAnchor
        def("WordWrap") = .WordWrap
    End With
    On Error Resume Next
    def("AutoSize") = shp.TextFrame2.AutoSize
    On Error GoTo 0
    If shp.Fill.Visible Then
        def("FillColor") = shp.Fill.ForeColor.RGB
        def("FillTransparency") = shp.Fill.Transparency
    End If
    If shp.line.Visible Then
        def("LineColor") = shp.line.ForeColor.RGB
        def("LineWeight") = shp.line.Weight
    End If
    def("Dual") = True
    
    
On Error Resume Next
def("TextShadow") = shp.TextFrame2.TextRange.Font.Shadow.Visible
If shp.TextFrame2.TextRange.Font.Shadow.Visible Then
    def("TextShadowColor") = shp.TextFrame2.TextRange.Font.Shadow.ForeColor.RGB
    def("TextShadowOffsetX") = shp.TextFrame2.TextRange.Font.Shadow.OffsetX
    def("TextShadowOffsetY") = shp.TextFrame2.TextRange.Font.Shadow.OffsetY
    def("TextShadowTransparency") = shp.TextFrame2.TextRange.Font.Shadow.Transparency
End If
On Error GoTo 0


    
    Set defs(styleName) = def
    SaveStyleDefinitions defs
    Debug.Print "Saved dual style: " & styleName
End Sub

'====================================================
' (C) DEFINE STYLE (Combined)
' Presents a list of 9 style categories:
'   1. Title 1
'   2. Subtitle
'   3. Body
'   4. Title 2
'   5. Title 3
'   6. Dual
'   7. Date Box
'   8. Entry Box
'   9. Entry Body w/ Title  (will not auto-update)
'
' If the selected shape is a group, automatically list its text boxes
' by number so the user can simply enter a number.
'====================================================
Sub DefineStyleCombined()
    Dim choice As String, styleName As String
    Dim prompt As String
    Dim shp As Shape
    Dim paraCount As Long
    
    If ActiveWindow.Selection Is Nothing Or ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a shape with text.", vbExclamation
        Exit Sub
    End If
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    ' If a group is selected, check for text-containing sub-shapes.
    If shp.Type = msoGroup Then
        Dim subShp As Shape, possibleList As String, countItems As Long
        countItems = 0
        possibleList = ""
        For Each subShp In shp.GroupItems
            If subShp.HasTextFrame Then
                If subShp.TextFrame.HasText Then
                    countItems = countItems + 1
                    possibleList = possibleList & countItems & ": " & subShp.Name & vbCrLf
                End If
            End If
        Next subShp
        If countItems = 0 Then
            MsgBox "No text-containing shape found in the group.", vbExclamation
            Exit Sub
        ElseIf countItems = 1 Then
            For Each subShp In shp.GroupItems
                If subShp.HasTextFrame And subShp.TextFrame.HasText Then
                    Set shp = subShp
                    Exit For
                End If
            Next subShp
        Else
            Dim chosenNumber As String, chosenIndex As Long, j As Long
            chosenNumber = InputBox("Multiple text boxes found in the group:" & vbCrLf & possibleList & _
                                      "Enter the number of the shape to define:", "Select Text Box")
            If Trim(chosenNumber) = "" Then
                MsgBox "No shape selected.", vbExclamation
                Exit Sub
            End If
            If Not IsNumeric(Trim(chosenNumber)) Then
                MsgBox "Please enter a number.", vbExclamation
                Exit Sub
            End If
            chosenIndex = CLng(Trim(chosenNumber))
            If chosenIndex < 1 Or chosenIndex > countItems Then
                MsgBox "Invalid number selected.", vbExclamation
                Exit Sub
            End If
            j = 0
            For Each subShp In shp.GroupItems
                If subShp.HasTextFrame And subShp.TextFrame.HasText Then
                    j = j + 1
                    If j = chosenIndex Then
                        Set shp = subShp
                        Exit For
                    End If
                End If
            Next subShp
        End If
    End If
    
    paraCount = shp.TextFrame.TextRange.Paragraphs.count
    
    prompt = "Select a style to define:" & vbCrLf & _
             "1. Title 1" & vbCrLf & _
             "2. Subtitle" & vbCrLf & _
             "3. Body" & vbCrLf & _
             "4. Title 2" & vbCrLf & _
             "5. Title 3" & vbCrLf & _
             "6. Dual" & vbCrLf & _
             "7. Date Box" & vbCrLf & _
             "8. Entry Box" & vbCrLf & _
             "9. Entry Body w/ Title"
    choice = InputBox(prompt, "Define Style")
    If Trim(choice) = "" Then
        MsgBox "No style selected.", vbExclamation
        Exit Sub
    End If
    
    Select Case Trim(choice)
        Case "1": styleName = "Title 1"
        Case "2": styleName = "Subtitle"
        Case "3": styleName = "Body"
        Case "4": styleName = "Title 2"
        Case "5": styleName = "Title 3"
        Case "6": styleName = "Dual"
        Case "7": styleName = "Date Box"
        Case "8": styleName = "Entry Box"
        Case "9": styleName = "Entry Body w/ Title"
        Case Else:
            MsgBox "Invalid selection.", vbExclamation
            Exit Sub
    End Select
    
    Debug.Print "Defining style: " & styleName
    If styleName = "Dual" Or styleName = "Entry Body w/ Title" Then
    If paraCount < 2 Then
        MsgBox "Selected shape does not contain two or more paragraphs. Please use a real paragraph break.", vbExclamation
        Exit Sub
    Else
        SaveDualStyleDefinition styleName, shp
    End If
Else
    SaveSingleStyleDefinition styleName, shp
End If

    MsgBox "Style '" & styleName & "' defined.", vbInformation
    UpdateAllGroupStylesCombined
End Sub

'====================================================
' (D) ASSIGN STYLE (Combined)
' Presents the same 9 options; the chosen style name is stored in the shape's "GroupStyle" tag.
'====================================================
Sub AssignToGroupStyleCombined(control As IRibbonControl)
    Dim choice As String, styleName As String
    Dim prompt As String
    Dim shp As Shape, shpRange As ShapeRange
    
    If ActiveWindow.Selection Is Nothing Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set shpRange = ActiveWindow.Selection.ShapeRange
    Else
        MsgBox "Please select shapes.", vbExclamation
        Exit Sub
    End If
    
    prompt = "Select a style for the selected shapes:" & vbCrLf & _
             "1. Title 1" & vbCrLf & _
             "2. Subtitle" & vbCrLf & _
             "3. Body" & vbCrLf & _
             "4. Title 2" & vbCrLf & _
             "5. Title 3" & vbCrLf & _
             "6. Dual" & vbCrLf & _
             "7. Date Box" & vbCrLf & _
             "8. Entry Box" & vbCrLf & _
             "9. Entry Body w/ Title"
    choice = InputBox(prompt, "Assign Group Style")
    If Trim(choice) = "" Then
        MsgBox "No style selected.", vbExclamation
        Exit Sub
    End If
    Select Case Trim(choice)
        Case "1": styleName = "Title 1"
        Case "2": styleName = "Subtitle"
        Case "3": styleName = "Body"
        Case "4": styleName = "Title 2"
        Case "5": styleName = "Title 3"
        Case "6": styleName = "Dual"
        Case "7": styleName = "Date Box"
        Case "8": styleName = "Entry Box"
        Case "9": styleName = "Entry Body w/ Title"
        Case Else:
            MsgBox "Invalid selection.", vbExclamation
            Exit Sub
    End Select
    
    Dim shpItem As Shape, tagged As Long
    tagged = 0
    For Each shpItem In shpRange
        TagWithStyle shpItem, styleName, tagged
    Next shpItem
    If tagged = 0 Then
        MsgBox "None of the selected shapes contain text to style.", vbExclamation
        Exit Sub
    End If
    MsgBox tagged & " shape(s) assigned to style '" & styleName & "'.", vbInformation
    UpdateAllGroupStylesCombined
End Sub

' Tag a shape - or, for a GROUP, its text-bearing descendants - with the GroupStyle tag.
' UpdateShapesInContainer only reads the tag on non-group leaves, so tagging a group
' itself is a silent no-op; drill in instead. Mirrors TimelineEntry's per-child tagging.
Private Sub TagWithStyle(ByVal shp As Shape, ByVal styleName As String, ByRef tagged As Long)
    Dim child As Shape
    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            TagWithStyle child, styleName, tagged
        Next child
    ElseIf shp.HasTextFrame Then
        shp.Tags.Add "GroupStyle", styleName
        tagged = tagged + 1
    End If
End Sub

'====================================================
' (E) UPDATE ALL GROUP STYLES (Combined)
' Recursively processes all shapes (and group items) in the presentation.
' It skips auto-updating shapes assigned "Entry Body w/ Title".
'====================================================
Sub UpdateAllGroupStylesCombined()
    Dim defs As Object
    Set defs = LoadStyleDefinitions()
    
    Dim sld As slide
    For Each sld In ActivePresentation.Slides
    UpdateShapesInContainer sld.Shapes, defs
Next sld

 
End Sub

Sub UpdateShapesInContainer(shapesCollection As Object, defs As Object)
    Dim shp As Object, styleName As String
    For Each shp In shapesCollection
        If shp.Type = msoGroup Then
            UpdateShapesInContainer shp.GroupItems, defs
        Else
            styleName = ""                          ' reset; don't carry the previous shape's tag
            On Error Resume Next
            styleName = shp.Tags("GroupStyle")
            On Error GoTo 0
            ' "Entry Body w/ Title" is assign-only: never auto-restyled (its title line is hand-tuned per entry)
            If styleName <> "" And styleName <> "Entry Body w/ Title" Then
                If defs.exists(styleName) Then
                    Dim styleDef As Object
                    Set styleDef = defs(styleName)
                    On Error Resume Next            ' one problem shape must not abort the whole deck sweep
                    If styleDef.exists("Dual") And styleDef("Dual") = True Then
                        ApplyDualStyleDefinition styleDef, shp
                    Else
                        ApplySingleStyleDefinition styleDef, shp
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next shp
End Sub

'====================================================
' (F) APPLY SINGLE STYLE DEFINITION
'====================================================
Sub ApplySingleStyleDefinition(styleDef As Object, shp As Shape)
    On Error Resume Next                 ' a missing/odd key degrades gracefully instead of aborting the shape
    With shp.TextFrame.TextRange.Font
        If styleDef.exists("FontName") Then .Name = styleDef("FontName")
        If styleDef.exists("FontSize") Then .Size = styleDef("FontSize")
        If styleDef.exists("Bold") Then .Bold = styleDef("Bold")
        If styleDef.exists("Italic") Then .Italic = styleDef("Italic")
        If styleDef.exists("Underline") Then .Underline = styleDef("Underline")
        If styleDef.exists("FontColor") Then .Color.RGB = styleDef("FontColor")
    End With
    With shp.TextFrame.TextRange.ParagraphFormat
        If styleDef.exists("Alignment") Then .Alignment = CInt(Val(styleDef("Alignment")))
        If styleDef.exists("SpaceBefore") Then .SpaceBefore = CDbl(styleDef("SpaceBefore"))
        If styleDef.exists("SpaceAfter") Then .SpaceAfter = CDbl(styleDef("SpaceAfter"))
    End With
    With shp.TextFrame
        If styleDef.exists("MarginLeft") Then .MarginLeft = styleDef("MarginLeft")
        If styleDef.exists("MarginRight") Then .MarginRight = styleDef("MarginRight")
        If styleDef.exists("MarginTop") Then .MarginTop = styleDef("MarginTop")
        If styleDef.exists("MarginBottom") Then .MarginBottom = styleDef("MarginBottom")
        If styleDef.exists("VerticalAnchor") Then .VerticalAnchor = styleDef("VerticalAnchor")
        If styleDef.exists("WordWrap") Then .WordWrap = CBool(styleDef("WordWrap"))
    End With
    If styleDef.exists("AutoSize") Then shp.TextFrame2.AutoSize = styleDef("AutoSize")

    If styleDef.exists("FillColor") Then
        With shp.Fill
            .ForeColor.RGB = styleDef("FillColor")
            .Transparency = styleDef("FillTransparency")
            .Solid
        End With
    End If
    If styleDef.exists("LineColor") Then
        With shp.line
            .ForeColor.RGB = styleDef("LineColor")
            .Weight = styleDef("LineWeight")
            .Visible = msoTrue
        End With
    End If

    ApplyTextShadow styleDef, shp
    On Error GoTo 0
End Sub

' Shared text-shadow apply for single + dual styles. Absent or false -> shadow off.
Private Sub ApplyTextShadow(ByVal styleDef As Object, ByVal shp As Shape)
    On Error Resume Next
    Dim wantShadow As Boolean
    wantShadow = False
    If styleDef.exists("TextShadow") Then
        wantShadow = (styleDef("TextShadow") = "True" Or styleDef("TextShadow") = "-1" Or CBool(styleDef("TextShadow")) = True)
    End If
    If wantShadow Then
        With shp.TextFrame2.TextRange.Font.Shadow
            .Visible = msoTrue
            .ForeColor.RGB = styleDef("TextShadowColor")
            .OffsetX = styleDef("TextShadowOffsetX")
            .OffsetY = styleDef("TextShadowOffsetY")
            .Transparency = styleDef("TextShadowTransparency")
        End With
    Else
        shp.TextFrame2.TextRange.Font.Shadow.Visible = msoFalse
    End If
End Sub

'====================================================
' (G) APPLY DUAL STYLE DEFINITION
' Applies dual formatting: Paragraph(1) gets Line1 formatting; Paragraphs(2+) get Line2 formatting.
'====================================================
Sub ApplyDualStyleDefinition(styleDef As Object, shp As Shape)
    On Error Resume Next                 ' degrade gracefully on missing/odd keys
    Dim tr As TextRange
    Dim paraCount As Long, p As Long
    Dim line1TR As TextRange, line2TR As TextRange
    Dim startOfLine2 As Long, lengthOfLine2 As Long
    Dim a1 As Long, a2 As Long

    Set tr = shp.TextFrame.TextRange
    paraCount = tr.Paragraphs.count
    If paraCount >= 2 Then
        Set line1TR = tr.Paragraphs(1)
        startOfLine2 = tr.Paragraphs(2).Start
        lengthOfLine2 = tr.Length - (startOfLine2 - 1)
        Set line2TR = tr.Characters(startOfLine2, lengthOfLine2)
    Else
        Set line1TR = tr
        Set line2TR = Nothing
    End If

    With line1TR.Font
        If styleDef.exists("Line1_FontName") Then .Name = styleDef("Line1_FontName")
        If styleDef.exists("Line1_FontSize") Then .Size = styleDef("Line1_FontSize")
        If styleDef.exists("Line1_Bold") Then .Bold = styleDef("Line1_Bold")
        If styleDef.exists("Line1_Italic") Then .Italic = styleDef("Line1_Italic")
        If styleDef.exists("Line1_Underline") Then .Underline = styleDef("Line1_Underline")
        If styleDef.exists("Line1_FontColor") Then .Color.RGB = styleDef("Line1_FontColor")
    End With
    a1 = CLng(Val(styleDef("Line1_Alignment")))
    If a1 < 1 Or a1 > 4 Then a1 = 2
    With line1TR.ParagraphFormat
        .Alignment = a1
        If styleDef.exists("Line1_SpaceBefore") Then .SpaceBefore = CDbl(styleDef("Line1_SpaceBefore"))
        If styleDef.exists("Line1_SpaceAfter") Then .SpaceAfter = CDbl(styleDef("Line1_SpaceAfter"))
    End With

    If Not line2TR Is Nothing Then
        With line2TR.Font
            If styleDef.exists("Line2_FontName") Then .Name = styleDef("Line2_FontName")
            If styleDef.exists("Line2_FontSize") Then .Size = styleDef("Line2_FontSize")
            If styleDef.exists("Line2_Bold") Then .Bold = styleDef("Line2_Bold")
            If styleDef.exists("Line2_Italic") Then .Italic = styleDef("Line2_Italic")
            If styleDef.exists("Line2_Underline") Then .Underline = styleDef("Line2_Underline")
            If styleDef.exists("Line2_FontColor") Then .Color.RGB = styleDef("Line2_FontColor")
        End With
        a2 = CLng(Val(styleDef("Line2_Alignment")))
        If a2 < 1 Or a2 > 4 Then a2 = 1
        ' Line2 font already spans paragraphs 2..N; apply its paragraph format to all of them too
        For p = 2 To paraCount
            With tr.Paragraphs(p).ParagraphFormat
                .Alignment = a2
                If styleDef.exists("Line2_SpaceBefore") Then .SpaceBefore = CDbl(styleDef("Line2_SpaceBefore"))
                If styleDef.exists("Line2_SpaceAfter") Then .SpaceAfter = CDbl(styleDef("Line2_SpaceAfter"))
            End With
        Next p
    End If

    With shp.TextFrame
        If styleDef.exists("MarginLeft") Then .MarginLeft = styleDef("MarginLeft")
        If styleDef.exists("MarginRight") Then .MarginRight = styleDef("MarginRight")
        If styleDef.exists("MarginTop") Then .MarginTop = styleDef("MarginTop")
        If styleDef.exists("MarginBottom") Then .MarginBottom = styleDef("MarginBottom")
        If styleDef.exists("VerticalAnchor") Then .VerticalAnchor = styleDef("VerticalAnchor")
        If styleDef.exists("WordWrap") Then .WordWrap = CBool(styleDef("WordWrap"))
    End With
    If styleDef.exists("AutoSize") Then shp.TextFrame2.AutoSize = styleDef("AutoSize")

    If styleDef.exists("FillColor") Then
        With shp.Fill
            .ForeColor.RGB = styleDef("FillColor")
            .Transparency = styleDef("FillTransparency")
            .Solid
        End With
    End If
    If styleDef.exists("LineColor") Then
        With shp.line
            .ForeColor.RGB = styleDef("LineColor")
            .Weight = styleDef("LineWeight")
            .Visible = msoTrue
        End With
    End If

    ApplyTextShadow styleDef, shp
    On Error GoTo 0
End Sub


