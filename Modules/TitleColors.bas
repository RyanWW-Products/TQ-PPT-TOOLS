Attribute VB_Name = "TitleColors"
Option Explicit

' ============================================================================
' TitleColors - palette catalog + persistent registry for the Timeline Title
' Manager. Mirrors EntryGroups.bas: INI text in a CustomDocumentProperty, one
' [section] per title, with a truncation read-back on save. Layout:
'   [_Settings]
'   Palette=Cool Professional
'   SingletonUnique=0
'   [Dr. Mayer]
'   Color=6697728
'   Manual=1
' A title with no Color key is "unassigned" (its date box shows default navy).
' Manual=1 means the user set the color by hand (Auto-Assign leaves it alone
' until it is Released back to the auto pool).
' ============================================================================
Public Const TITLE_PROP As String = "TimelineTitles"
Public Const TSETTINGS As String = "_Settings"
Public Const PALETTE_SIZE As Long = 12

' The stock date-box navy (matches CreateTimelineEntry's RGB(4,61,102)).
Public Function DefaultNavy() As Long
    DefaultNavy = RGB(4, 61, 102)
End Function

' One neutral slate shared by all single-occurrence titles (when they are not
' given their own colors). Adaptive text keeps it readable.
Public Function SingletonColor() As Long
    SingletonColor = RGB(96, 104, 120)
End Function

' --- palettes ---------------------------------------------------------------
Public Function PaletteNamesList() As Variant
    PaletteNamesList = Array("Cool Professional", "Jewel Tones", "Slate / Muted", "High-Contrast")
End Function

Public Function PaletteIndexOf(ByVal name As String) As Long
    Dim v As Variant, i As Long
    v = PaletteNamesList()
    For i = LBound(v) To UBound(v)
        If StrComp(CStr(v(i)), name, vbTextCompare) = 0 Then
            PaletteIndexOf = i
            Exit Function
        End If
    Next i
    PaletteIndexOf = 0            ' default to the first palette
End Function

' Each palette is 12 colors ordered COOL first, warm last (see CoolCount). All
' are chosen to read under EITHER white or black text (the manager picks the ink
' per color), so cool colors are prioritised without excluding warm ones.
Public Function TitlePaletteColor(ByVal pIdx As Long, ByVal cIdx As Long) As Long
    Dim p(0 To 11) As Long
    Select Case pIdx
        Case 1      ' Jewel Tones - rich, saturated (cool 0-7)
            p(0) = RGB(15, 82, 186)     ' Sapphire
            p(1) = RGB(0, 140, 90)      ' Emerald
            p(2) = RGB(120, 60, 170)    ' Amethyst
            p(3) = RGB(0, 128, 128)     ' Teal
            p(4) = RGB(40, 70, 180)     ' Cobalt
            p(5) = RGB(0, 150, 110)     ' Jade
            p(6) = RGB(90, 50, 160)     ' Violet
            p(7) = RGB(0, 110, 140)     ' Peacock
            p(8) = RGB(160, 30, 60)     ' Ruby (warm)
            p(9) = RGB(200, 140, 0)     ' Amber (warm)
            p(10) = RGB(130, 20, 50)    ' Garnet (warm)
            p(11) = RGB(190, 100, 20)   ' Topaz (warm)
        Case 2      ' Slate / Muted - desaturated (cool 0-8)
            p(0) = RGB(70, 85, 100)     ' Slate
            p(1) = RGB(95, 120, 150)    ' Dusty Blue
            p(2) = RGB(110, 130, 110)   ' Sage
            p(3) = RGB(80, 120, 120)    ' Muted Teal
            p(4) = RGB(90, 105, 130)    ' Blue Gray
            p(5) = RGB(100, 120, 90)    ' Moss
            p(6) = RGB(75, 90, 110)     ' Storm
            p(7) = RGB(120, 140, 160)   ' Fog Blue
            p(8) = RGB(110, 115, 120)   ' Pewter
            p(9) = RGB(150, 110, 95)    ' Clay (warm)
            p(10) = RGB(150, 110, 130)  ' Mauve (warm)
            p(11) = RGB(140, 120, 100)  ' Taupe (warm)
        Case 3      ' High-Contrast - bold, distinct (cool 0-7)
            p(0) = RGB(20, 30, 60)      ' Midnight
            p(1) = RGB(0, 110, 110)     ' Strong Teal
            p(2) = RGB(20, 70, 200)     ' Vivid Blue
            p(3) = RGB(0, 110, 60)      ' Deep Green
            p(4) = RGB(100, 30, 160)    ' Purple
            p(5) = RGB(0, 120, 140)     ' Dark Cyan
            p(6) = RGB(45, 45, 140)     ' Indigo
            p(7) = RGB(30, 130, 70)     ' Green
            p(8) = RGB(190, 30, 40)     ' Red (warm)
            p(9) = RGB(210, 110, 20)    ' Orange (warm)
            p(10) = RGB(170, 30, 110)   ' Magenta (warm)
            p(11) = RGB(120, 70, 40)    ' Brown (warm)
        Case Else   ' 0 = Cool Professional - muted pro (cool 0-9)
            p(0) = RGB(4, 61, 102)      ' Navy
            p(1) = RGB(0, 128, 128)     ' Teal
            p(2) = RGB(30, 90, 168)     ' Royal Blue
            p(3) = RGB(72, 84, 140)     ' Slate Blue
            p(4) = RGB(34, 110, 80)     ' Forest Green
            p(5) = RGB(20, 130, 150)    ' Deep Cyan
            p(6) = RGB(63, 60, 140)     ' Indigo
            p(7) = RGB(70, 110, 150)    ' Steel Blue
            p(8) = RGB(46, 139, 120)    ' Sea Green
            p(9) = RGB(110, 70, 140)    ' Plum
            p(10) = RGB(130, 45, 60)    ' Burgundy (warm)
            p(11) = RGB(150, 80, 40)    ' Rust (warm)
    End Select
    If cIdx < 0 Then cIdx = 0
    TitlePaletteColor = p(cIdx Mod PALETTE_SIZE)
End Function

Public Function PaletteColorByName(ByVal name As String, ByVal cIdx As Long) As Long
    PaletteColorByName = TitlePaletteColor(PaletteIndexOf(name), cIdx)
End Function

' How many of the 12 swatches (from index 0) are the "cool" block. Auto-Assign
' hands these out before the warm tail.
Public Function CoolCount(ByVal pIdx As Long) As Long
    Select Case pIdx
        Case 1: CoolCount = 8       ' Jewel
        Case 2: CoolCount = 9       ' Slate
        Case 3: CoolCount = 8       ' High-Contrast
        Case Else: CoolCount = 10   ' Cool Professional
    End Select
End Function

' --- load / save (verbatim pattern from EntryGroups.bas) --------------------
Public Function LoadTitleReg() As Object
    Dim reg As Object, cur As Object
    Dim ini As String, lines() As String, i As Long, ln As String, secName As String
    Dim parts() As String
    Set reg = CreateObject("Scripting.Dictionary")
    reg.CompareMode = vbTextCompare

    On Error Resume Next
    ini = ActivePresentation.CustomDocumentProperties(TITLE_PROP).Value
    On Error GoTo 0

    If ini <> "" Then
        lines = Split(ini, vbCrLf)
        For i = 0 To UBound(lines)
            ln = Trim(lines(i))
            If ln <> "" And Left(ln, 1) <> ";" Then
                If Left(ln, 1) = "[" And Right(ln, 1) = "]" Then
                    secName = Mid(ln, 2, Len(ln) - 2)
                    Set cur = CreateObject("Scripting.Dictionary")
                    cur.CompareMode = vbTextCompare
                    Set reg(secName) = cur
                ElseIf InStr(ln, "=") > 0 Then
                    parts = Split(ln, "=", 2)
                    If Not cur Is Nothing Then cur(Trim(parts(0))) = Trim(parts(1))
                End If
            End If
        Next i
    End If
    Set LoadTitleReg = reg
End Function

Public Sub SaveTitleReg(reg As Object)
    Dim ini As String, secName As Variant, k As Variant, sec As Object
    Dim chk As String
    For Each secName In reg.Keys
        ini = ini & "[" & secName & "]" & vbCrLf
        Set sec = reg(secName)
        For Each k In sec.Keys
            ini = ini & k & "=" & CleanValue(CStr(sec(k))) & vbCrLf
        Next k
        ini = ini & vbCrLf
    Next secName

    On Error Resume Next
    ActivePresentation.CustomDocumentProperties(TITLE_PROP).Value = ini
    If Err.Number <> 0 Then
        Err.Clear
        ActivePresentation.CustomDocumentProperties.Add Name:=TITLE_PROP, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=ini
    End If
    On Error GoTo 0

    On Error Resume Next
    chk = ActivePresentation.CustomDocumentProperties(TITLE_PROP).Value
    On Error GoTo 0
    If Len(Replace(chk, vbCr, "")) < Len(Replace(ini, vbCr, "")) Then
        MsgBox "The title-color registry may not have fully saved (" & Len(chk) & " of " & _
               Len(ini) & " characters stored). You likely have more titles than this " & _
               "presentation's property storage allows.", vbExclamation, "Title Manager"
    End If
End Sub

Private Function CleanValue(ByVal v As String) As String
    CleanValue = Replace(Replace(v, vbCr, " "), vbLf, " ")
End Function

' --- settings ---------------------------------------------------------------
Public Sub EnsureTitleSettings(reg As Object)
    Dim s As Object
    If Not reg.Exists(TSETTINGS) Then
        Set s = CreateObject("Scripting.Dictionary")
        s.CompareMode = vbTextCompare
        Set reg(TSETTINGS) = s
    End If
    Set s = reg(TSETTINGS)
    If Not s.Exists("Palette") Then s("Palette") = "Cool Professional"
    If Not s.Exists("SingletonUnique") Then s("SingletonUnique") = "0"
End Sub

Public Function GetPalette(reg As Object) As String
    EnsureTitleSettings reg
    GetPalette = CStr(reg(TSETTINGS)("Palette"))
End Function

Public Sub SetPalette(reg As Object, ByVal name As String)
    EnsureTitleSettings reg
    reg(TSETTINGS)("Palette") = name
End Sub

Public Function GetSingletonUnique(reg As Object) As Boolean
    EnsureTitleSettings reg
    GetSingletonUnique = (Val(reg(TSETTINGS)("SingletonUnique")) <> 0)
End Function

Public Sub SetSingletonUnique(reg As Object, ByVal isOn As Boolean)
    EnsureTitleSettings reg
    reg(TSETTINGS)("SingletonUnique") = IIf(isOn, "1", "0")
End Sub

' --- title queries / mutations ----------------------------------------------
Public Function SanitizeTitleName(ByVal s As String) As String
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "=", "")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    SanitizeTitleName = Trim(s)
End Function

Public Function TitleExists(reg As Object, ByVal key As String) As Boolean
    If key = TSETTINGS Then Exit Function
    TitleExists = reg.Exists(key)
End Function

' Create an empty (unassigned) section so the title is tracked/persisted.
Public Sub EnsureTitle(reg As Object, ByVal key As String)
    Dim g As Object
    key = SanitizeTitleName(key)
    If key = "" Or key = TSETTINGS Then Exit Sub
    If reg.Exists(key) Then Exit Sub
    Set g = CreateObject("Scripting.Dictionary")
    g.CompareMode = vbTextCompare
    Set reg(key) = g
End Sub

Public Function HasColor(reg As Object, ByVal key As String) As Boolean
    If Not reg.Exists(key) Then Exit Function
    HasColor = reg(key).Exists("Color")
End Function

Public Function GetTitleColor(reg As Object, ByVal key As String) As Long
    If reg.Exists(key) Then
        If reg(key).Exists("Color") Then
            GetTitleColor = CLng(Val(reg(key)("Color")))
            Exit Function
        End If
    End If
    GetTitleColor = DefaultNavy()
End Function

Public Sub SetTitleColor(reg As Object, ByVal key As String, ByVal color As Long)
    EnsureTitle reg, key
    key = SanitizeTitleName(key)
    If reg.Exists(key) Then reg(key)("Color") = CStr(color)
End Sub

' Remove the color (back to unassigned/navy) and clear the manual flag.
Public Sub ClearTitleColor(reg As Object, ByVal key As String)
    If Not reg.Exists(key) Then Exit Sub
    On Error Resume Next
    reg(key).Remove "Color"
    On Error GoTo 0
    reg(key)("Manual") = "0"
End Sub

Public Function GetTitleManual(reg As Object, ByVal key As String) As Boolean
    If reg.Exists(key) Then
        If reg(key).Exists("Manual") Then GetTitleManual = (Val(reg(key)("Manual")) <> 0)
    End If
End Function

Public Sub SetTitleManual(reg As Object, ByVal key As String, ByVal isManual As Boolean)
    EnsureTitle reg, key
    key = SanitizeTitleName(key)
    If reg.Exists(key) Then reg(key)("Manual") = IIf(isManual, "1", "0")
End Sub
