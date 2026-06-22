Attribute VB_Name = "EntryGroups"
Option Explicit

' ============================================================================
' EntryGroups - the named-group registry for the Calendar Entry Manager.
' ============================================================================
' Stored presentation-wide as INI text in a CustomDocumentProperty named
' "EntryGroups" (the same proven pattern as GroupStyles.bas, including the
' CR-stripped truncation read-back). Group MEMBERSHIP is NOT stored here - it
' lives on each entry shape's "EntryGroup" tag - so the registry stays tiny
' regardless of how many entries exist. Layout:
'   [_Settings]
'   LegendOn=1
'   LegendFontSize=10
'   [Dr. A]
'   Color=12611584
'   Order=1
Public Const ENTRY_PROP As String = "EntryGroups"
Public Const SETTINGS_SECTION As String = "_Settings"
Public Const DEFAULT_GROUP As String = "Default"
Public Const PALETTE_COUNT As Long = 12

' --- palette ----------------------------------------------------------------
' A 12-swatch palette extending the legacy 4-color cycle. 0-based index.
Public Function PaletteColor(ByVal idx As Long) As Long
    Dim p(0 To 11) As Long
    p(0) = RGB(173, 216, 230)   ' Light Blue
    p(1) = RGB(144, 238, 144)   ' Light Green
    p(2) = RGB(255, 165, 0)     ' Orange
    p(3) = RGB(255, 182, 193)   ' Light Pink
    p(4) = RGB(221, 160, 221)   ' Plum
    p(5) = RGB(255, 255, 153)   ' Light Yellow
    p(6) = RGB(176, 224, 230)   ' Powder Blue
    p(7) = RGB(255, 200, 124)   ' Light Orange
    p(8) = RGB(200, 200, 200)   ' Grey
    p(9) = RGB(152, 251, 152)   ' Pale Green
    p(10) = RGB(240, 128, 128)  ' Light Coral
    p(11) = RGB(175, 238, 238)  ' Pale Turquoise
    If idx < 0 Then idx = 0
    PaletteColor = p(idx Mod PALETTE_COUNT)
End Function

Public Function PaletteNames() As Variant
    PaletteNames = Array("Light Blue", "Light Green", "Orange", "Light Pink", _
        "Plum", "Light Yellow", "Powder Blue", "Light Orange", _
        "Grey", "Pale Green", "Light Coral", "Pale Turquoise")
End Function

' --- load / save ------------------------------------------------------------
Public Function LoadRegistry() As Object
    Dim reg As Object, cur As Object
    Dim ini As String, lines() As String, i As Long, ln As String, secName As String
    Dim parts() As String
    Set reg = CreateObject("Scripting.Dictionary")
    reg.CompareMode = vbTextCompare

    On Error Resume Next
    ini = ActivePresentation.CustomDocumentProperties(ENTRY_PROP).Value
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
    Set LoadRegistry = reg
End Function

Public Sub SaveRegistry(reg As Object)
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
    ActivePresentation.CustomDocumentProperties(ENTRY_PROP).Value = ini
    If Err.Number <> 0 Then
        Err.Clear
        ActivePresentation.CustomDocumentProperties.Add Name:=ENTRY_PROP, _
            LinkToContent:=False, Type:=msoPropertyTypeString, Value:=ini
    End If
    On Error GoTo 0

    On Error Resume Next
    chk = ActivePresentation.CustomDocumentProperties(ENTRY_PROP).Value
    On Error GoTo 0
    If Len(Replace(chk, vbCr, "")) < Len(Replace(ini, vbCr, "")) Then
        MsgBox "The entry-group registry may not have fully saved (" & Len(chk) & " of " & _
               Len(ini) & " characters stored). You likely have more groups than this " & _
               "presentation's property storage allows.", vbExclamation, "Entry Manager"
    End If
End Sub

' Keep a serialized value on one line so it can't fabricate INI structure.
Private Function CleanValue(ByVal v As String) As String
    CleanValue = Replace(Replace(v, vbCr, " "), vbLf, " ")
End Function

' --- settings ---------------------------------------------------------------
Public Sub EnsureSettings(reg As Object)
    Dim s As Object
    If Not reg.Exists(SETTINGS_SECTION) Then
        Set s = CreateObject("Scripting.Dictionary")
        s.CompareMode = vbTextCompare
        Set reg(SETTINGS_SECTION) = s
    End If
    Set s = reg(SETTINGS_SECTION)
    If Not s.Exists("LegendOn") Then s("LegendOn") = "1"
    If Not s.Exists("LegendFontSize") Then s("LegendFontSize") = "10"
End Sub

Public Sub EnsureDefaultGroup(reg As Object)
    EnsureSettings reg
    If GroupCount(reg) = 0 Then AddGroup reg, DEFAULT_GROUP
End Sub

Public Function LegendOn(reg As Object) As Boolean
    EnsureSettings reg
    LegendOn = (Val(reg(SETTINGS_SECTION)("LegendOn")) <> 0)
End Function

Public Function LegendFontSize(reg As Object) As Single
    EnsureSettings reg
    LegendFontSize = Val(reg(SETTINGS_SECTION)("LegendFontSize"))
    If LegendFontSize < 6 Then LegendFontSize = 10
End Function

Public Sub SetLegendOn(reg As Object, ByVal isOn As Boolean)
    EnsureSettings reg
    reg(SETTINGS_SECTION)("LegendOn") = IIf(isOn, "1", "0")
End Sub

Public Sub SetLegendFontSize(reg As Object, ByVal sz As Single)
    EnsureSettings reg
    If sz < 6 Then sz = 6
    If sz > 36 Then sz = 36
    reg(SETTINGS_SECTION)("LegendFontSize") = CStr(sz)
End Sub

' --- group queries ----------------------------------------------------------
Public Function GroupCount(reg As Object) As Long
    Dim k As Variant, n As Long
    For Each k In reg.Keys
        If CStr(k) <> SETTINGS_SECTION Then n = n + 1
    Next k
    GroupCount = n
End Function

Public Function GroupExists(reg As Object, ByVal name As String) As Boolean
    If name = SETTINGS_SECTION Then Exit Function
    GroupExists = reg.Exists(name)
End Function

Public Function GroupOrder(reg As Object, ByVal name As String) As Long
    If reg.Exists(name) Then
        If reg(name).Exists("Order") Then GroupOrder = CLng(Val(reg(name)("Order")))
    End If
End Function

Public Function GroupColor(reg As Object, ByVal name As String) As Long
    If reg.Exists(name) Then
        If reg(name).Exists("Color") Then GroupColor = CLng(Val(reg(name)("Color")))
    End If
End Function

' Group names sorted by their Order key (the _Settings section is excluded).
Public Function GroupNamesInOrder(reg As Object) As Variant
    Dim names() As String, ords() As Long, n As Long, k As Variant
    Dim i As Long, j As Long, tn As String, tordr As Long
    n = GroupCount(reg)
    If n = 0 Then GroupNamesInOrder = Array(): Exit Function
    ReDim names(0 To n - 1)
    ReDim ords(0 To n - 1)
    i = 0
    For Each k In reg.Keys
        If CStr(k) <> SETTINGS_SECTION Then
            names(i) = CStr(k)
            ords(i) = GroupOrder(reg, CStr(k))
            i = i + 1
        End If
    Next k
    For i = 1 To n - 1
        tn = names(i): tordr = ords(i): j = i - 1
        Do While j >= 0
            If ords(j) <= tordr Then Exit Do
            names(j + 1) = names(j): ords(j + 1) = ords(j): j = j - 1
        Loop
        names(j + 1) = tn: ords(j + 1) = tordr
    Next i
    GroupNamesInOrder = names
End Function

' --- group mutations (registry only; shape side-effects live in EntryManager) ---
Public Function SanitizeGroupName(ByVal s As String) As String
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "=", "")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    SanitizeGroupName = Trim(s)
End Function

' Add a group with the next palette color. Returns the (sanitized) name, or "" on failure.
Public Function AddGroup(reg As Object, ByVal name As String) As String
    Dim g As Object, ord As Long
    name = SanitizeGroupName(name)
    If name = "" Then Exit Function
    If name = SETTINGS_SECTION Then Exit Function
    If reg.Exists(name) Then Exit Function
    Set g = CreateObject("Scripting.Dictionary")
    g.CompareMode = vbTextCompare
    ord = GroupCount(reg)                 ' 0-based slot = palette index
    g("Color") = CStr(PaletteColor(ord))
    g("Order") = CStr(ord + 1)
    Set reg(name) = g
    AddGroup = name
End Function

Public Sub SetGroupColor(reg As Object, ByVal name As String, ByVal color As Long)
    If reg.Exists(name) Then reg(name)("Color") = CStr(color)
End Sub

Public Sub SetGroupOrder(reg As Object, ByVal name As String, ByVal ord As Long)
    If reg.Exists(name) Then reg(name)("Order") = CStr(ord)
End Sub

' Move the registry section key (dictionaries can't rename a key in place).
Public Sub RenameGroupKey(reg As Object, ByVal oldName As String, ByVal newName As String)
    Dim g As Object
    If Not reg.Exists(oldName) Then Exit Sub
    Set g = reg(oldName)
    reg.Remove oldName
    Set reg(newName) = g
End Sub

Public Sub RemoveGroupKey(reg As Object, ByVal name As String)
    If reg.Exists(name) Then reg.Remove name
End Sub
