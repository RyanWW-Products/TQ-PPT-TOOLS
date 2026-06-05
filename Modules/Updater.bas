Attribute VB_Name = "Updater"
Option Explicit

' ============================================================================
' UPDATER  -  GitHub-backed self-update for the Trial Quest add-in
' ============================================================================
' Ribbon entry points:
'     CheckForUpdates(control)  -> "Check for Updates" button
'     ShowAbout(control)        -> "About" button
'     Auto_Open()               -> optional silent check on add-in load
'
' How it works
'   * version.json (repo root) declares the latest addinVersion + assetsVersion.
'   * manifest.json (repo root) lists every template asset + its byte size.
'   * Downloads use the GitHub *contents* API with a read-only token, so the
'     repository can be PRIVATE. The token lives in the registry (written by the
'     installer), never in this file or in git.
'   * Template assets are not locked, so they are overwritten in place at once.
'   * The .ppam IS locked while PowerPoint runs, so the new copy is staged and a
'     tiny PowerShell "swapper" copies it into place after PowerPoint closes,
'     then relaunches PowerPoint. No PowerPoint re-registration is needed because
'     the installed file name never changes.
'   * The installed version is NOT baked into this code. It lives in the registry
'     (InstalledVersion), written by the installer on first install and by the
'     swapper after each update. So a version bump is just version.json + push --
'     no editing a constant and no re-importing this module.
'
' SET THESE TWO BEFORE BUILDING (must match installer.iss + the repo):
Private Const GH_OWNER     As String = "RyanWW-Products"
Private Const GH_REPO      As String = "TQ-PPT-TOOLS"
' The branch is chosen by the install CHANNEL (registry "Channel"):
'   stable -> main branch (live users) ;  beta -> beta branch (you / testers).
Private Const BRANCH_STABLE As String = "main"
Private Const BRANCH_BETA   As String = "beta"
' ============================================================================

' --- Fixed install layout (matches the installer) ---------------------------
Private Const ADDIN_SUBDIR   As String = "Trial Ex Addin\"   ' under %APPDATA%\Microsoft\AddIns\
Private Const ADDIN_FILENAME As String = "TrialQuest.ppam"
Private Const STAGING_SUBDIR As String = "TrialQuestUpdate\" ' under %TEMP% (NOT the add-in dir, so the swapper can clean up without locking itself)

' --- Registry (HKCU\Software\TrialQuest\Addin\) ------------------------------
Private Const REG_ROOT As String = "HKEY_CURRENT_USER\Software\TrialQuest\Addin\"
Private Const REG_TOKEN  As String = "GitHubToken"
Private Const REG_ASSETS As String = "AssetsVersion"
Private Const REG_INSTALLED As String = "InstalledVersion"
Private Const REG_CHANNEL As String = "Channel"
Private Const REG_AUTOCHK As String = "AutoCheck"

' ============================================================================
' RIBBON ENTRY POINTS
' ============================================================================

Public Sub CheckForUpdates(control As IRibbonControl)
    DoUpdateCheck True
End Sub

Public Sub ShowAbout(control As IRibbonControl)
    Dim tok As String, msg As String
    tok = GetToken()
    msg = "Trial Quest PowerPoint Add-in" & vbCrLf & String(34, "-") & vbCrLf & _
          "Installed version : " & InstalledVersion() & vbCrLf & _
          "Assets version    : " & RegGet(REG_ASSETS, "0") & vbCrLf & _
          "Channel           : " & ChannelName() & vbCrLf & _
          "Update source     : " & GH_OWNER & "/" & GH_REPO & " (" & BranchRef() & ")" & vbCrLf & _
          "Update token      : " & IIf(tok = "", "NOT configured", "configured") & vbCrLf & _
          "Install folder    : " & AddInDir()
    If tok = "" Then
        msg = msg & vbCrLf & vbCrLf & "No update token is configured. Set one now?"
        If MsgBox(msg, vbExclamation + vbYesNo, "About Trial Quest") = vbYes Then PromptForToken
    Else
        MsgBox msg, vbInformation, "About Trial Quest"
    End If
End Sub

Public Sub Auto_Open()
    ' Optional, off by default. Turn on with:
    '   HKCU\Software\TrialQuest\Addin\AutoCheck = "1"
    On Error Resume Next
    If RegGet(REG_AUTOCHK, "0") = "1" Then DoUpdateCheck False
End Sub

' ============================================================================
' CORE
' ============================================================================

Private Sub DoUpdateCheck(ByVal interactive As Boolean)
    Dim tok As String, vj As String
    Dim latestVer As String, latestAssets As Long, localAssets As Long
    Dim needAddin As Boolean, needAssets As Boolean
    Dim lastErr As String

    tok = GetToken()
    If tok = "" Then
        If interactive Then
            If MsgBox("No update token is configured on this machine." & vbCrLf & vbCrLf & _
                      "Re-run the installer, or set a token now?", _
                      vbExclamation + vbYesNo, "Check for Updates") = vbYes Then PromptForToken
        End If
        Exit Sub
    End If

    vj = HttpGetText("version.json", lastErr)
    If vj = "" Then
        If interactive Then MsgBox "Could not reach the update server." & vbCrLf & vbCrLf & lastErr, _
                                   vbExclamation, "Check for Updates"
        Exit Sub
    End If

    latestVer = JsonStr(vj, "addinVersion")
    latestAssets = JsonNum(vj, "assetsVersion")
    localAssets = CLng(Val(RegGet(REG_ASSETS, "0")))

    needAddin = (CompareVersions(latestVer, InstalledVersion()) > 0)
    needAssets = (latestAssets > localAssets)

    If Not needAddin And Not needAssets Then
        If interactive Then MsgBox "You are up to date." & vbCrLf & vbCrLf & _
                                   "Add-in version: " & InstalledVersion() & vbCrLf & _
                                   "Assets version: " & localAssets, vbInformation, "Check for Updates"
        Exit Sub
    End If

    Dim prompt As String
    prompt = "An update is available:" & vbCrLf & vbCrLf
    If needAddin Then prompt = prompt & "  - Add-in:  " & InstalledVersion() & "  ->  " & latestVer & vbCrLf
    If needAssets Then prompt = prompt & "  - Assets:  v" & localAssets & "  ->  v" & latestAssets & vbCrLf
    prompt = prompt & vbCrLf & "Download and install now?"
    If MsgBox(prompt, vbQuestion + vbYesNo, "Check for Updates") <> vbYes Then Exit Sub

    ' 1) Assets first (applied in place immediately; not locked) ---------------
    If needAssets Then
        If Not SyncAssets(latestAssets, lastErr) Then
            MsgBox "Asset update failed:" & vbCrLf & vbCrLf & lastErr, vbExclamation, "Check for Updates"
            Exit Sub
        End If
    End If

    ' 2) Add-in (staged; swapped after PowerPoint closes) ----------------------
    If needAddin Then
        If StageAddinAndSwap(latestVer, lastErr) Then
            MsgBox "Update downloaded." & vbCrLf & vbCrLf & _
                   "Please SAVE your work and CLOSE all PowerPoint windows." & vbCrLf & _
                   "The new version will install automatically and PowerPoint will reopen.", _
                   vbInformation, "Check for Updates"
        Else
            MsgBox "Add-in update failed:" & vbCrLf & vbCrLf & lastErr, vbExclamation, "Check for Updates"
        End If
    Else
        MsgBox "Template assets updated to v" & latestAssets & ".", vbInformation, "Check for Updates"
    End If
End Sub

' ----------------------------------------------------------------------------
' Download every asset whose local copy is missing or a different size.
' ----------------------------------------------------------------------------
Private Function SyncAssets(ByVal newVersion As Long, ByRef errOut As String) As Boolean
    Dim mf As String, root As String
    Dim paths() As String, sizes() As String, n As Long, i As Long
    Dim target As String, downloaded As Long

    mf = HttpGetText("manifest.json", errOut)
    If mf = "" Then SyncAssets = False: Exit Function

    n = ParseManifest(mf, paths, sizes)
    If n = 0 Then errOut = "Manifest contained no files.": SyncAssets = False: Exit Function

    root = AddInsRoot()
    For i = 0 To n - 1
        target = root & Replace(paths(i), "/", "\")
        If (Dir(target) = "") Or (FileLenSafe(target) <> CLng(Val(sizes(i)))) Then
            EnsureFolder Left(target, InStrRev(target, "\"))
            If Not HttpDownloadFile("assets/" & paths(i), target, errOut) Then
                errOut = "Failed downloading " & paths(i) & vbCrLf & errOut
                SyncAssets = False: Exit Function
            End If
            downloaded = downloaded + 1
        End If
        DoEvents
    Next i

    RegSet REG_ASSETS, CStr(newVersion)
    Debug.Print "SyncAssets: " & downloaded & " of " & n & " file(s) downloaded; assets now v" & newVersion
    SyncAssets = True
End Function

' ----------------------------------------------------------------------------
' Download the new .ppam to a staging folder and launch the swapper.
' ----------------------------------------------------------------------------
Private Function StageAddinAndSwap(ByVal newVersion As String, ByRef errOut As String) As Boolean
    Dim staging As String, staged As String, target As String, script As String, pptExe As String

    staging = Environ$("TEMP") & "\" & STAGING_SUBDIR
    EnsureFolder staging
    staged = staging & ADDIN_FILENAME
    target = AddInDir() & ADDIN_FILENAME
    pptExe = Application.Path & "\POWERPNT.EXE"

    If Not HttpDownloadFile("dist/" & ADDIN_FILENAME, staged, errOut) Then
        StageAddinAndSwap = False: Exit Function
    End If

    script = staging & "apply_update.ps1"
    WriteSwapperScript script, staged, target, pptExe, newVersion

    ' Launch detached & hidden so it survives PowerPoint closing.
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & script & """", 0, False

    StageAddinAndSwap = True
End Function

Private Sub WriteSwapperScript(ByVal scriptPath As String, ByVal staged As String, _
                               ByVal target As String, ByVal pptExe As String, ByVal newVersion As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(scriptPath, True)
    ts.WriteLine "$ErrorActionPreference = 'SilentlyContinue'"
    ts.WriteLine "$staged = '" & staged & "'"
    ts.WriteLine "$target = '" & target & "'"
    ts.WriteLine "$ppt    = '" & pptExe & "'"
    ts.WriteLine "while (Get-Process POWERPNT -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 1 }"
    ts.WriteLine "Start-Sleep -Seconds 1"
    ts.WriteLine "$copied = $false"
    ts.WriteLine "for ($i = 0; $i -lt 30; $i++) {"
    ts.WriteLine "  try { Copy-Item -LiteralPath $staged -Destination $target -Force; $copied = $true; break }"
    ts.WriteLine "  catch { Start-Sleep -Seconds 1 }"
    ts.WriteLine "}"
    ts.WriteLine "if ($copied) {"
    ts.WriteLine "  New-Item -Path 'HKCU:\Software\TrialQuest\Addin' -Force | Out-Null"
    ts.WriteLine "  Set-ItemProperty -Path 'HKCU:\Software\TrialQuest\Addin' -Name 'InstalledVersion' -Value '" & newVersion & "'"
    ts.WriteLine "}"
    ts.WriteLine "Remove-Item -LiteralPath $staged -Force -ErrorAction SilentlyContinue"
    ts.WriteLine "if (Test-Path $ppt) { Start-Process -FilePath $ppt } else { Start-Process powerpnt }"
    ts.Close
End Sub

' ============================================================================
' HTTP (GitHub contents API, supports private repos via token)
' ============================================================================

Private Function HttpGetText(ByVal repoPath As String, ByRef errOut As String) As String
    Dim http As Object
    On Error GoTo fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", ContentsUrl(repoPath), False
    SetCommonHeaders http
    http.send
    If http.Status = 200 Then
        HttpGetText = http.responseText
    Else
        errOut = "HTTP " & http.Status & " " & http.statusText
        HttpGetText = ""
    End If
    Exit Function
fail:
    errOut = "Request error: " & Err.Description
    HttpGetText = ""
End Function

Private Function HttpDownloadFile(ByVal repoPath As String, ByVal destPath As String, _
                                  ByRef errOut As String) As Boolean
    Dim http As Object, stream As Object
    On Error GoTo fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", ContentsUrl(repoPath), False
    SetCommonHeaders http
    http.send
    If http.Status <> 200 Then
        errOut = "HTTP " & http.Status & " " & http.statusText
        HttpDownloadFile = False: Exit Function
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                 ' binary
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile destPath, 2   ' overwrite
    stream.Close
    HttpDownloadFile = True
    Exit Function
fail:
    errOut = "Download error: " & Err.Description
    HttpDownloadFile = False
End Function

Private Sub SetCommonHeaders(ByVal http As Object)
    http.setRequestHeader "Authorization", "Bearer " & GetToken()
    http.setRequestHeader "Accept", "application/vnd.github.raw"
    http.setRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    http.setRequestHeader "User-Agent", "TrialQuest-Addin/" & InstalledVersion()
End Sub

Private Function ContentsUrl(ByVal repoPath As String) As String
    ContentsUrl = "https://api.github.com/repos/" & GH_OWNER & "/" & GH_REPO & _
                  "/contents/" & UrlEncodePath(repoPath) & "?ref=" & BranchRef()
End Function

' The install channel ("stable" or "beta") and the branch it maps to.
Private Function ChannelName() As String
    ChannelName = LCase(RegGet(REG_CHANNEL, "stable"))
End Function
Private Function BranchRef() As String
    If ChannelName() = "beta" Then BranchRef = BRANCH_BETA Else BranchRef = BRANCH_STABLE
End Function

' Percent-encode a repo path for the GitHub URL. Preserves "/" and the RFC 3986
' unreserved set; encodes everything else (spaces, &, +, parentheses, ...).
' Assumes ASCII asset names (true for this project).
Private Function UrlEncodePath(ByVal p As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(p)
        ch = Mid(p, i, 1)
        If ch Like "[A-Za-z0-9]" Or ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Or ch = "/" Then
            out = out & ch
        Else
            out = out & "%" & Right("0" & Hex(Asc(ch)), 2)
        End If
    Next i
    UrlEncodePath = out
End Function

' ============================================================================
' JSON helpers (tiny, regex-based; sufficient for our flat version/manifest)
' ============================================================================

Private Function JsonStr(ByVal json As String, ByVal key As String) As String
    JsonStr = RegexFirst(json, """" & key & """\s*:\s*""([^""]*)""")
End Function

Private Function JsonNum(ByVal json As String, ByVal key As String) As Long
    Dim s As String
    s = RegexFirst(json, """" & key & """\s*:\s*(-?\d+)")
    If s = "" Then JsonNum = 0 Else JsonNum = CLng(s)
End Function

' Parse manifest.json "files":[{"path":"..","size":N}, ...] -> parallel arrays.
Private Function ParseManifest(ByVal json As String, ByRef paths() As String, _
                               ByRef sizes() As String) As Long
    Dim re As Object, matches As Object, m As Object, i As Long
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = """path""\s*:\s*""([^""]*)""\s*,\s*""size""\s*:\s*(\d+)"
    Set matches = re.Execute(json)
    If matches.count = 0 Then ParseManifest = 0: Exit Function
    ReDim paths(matches.count - 1)
    ReDim sizes(matches.count - 1)
    For Each m In matches
        paths(i) = m.SubMatches(0)
        sizes(i) = m.SubMatches(1)
        i = i + 1
    Next m
    ParseManifest = matches.count
End Function

Private Function RegexFirst(ByVal s As String, ByVal pattern As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = pattern
    Set matches = re.Execute(s)
    If matches.count > 0 Then RegexFirst = matches(0).SubMatches(0) Else RegexFirst = ""
End Function

' ============================================================================
' Version comparison ("5.4.10" > "5.4.9")
' ============================================================================
Private Function CompareVersions(ByVal a As String, ByVal b As String) As Integer
    Dim pa() As String, pb() As String, i As Long, na As Long, nb As Long, n As Long
    pa = Split(a, ".")
    pb = Split(b, ".")
    n = UBound(pa)
    If UBound(pb) > n Then n = UBound(pb)
    For i = 0 To n
        na = PartAt(pa, i)
        nb = PartAt(pb, i)
        If na > nb Then CompareVersions = 1: Exit Function
        If na < nb Then CompareVersions = -1: Exit Function
    Next i
    CompareVersions = 0
End Function

Private Function PartAt(ByRef parts() As String, ByVal idx As Long) As Long
    On Error Resume Next
    If idx <= UBound(parts) Then PartAt = CLng(Val(parts(idx))) Else PartAt = 0
End Function

' ============================================================================
' Registry / paths / filesystem helpers
' ============================================================================

Private Function GetToken() As String
    GetToken = RegGet(REG_TOKEN, "")
End Function

' The locally installed add-in version. Set by the installer (first install) and
' by the swapper after each update; defaults low so an update is offered if unset.
Private Function InstalledVersion() As String
    InstalledVersion = RegGet(REG_INSTALLED, "0.0.0")
End Function

Private Sub PromptForToken()
    Dim t As String
    t = InputBox("Paste the GitHub read-only access token for updates:", "Set Update Token")
    If t <> "" Then
        RegSet REG_TOKEN, t
        MsgBox "Token saved. Try 'Check for Updates' again.", vbInformation, "Set Update Token"
    End If
End Sub

Private Function RegGet(ByVal name As String, ByVal dflt As String) As String
    Dim sh As Object
    On Error GoTo useDefault
    Set sh = CreateObject("WScript.Shell")
    RegGet = sh.RegRead(REG_ROOT & name)
    Exit Function
useDefault:
    RegGet = dflt
End Function

Private Sub RegSet(ByVal name As String, ByVal value As String)
    Dim sh As Object
    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    sh.RegWrite REG_ROOT & name, value, "REG_SZ"
End Sub

Private Function AddInsRoot() As String
    AddInsRoot = Environ$("APPDATA") & "\Microsoft\AddIns\"
End Function

Private Function AddInDir() As String
    AddInDir = AddInsRoot() & ADDIN_SUBDIR
End Function

Private Function FileLenSafe(ByVal path As String) As Long
    On Error Resume Next
    FileLenSafe = FileLen(path)
End Function

' Create a folder and any missing parents. Accepts a trailing-slash path.
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object, parts() As String, build As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1)
    parts = Split(folderPath, "\")
    build = parts(0)                         ' drive, e.g. "C:"
    For i = 1 To UBound(parts)
        build = build & "\" & parts(i)
        If Not fso.FolderExists(build) Then fso.CreateFolder build
    Next i
End Sub
