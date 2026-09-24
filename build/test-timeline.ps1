<#
.SYNOPSIS
    Runs the production timeline parser and placement routines in an isolated
    temporary Office VBA project. No existing presentation or source workbook is modified.
.DESCRIPTION
    Requires Excel for import tests and existing VBA project object model access
    in the selected host. Use -HostPowerPoint if Excel project access is disabled.
    This script does not change Office trust settings. The routines are extracted
    from TimelineCreator.bas so these checks execute VBA, not a port of the logic.
    FixturePath accepts the 30-event timeline_import_2099_corrected.xlsx;
    OriginalFixturePath accepts the 37-event original, with seven untimed rows.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\build\test-timeline.ps1 -FixturePath 'C:\path\timeline_import_2099_corrected.xlsx'
.EXAMPLE
    .\build\test-timeline.ps1 -EmitModule '.\build\Output\TimelineRegression.bas'
.EXAMPLE
    .\build\test-timeline.ps1 -HostPowerPoint -FixturePath 'C:\path\timeline_import_2099_corrected.xlsx'
#>
[CmdletBinding()]
param(
    [string]$ModulePath,
    [string]$FixturePath,
    [string]$OriginalFixturePath,
    [string]$EmitModule,
    [switch]$HostPowerPoint
)

$ErrorActionPreference = 'Stop'
if (-not $ModulePath) { $ModulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules\TimelineCreator.bas' }
$source = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $ModulePath))
if ($FixturePath) { $FixturePath = (Resolve-Path -LiteralPath $FixturePath).Path }
if ($OriginalFixturePath) { $OriginalFixturePath = (Resolve-Path -LiteralPath $OriginalFixturePath).Path }

function Get-VbaRoutine([string]$Name) {
    $pattern = '(?ms)^(?:Private|Public) (Function|Sub) ' + [regex]::Escape($Name) + '\b.*?^End \1\s*$'
    $match = [regex]::Match($source, $pattern)
    if (-not $match.Success) { throw "Production VBA routine not found: $Name" }
    $match.Value.TrimEnd()
}

$type = [regex]::Match($source, '(?ms)^Private Type TLEvent\b.*?^End Type\s*$')
if (-not $type.Success) { throw 'Production TLEvent type not found.' }
$routines = @(
    'ReadEvents', 'ParseDateCell', 'MonthFromName',
    'TryParseTimeCell', 'TryClockTime', 'ClockLabel', 'NormalizeMeridiem',
    'PrepareEventPositions', 'PrepareEventUnits', 'EventsHaveTimes', 'SortEvents',
    'DetectType', 'IntervalCode', 'UnitStartOf', 'UnitFracOf', 'ComputeColumns', 'ClampD'
)
$testCode = @'
Private checkCount As Long

Public Function RunTimelineRegressionTests(ByVal fixturePath As String, _
                                          Optional ByVal originalFixturePath As String = "") As String
    On Error GoTo Failed
    checkCount = 0
    CheckTime "22:17:51", 22, 17, 51, "10:17:51 PM"
    CheckTime "23:06", 23, 6, 0, "11:06 PM"
    CheckTime "10:30 PM", 22, 30, 0, "10:30 PM"
    CheckTime "10:30 p.m.", 22, 30, 0, "10:30 PM"
    CheckTime "10:30 p. m.", 22, 30, 0, "10:30 PM"
    CheckTime "12:00 AM", 0, 0, 0, "12:00 AM"
    CheckTime "12:00 PM", 12, 0, 0, "12:00 PM"
    CheckTime 0, 0, 0, 0, "12:00 AM"
    CheckTime 0.5, 12, 0, 0, "12:00 PM"
    CheckTime TimeSerial(23, 14, 58), 23, 14, 58, "11:14:58 PM"
    CheckRange "22:47:35 - 22:56:28", 22, 47, 35, "10:47:35 PM", "10:56:28 PM"
    CheckRange "23:01:07-23:01:23", 23, 1, 7, "11:01:07 PM", "11:01:23 PM"
    CheckRange "23:35 " & ChrW(8211) & " 23:43", 23, 35, 0, "11:35 PM", "11:43 PM"
    CheckRange "22:10" & ChrW(8212) & "22:20", 22, 10, 0, "10:10 PM", "10:20 PM"
    CheckRange "22:10 to 22:20", 22, 10, 0, "10:10 PM", "10:20 PM"
    CheckRange "11:50 PM - 12:10 AM", 23, 50, 0, "11:50 PM", "12:10 AM"
    CheckRejected ""
    CheckRejected "   "
    CheckRejected Empty
    CheckRejected Null
    CheckRejected "unknown"
    CheckRejected "25:00"
    CheckRejected "22:10 - unknown"
    CheckRejected "22:10-22:20-22:30"
    CheckMixedPlacement
    CheckLeadingUntimed
    CheckDateOnlyPlacement
    CheckMidnightPlacement
    If Len(fixturePath) > 0 Then CheckCorrectedFixture fixturePath
    If Len(originalFixturePath) > 0 Then CheckOriginalFixture originalFixturePath
    RunTimelineRegressionTests = "PASS: " & checkCount & " timeline regression assertions" & _
        IIf(Len(fixturePath) > 0, " (corrected 30-event workbook included)", "") & _
        IIf(Len(originalFixturePath) > 0, " (original 37-event workbook included)", "") & "."
    Exit Function
Failed:
    RunTimelineRegressionTests = "FAIL after " & checkCount & " passing assertions: " & Err.Description
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2199, "TimelineRegression", message
    checkCount = checkCount + 1
End Sub

Private Sub CheckTime(ByVal value As Variant, ByVal h As Long, ByVal m As Long, _
                      ByVal s As Long, ByVal wantedLabel As String)
    Dim fraction As Double, label As String, accepted As Boolean
    accepted = TryParseTimeCell(value, fraction, label)
    Check accepted, "Time rejected: " & CStr(value)
    Check Abs(fraction * 86400 - (h * 3600 + m * 60 + s)) < 0.001, "Incorrect clock position: " & CStr(value)
    Check label = wantedLabel, "Incorrect clock label for " & CStr(value) & ": " & label
End Sub

Private Sub CheckRange(ByVal value As String, ByVal h As Long, ByVal m As Long, _
                       ByVal s As Long, ByVal firstLabel As String, ByVal lastLabel As String)
    Dim fraction As Double, label As String, accepted As Boolean
    accepted = TryParseTimeCell(value, fraction, label)
    Check accepted, "Time range rejected: " & value
    Check Abs(fraction * 86400 - (h * 3600 + m * 60 + s)) < 0.001, "Incorrect range start position: " & value
    Check InStr(1, label, firstLabel, vbBinaryCompare) = 1, "Range start label missing: " & label
    Check InStr(1, label, lastLabel, vbBinaryCompare) > Len(firstLabel), "Range end label missing: " & label
End Sub

Private Sub CheckRejected(ByVal value As Variant)
    Dim fraction As Double, label As String
    Check Not TryParseTimeCell(value, fraction, label), "Blank or invalid clock value became a real time"
End Sub

Private Sub SetEvent(ByRef ev As TLEvent, ByVal dayValue As Date, ByVal rowNumber As Long, _
                     ByVal timed As Boolean, Optional ByVal h As Long = 0, _
                     Optional ByVal m As Long = 0, Optional ByVal s As Long = 0)
    ev.RawDate = dayValue + TimeSerial(h, m, s)
    ev.OrigIndex = rowNumber
    ev.HasTime = timed
    ev.Prec = IIf(timed, 4, 3)
    ev.DateLabel = Format$(dayValue, "mmm d")
    If timed Then ev.DateLabel = ev.DateLabel & "  " & ClockLabel(ev.RawDate)
End Sub

Private Sub CheckMixedPlacement()
    Dim ev(1 To 5) As TLEvent, cols() As Date, d As Date, i As Long, n As Long
    d = DateSerial(2099, 1, 1)
    SetEvent ev(1), d, 2, True, 22, 17, 51
    SetEvent ev(2), d, 3, False
    SetEvent ev(3), d, 4, True, 23, 10
    SetEvent ev(4), d, 5, False
    SetEvent ev(5), d, 6, True, 23, 14, 58
    PrepareEventPositions ev, 5
    SortEvents ev, 5
    PrepareEventUnits ev, 5, "Hours"
    For i = 1 To 5
        Check ev(i).OrigIndex = i + 1, "Mixed untimed row lost its spreadsheet order"
    Next i
    Check ev(2).RawDate = d And Not ev(2).HasTime, "Untimed event acquired a fabricated timestamp"
    Check ev(4).RawDate = d And Not ev(4).HasTime, "Second untimed event acquired a fabricated timestamp"
    Check ev(2).PlacementDate = ev(1).RawDate, "Untimed event left the preceding timed event's position"
    Check ev(4).PlacementDate = ev(3).RawDate, "Untimed event left the second preceding timed event's position"
    Check EventsHaveTimes(ev, 5), "Mixed timeline was classified as date-only"
    Check DetectType(ev, 5) = "Hours", "Mixed same-night events were not classified as Hours"
    n = ComputeColumns(ev, 5, "Hours", False, cols)
    Check n = 2, "Mixed timeline must have two hour shapes, with no midnight shape"
    Check Hour(cols(1)) = 22 And Hour(cols(2)) = 23, "Mixed timeline should begin at 10 PM and then 11 PM"
    Check Abs((CDbl(DateAdd("h", 1, cols(2))) - CDbl(DateSerial(2099, 1, 2))) * 86400#) < 0.001, _
        "Midnight must be the right edge of the final hour. " & HourEdgeDiagnostic(cols(1), cols(2), DateSerial(2099, 1, 2))
    CheckColumnMembership ev, 5, cols, n
    n = ComputeColumns(ev, 5, "Hours", True, cols)
    Check n = 2, "Compact mixed timeline gained an extra hour shape"
    CheckColumnMembership ev, 5, cols, n
End Sub

Private Sub CheckColumnMembership(ByRef ev() As TLEvent, ByVal eventCount As Long, ByRef cols() As Date, ByVal columnCount As Long)
    Dim i As Long, ci As Long, found As Boolean
    For i = 1 To eventCount
        found = False
        For ci = 1 To columnCount
            If ev(i).UnitStart = cols(ci) Then found = True: Exit For
        Next ci
        Check found, "Event row " & ev(i).OrigIndex & " has no exactly matching column key: " & DateDiagnostic(ev(i).UnitStart)
    Next i
End Sub

Private Function HourEdgeDiagnostic(ByVal firstColumn As Date, ByVal lastColumn As Date, ByVal expectedEdge As Date) As String
    Dim edge As Date
    edge = DateAdd("h", 1, lastColumn)
    HourEdgeDiagnostic = "first=" & DateDiagnostic(firstColumn) & "; last=" & DateDiagnostic(lastColumn) & _
        "; actualEdge=" & DateDiagnostic(edge) & "; expectedEdge=" & DateDiagnostic(expectedEdge) & _
        "; differenceSeconds=" & Format$((CDbl(edge) - CDbl(expectedEdge)) * 86400#, "0.000000000")
End Function

Private Function DateDiagnostic(ByVal value As Date) As String
    DateDiagnostic = Format$(value, "yyyy-mm-dd hh:nn:ss") & " (serial=" & Format$(CDbl(value), "0.00000000000000000") & ")"
End Function

Private Sub CheckLeadingUntimed()
    Dim ev(1 To 3) As TLEvent, d As Date
    d = DateSerial(2099, 1, 1)
    SetEvent ev(1), d, 2, False
    SetEvent ev(2), d, 3, True, 22, 30
    SetEvent ev(3), d, 4, False
    PrepareEventPositions ev, 3
    SortEvents ev, 3
    Check ev(1).OrigIndex = 2 And ev(2).OrigIndex = 3 And ev(3).OrigIndex = 4, "Leading untimed event lost spreadsheet order"
    Check ev(1).PlacementDate = d + TimeSerial(22, 30, 0), "Leading untimed event should use the next timed same-day entry for layout"
    Check ev(1).RawDate = d And Not ev(1).HasTime, "Leading untimed event gained a displayed time"
End Sub

Private Sub CheckDateOnlyPlacement()
    Dim ev(1 To 3) As TLEvent, d As Date, i As Long
    d = DateSerial(2099, 1, 1)
    SetEvent ev(1), d, 2, False
    SetEvent ev(2), d, 3, False
    SetEvent ev(3), DateAdd("d", 1, d), 4, False
    PrepareEventPositions ev, 3
    SortEvents ev, 3
    PrepareEventUnits ev, 3, "Days"
    Check Not EventsHaveTimes(ev, 3), "Date-only events were classified as timed"
    For i = 1 To 3
        Check Abs(ev(i).UnitFrac - 0.5) < 0.000001, "Date-only leader should meet the middle of its day"
        Check Int(CDbl(ev(i).RawDate)) = CDbl(ev(i).RawDate), "Date-only input was overwritten with a fabricated time"
    Next i
End Sub

Private Sub CheckMidnightPlacement()
    Dim ev(1 To 3) As TLEvent, d As Date, cols() As Date, n As Long
    d = DateSerial(2099, 1, 1)
    SetEvent ev(1), d, 2, True, 23, 50
    SetEvent ev(2), DateAdd("d", 1, d), 3, True, 0, 0
    SetEvent ev(3), DateAdd("d", 1, d), 4, False
    PrepareEventPositions ev, 3
    SortEvents ev, 3
    PrepareEventUnits ev, 3, "Hours"
    Check ev(2).HasTime, "Real midnight must remain distinguishable from no time"
    Check Not ev(3).HasTime, "Untimed row after midnight must remain untimed"
    Check ev(3).RawDate = DateAdd("d", 1, d), "Untimed event crossed to the previous date"
    Check ev(3).OrigIndex = 4, "Untimed event crossed ahead of midnight"
    n = ComputeColumns(ev, 3, "Hours", False, cols)
    Check n = 2, "An actual midnight event needs its own hour shape"
    Check cols(1) = d + TimeSerial(23, 0, 0) And cols(2) = DateAdd("d", 1, d), "Overnight columns are incorrect"
    CheckColumnMembership ev, 3, cols, n
End Sub

Private Sub CheckCorrectedFixture(ByVal filePath As String)
    Dim ev() As TLEvent, cols() As Date, n As Long, nCols As Long
    Dim errLog As String, i As Long, rangeCount As Long, d As Date
    n = ReadEvents(filePath, ev, errLog)
    Check n = 30, "Corrected workbook should import all 30 entries; got " & n & ". " & errLog
    Check Len(errLog) = 0, "Corrected workbook reported import errors: " & errLog
    Check DetectType(ev, n) = "Hours", "Corrected workbook must be an Hours timeline"
    d = DateSerial(2099, 1, 1)
    PrepareEventUnits ev, n, "Hours"
    For i = 1 To n
        Check ev(i).HasTime, "Corrected workbook lost a valid time on row " & ev(i).OrigIndex
        Check ev(i).RawDate >= d + TimeSerial(22, 0, 0) And ev(i).RawDate < d + 1, "Corrected workbook event fell outside 10 PM to midnight"
        If InStr(ev(i).DateLabel, "10:56:28 PM") > 0 Or InStr(ev(i).DateLabel, "10:59:35 PM") > 0 _
          Or InStr(ev(i).DateLabel, "11:01:23 PM") > 0 Or InStr(ev(i).DateLabel, "11:43 PM") > 0 Then rangeCount = rangeCount + 1
    Next i
    Check rangeCount = 4, "Corrected workbook did not preserve all four range end labels"
    Check Abs((CDbl(ev(1).RawDate) - CDbl(d)) * 86400 - 80271) < 0.001, "First event must retain 22:17:51 including seconds"
    nCols = ComputeColumns(ev, n, "Hours", False, cols)
    Check nCols = 2, "Corrected workbook should have exactly two hour shapes"
    Check cols(1) = d + TimeSerial(22, 0, 0) And cols(2) = d + TimeSerial(23, 0, 0), "Corrected workbook columns must be 10 PM and 11 PM"
    Check Abs((CDbl(DateAdd("h", 1, cols(2))) - CDbl(d + 1)) * 86400#) < 0.001, _
        "Corrected workbook must end at the midnight edge. " & HourEdgeDiagnostic(cols(1), cols(2), d + 1)
    CheckColumnMembership ev, n, cols, nCols
    nCols = ComputeColumns(ev, n, "Hours", True, cols)
    Check nCols = 2, "Corrected workbook compact mode should also have two hour shapes"
    CheckColumnMembership ev, n, cols, nCols
End Sub

Private Sub CheckOriginalFixture(ByVal filePath As String)
    Dim ev() As TLEvent, cols() As Date, n As Long, nCols As Long
    Dim errLog As String, i As Long, untimedCount As Long, d As Date
    n = ReadEvents(filePath, ev, errLog)
    Check n = 37, "Original workbook should import all 37 entries; got " & n & ". " & errLog
    Check Len(errLog) = 0, "Original workbook reported import errors: " & errLog
    Check DetectType(ev, n) = "Hours", "Original workbook must be an Hours timeline"
    d = DateSerial(2099, 1, 1)
    PrepareEventUnits ev, n, "Hours"
    For i = 1 To n
        Check ev(i).OrigIndex = i + 1, "Original workbook row moved out of spreadsheet order: " & ev(i).OrigIndex
        If Not ev(i).HasTime Then
            untimedCount = untimedCount + 1
            Check ev(i).RawDate = d, "Original untimed event acquired a fabricated timestamp"
            Check InStr(ev(i).DateLabel, ":") = 0, "Original untimed event acquired a displayed time"
            Check ev(i).PlacementDate >= d + TimeSerial(22, 0, 0), "Original untimed event created a midnight column"
        End If
    Next i
    Check untimedCount = 7, "Original workbook should retain seven untimed events; got " & untimedCount
    nCols = ComputeColumns(ev, n, "Hours", False, cols)
    Check nCols = 2, "Original workbook should have exactly two hour shapes"
    Check cols(1) = d + TimeSerial(22, 0, 0) And cols(2) = d + TimeSerial(23, 0, 0), "Original workbook columns must be 10 PM and 11 PM"
    CheckColumnMembership ev, n, cols, nCols
    nCols = ComputeColumns(ev, n, "Hours", True, cols)
    Check nCols = 2, "Original workbook compact mode should also have two hour shapes"
    CheckColumnMembership ev, n, cols, nCols
End Sub
'@
$moduleCode = "Option Explicit`r`nPrivate Const MAX_COLS As Long = 75`r`n" +
    $type.Value.TrimEnd() + "`r`nPrivate checkCount As Long`r`n`r`n" +
    (($routines | ForEach-Object { Get-VbaRoutine $_ }) -join "`r`n`r`n") + "`r`n`r`n" +
    ($testCode -replace '(?m)^Private checkCount As Long\r?\n', '')

if ($HostPowerPoint) {
    $runId = [Guid]::NewGuid().ToString('N')
    $runDirectory = Join-Path $PSScriptRoot 'Output'
    [void][IO.Directory]::CreateDirectory($runDirectory)
    $presentationPath = Join-Path $runDirectory "TimelineRegression-$runId.pptm"
    $resultPath = Join-Path $runDirectory "TimelineRegression-$runId.txt"
    $fixtureLiteral = ([string]$FixturePath).Replace('"', '""')
    $originalLiteral = ([string]$OriginalFixturePath).Replace('"', '""')
    $resultLiteral = $resultPath.Replace('"', '""')
    $moduleCode += @"

Public Sub RunTimelineRegressionToFile()
    Dim result As String, outputNumber As Integer
    On Error GoTo Failed
    result = RunTimelineRegressionTests("$fixtureLiteral", "$originalLiteral")
WriteResult:
    On Error GoTo 0
    outputNumber = FreeFile
    Open "$resultLiteral" For Output As #outputNumber
    Print #outputNumber, result
    Close #outputNumber
    Exit Sub
Failed:
    result = "FAIL in regression wrapper: " & Err.Number & ": " & Err.Description
    Resume WriteResult
End Sub
"@
}

if ($EmitModule) {
    $outputPath = [IO.Path]::GetFullPath($EmitModule)
    [IO.File]::WriteAllText($outputPath, "Attribute VB_Name = `"TimelineRegression`"`r`n" + $moduleCode, [Text.Encoding]::Default)
    Write-Output "Extracted production VBA and regression assertions to $outputPath"
    return
}

if ($HostPowerPoint) {
    $powerPoint = $null
    $presentation = $null
    $component = $null
    $quitPowerPoint = $false
    try {
        # PowerPoint's PIA declares Run(string, ref object[]). PowerShell's COM
        # binder cannot handle that ByRef ParamArray reliably, so use a tiny
        # strongly typed C# bridge. This does not alter Office or macro security.
        if (-not ('TimelinePowerPointMacroRunner' -as [type])) {
            $interop = [Reflection.Assembly]::Load('Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
            Add-Type -ReferencedAssemblies $interop.Location -TypeDefinition @'
public static class TimelinePowerPointMacroRunner {
    public static object Run(object application, string macroName) {
        object[] arguments = new object[0];
        return ((Microsoft.Office.Interop.PowerPoint._Application)application).Run(macroName, ref arguments);
    }
}
'@
        }
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $quitPowerPoint = ($powerPoint.Presentations.Count -eq 0)
        $presentation = $powerPoint.Presentations.Add(0)
        $project = $presentation.VBProject
        if ($null -eq $project -or $null -eq $project.VBComponents) {
            throw 'PowerPoint VBA project access is unavailable. No Office trust settings were changed.'
        }
        $component = $project.VBComponents.Add(1)
        $component.Name = 'TimelineRegression'
        $component.CodeModule.AddFromString($moduleCode)
        # ppSaveAsOpenXMLPresentationMacroEnabled = 25. Saving makes the macro's
        # presentation qualifier stable and leaves an isolated diagnostic artifact.
        $presentation.SaveAs($presentationPath, 25)
        $macro = $presentation.Name + '!TimelineRegression.RunTimelineRegressionToFile'
        Write-Output "PowerPoint regression presentation: $presentationPath"
        Write-Output "PowerPoint regression result: $resultPath"
        [void][TimelinePowerPointMacroRunner]::Run($powerPoint, $macro)
        if (-not [IO.File]::Exists($resultPath)) {
            throw "PowerPoint did not write a regression result. Saved isolated test project: $presentationPath"
        }
        $result = [IO.File]::ReadAllText($resultPath).Trim()
        Write-Output $result
        if (-not $result.StartsWith('PASS:')) { throw 'Timeline regression checks failed.' }
    } finally {
        if ($null -ne $presentation) { $presentation.Close() }
        if ($null -ne $powerPoint -and $quitPowerPoint) { $powerPoint.Quit() }
        foreach ($comObject in @($component, $presentation, $powerPoint)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    return
}

$excel = $null
$book = $null
$component = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $project = $book.VBProject
    if ($null -eq $project -or $null -eq $project.VBComponents) {
        throw 'Excel VBA project access is unavailable. No Office trust settings were changed.'
    }
    $component = $project.VBComponents.Add(1)
    $component.Name = 'TimelineRegression'
    $component.CodeModule.AddFromString($moduleCode)
    $result = [string]$excel.Run("'$($book.Name)'!TimelineRegression.RunTimelineRegressionTests", [string]$FixturePath, [string]$OriginalFixturePath)
    Write-Output $result
    if (-not $result.StartsWith('PASS:')) { throw 'Timeline regression checks failed.' }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    foreach ($comObject in @($component, $book, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
