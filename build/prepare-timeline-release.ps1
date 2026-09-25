<#
.SYNOPSIS
    Imports the exported standard and class modules into a copy of the local master and
    saves a new PPTM and PPAM. Requires PowerPoint and trusted VBProject access.

.DESCRIPTION
    The source master and presentations already open in PowerPoint are left
    alone. Existing forms, their resources, references, and document modules are
    preserved. Tracked and new nonignored Modules/*.bas and Modules/*.cls files are imported.

    Macros are disabled while the private staging copy opens, and the original
    AutomationSecurity setting is restored immediately afterward. This script
    never runs a macro, changes Trust Center settings, or invokes the VBE UI.
    Saving is not proof that the VBA compiles: perform runtime validation before
    publishing. The optional -Build switch injects the ribbon and stamps the
    release version; nothing is committed, pushed, or installed.

.EXAMPLE
    .\build\prepare-timeline-release.ps1 -SourcePptm '.\TrialQuest Addin Master v5-6-03.pptm'

.EXAMPLE
    .\build\prepare-timeline-release.ps1 -Version 5.6.5 -Build
#>
[CmdletBinding()]
param(
    [string]$SourcePptm,
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version = '5.6.5',
    [switch]$Build
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$repoRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$releaseVersion = [version]$Version
$releaseSuffix = '{0}-{1}-{2:00}' -f $releaseVersion.Major, $releaseVersion.Minor, $releaseVersion.Build
$releaseName = 'TrialQuest Addin Master v' + $releaseSuffix
$masterOutput = Join-Path $repoRoot ($releaseName + '.pptm')
$addinOutput = Join-Path $repoRoot ($releaseName + '.ppam')

function Get-NormalizedModuleBody([string]$text) {
    # VBIDE hides exported Attribute records and may omit terminal empty lines.
    # Keep all other whitespace and letter case exact.
    $text = [regex]::Replace($text, '(?s)\AVERSION 1\.0 CLASS\r?\nBEGIN\r?\n.*?\r?\nEND\r?\n', '')
    $text = [regex]::Replace($text, '(?m)^Attribute [^\r\n]*(?:\r?\n|$)', '')
    $text = $text.Replace("`r`n", "`n").Replace("`r", "`n")
    $text.TrimEnd([char[]]"`n")
}

function Get-VbaCaseComparableBody([string]$text) {
    # VBIDE recases identifiers against its symbol table (fd.Title -> fd.title).
    # Protect quoted literals, date literals, and both VBA comment forms first;
    # only the intervening VBA code is case folded. Double quotes inside a
    # string and continued comment lines are part of their protected segment.
    $protectedPattern = '"(?:[^"\n]|"")*"|''(?:[^\n]*[ \t]_[ \t]*\n)*[^\n]*|(?i:\bRem\b)[ \t]+(?:[^\n]*[ \t]_[ \t]*\n)*[^\n]*|#[^#\n]*#'
    $builder = New-Object Text.StringBuilder
    $cursor = 0
    foreach ($segment in [regex]::Matches($text, $protectedPattern)) {
        [void]$builder.Append($text.Substring($cursor, $segment.Index - $cursor).ToUpperInvariant())
        $literal = $segment.Value
        if ($literal -match '^(?i:Rem)[ \t]') {
            [void]$builder.Append('REM') # REM itself is a VBA keyword.
            [void]$builder.Append($literal.Substring(3))
        }
        else { [void]$builder.Append($literal) }
        $cursor = $segment.Index + $segment.Length
    }
    [void]$builder.Append($text.Substring($cursor).ToUpperInvariant())
    $builder.ToString()
}

foreach ($outputPath in @($masterOutput, $addinOutput)) {
    if (Test-Path -LiteralPath $outputPath) {
        throw "Output already exists: $outputPath. Preserve or rename that file before preparing this version again."
    }
}

if ([string]::IsNullOrWhiteSpace($SourcePptm)) {
    $candidate = Get-ChildItem -LiteralPath $repoRoot -Filter 'TrialQuest Addin Master v*.pptm' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $candidate) { throw 'No local master PPTM found. Supply -SourcePptm.' }
    $SourcePptm = $candidate.FullName
}
$SourcePptm = (Resolve-Path -LiteralPath $SourcePptm).Path
if ([IO.Path]::GetExtension($SourcePptm) -ine '.pptm') { throw 'SourcePptm must be a PPTM file.' }

# Include a newly added module before its first commit, but never ignored files.
# safe.directory is scoped to this read-only git invocation, not global config.
$modulePaths = @(& git -c "safe.directory=$repoRoot" -C $repoRoot ls-files --cached --others --exclude-standard -- 'Modules/*.bas' 'Modules/*.cls')
if ($LASTEXITCODE -ne 0) { throw 'Could not enumerate the repository VBA modules.' }
$modulePaths = @($modulePaths | Sort-Object -Unique)
if ($modulePaths.Count -eq 0) { throw 'No Modules/*.bas source files found.' }

$sources = @()
$moduleNames = @{}
foreach ($relativePath in $modulePaths) {
    $sourcePath = Join-Path $repoRoot $relativePath
    $sourceText = [IO.File]::ReadAllText($sourcePath, [Text.Encoding]::Default)
    $match = [regex]::Match($sourceText, '(?m)^Attribute VB_Name = "([^"]+)"\s*$')
    if (-not $match.Success) { throw "Missing Attribute VB_Name in $relativePath" }
    $moduleName = $match.Groups[1].Value
    if ($moduleNames.ContainsKey($moduleName)) { throw "Duplicate module name: $moduleName" }
    $componentType = if ([IO.Path]::GetExtension($sourcePath) -eq '.cls') { 2 } else { 1 }
    $moduleNames[$moduleName] = $componentType
    $sources += [pscustomobject]@{
        Name = $moduleName
        Type = $componentType
        Path = $sourcePath
        Hash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        Body = Get-NormalizedModuleBody $sourceText
    }
}

$stagingDirectory = Join-Path $PSScriptRoot ('.timeline-release-' + [Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $stagingDirectory)
$stagingMaster = Join-Path $stagingDirectory ($releaseName + '.pptm')
$stagingAddin = Join-Path $stagingDirectory ($releaseName + '.ppam')
Copy-Item -LiteralPath $SourcePptm -Destination $stagingMaster

$pptApplication = $null
$presentation = $null
$project = $null
$components = $null
$ownsApplication = $false
$completed = $false
$preservedForms = @()
$preservedOtherStandardModules = @()
try {
    try { $pptApplication = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
    catch { $pptApplication = New-Object -ComObject PowerPoint.Application; $ownsApplication = $true }

    $previousSecurity = $pptApplication.AutomationSecurity
    try {
        $pptApplication.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
        $presentation = $pptApplication.Presentations.Open($stagingMaster, 0, 0, 0)
    }
    finally { $pptApplication.AutomationSecurity = $previousSecurity }
    if ($null -eq $presentation) { throw 'PowerPoint did not return the staging presentation.' }

    try { $project = $presentation.VBProject }
    catch { throw ('Cannot access the staging VBProject. Check trusted VBA project access. ' + $_.Exception.Message) }
    if ($null -eq $project) { throw 'PowerPoint returned a null VBProject. Check trusted VBA project access.' }
    if ([int]$project.Protection -ne 0) { throw 'The staging VBA project is protected; it cannot be imported.' }
    $components = $project.VBComponents
    if ($null -eq $components -or $components.Count -eq 0) { throw 'The staging VBA project has no accessible components.' }

    # Preflight every collision before removing even one copied component.
    for ($index = 1; $index -le $components.Count; $index++) {
        $component = $components.Item($index)
        $componentName = [string]$component.Name
        $componentType = [int]$component.Type
        if ($moduleNames.ContainsKey($componentName) -and $componentType -ne $moduleNames[$componentName]) {
            throw "Exported module collides with a different component type: $componentName"
        }
        if ($componentType -eq 3) { $preservedForms += $componentName }
        if ($componentType -eq 1 -and -not $moduleNames.ContainsKey($componentName)) {
            $preservedOtherStandardModules += $componentName
        }
    }

    foreach ($source in $sources) {
        $existing = $null
        for ($index = 1; $index -le $components.Count; $index++) {
            $candidateComponent = $components.Item($index)
            if ([string]$candidateComponent.Name -ieq $source.Name) { $existing = $candidateComponent; break }
        }
        if ($null -ne $existing) { $components.Remove($existing) }
        if ($source.Type -eq 2) {
            # PowerPoint VBIDE may import .cls files as standard modules. Create
            # the class explicitly and insert only its executable source body.
            $imported = $components.Add(2)
            $imported.Name = $source.Name
            $imported.CodeModule.AddFromString($source.Body)
        }
        else { $imported = $components.Import($source.Path) }
        if ($null -eq $imported -or [string]$imported.Name -ine $source.Name -or [int]$imported.Type -ne $source.Type) {
            throw "Module import did not retain the expected name and type: $($source.Name)"
        }
        if ($imported.CodeModule.CountOfLines -eq 0) { throw "Imported module has no code: $($source.Name)" }
        Write-Host ('Imported ' + $source.Name)
    }

    # Detect concurrent source edits so the prepared release cannot silently mix revisions.
    foreach ($source in $sources) {
        if ((Get-FileHash -LiteralPath $source.Path -Algorithm SHA256).Hash -ne $source.Hash) {
            throw "Source changed during preparation: $($source.Path). Rerun after source edits are finished."
        }
        # Verify after all imports because VBIDE can resolve identifiers while
        # importing later modules. A line count alone does not prove fidelity.
        $codeModule = $components.Item($source.Name).CodeModule
        $importedBody = Get-NormalizedModuleBody ($codeModule.Lines(1, $codeModule.CountOfLines))
        $expectedComparable = Get-VbaCaseComparableBody $source.Body
        $actualComparable = Get-VbaCaseComparableBody $importedBody
        if (-not [string]::Equals($expectedComparable, $actualComparable, [StringComparison]::Ordinal)) {
            $expectedLines = $source.Body -split "`n"
            $actualLines = $importedBody -split "`n"
            $expectedComparableLines = $expectedComparable -split "`n"
            $actualComparableLines = $actualComparable -split "`n"
            $firstDifferentLine = 1
            $lineLimit = [Math]::Min($expectedLines.Count, $actualLines.Count)
            for ($lineIndex = 0; $lineIndex -lt $lineLimit; $lineIndex++) {
                if (-not [string]::Equals($expectedComparableLines[$lineIndex], $actualComparableLines[$lineIndex], [StringComparison]::Ordinal)) {
                    $firstDifferentLine = $lineIndex + 1
                    break
                }
                $firstDifferentLine = $lineIndex + 2
            }
            $expectedLine = '<end of module>'
            $actualLine = '<end of module>'
            if ($firstDifferentLine -le $expectedLines.Count) { $expectedLine = $expectedLines[$firstDifferentLine - 1] }
            if ($firstDifferentLine -le $actualLines.Count) { $actualLine = $actualLines[$firstDifferentLine - 1] }
            throw ("Imported source differs from {0} at normalized body line {1}. Release was not saved.`r`nSOURCE: {2}`r`nACTUAL: {3}" -f $source.Name, $firstDifferentLine, $expectedLine, $actualLine)
        }
    }
    Write-Host ('Verified imported source for ' + $sources.Count + ' VBA modules (VBIDE code recasing allowed; literals/comments/spacing exact).')

    $presentation.Save() # The editable copy retains the original form resources.
    $presentation.SaveCopyAs($stagingAddin, 30) # ppSaveAsOpenXMLAddin
    $presentation.Close()
    $presentation = $null
    foreach ($preparedPath in @($stagingMaster, $stagingAddin)) {
        if (-not (Test-Path -LiteralPath $preparedPath) -or (Get-Item -LiteralPath $preparedPath).Length -eq 0) {
            throw "PowerPoint did not save the prepared file: $preparedPath"
        }
    }
    # No overwrite: another preparation must not replace an output produced while we ran.
    [IO.File]::Copy($stagingMaster, $masterOutput, $false)
    [IO.File]::Copy($stagingAddin, $addinOutput, $false)
    $completed = $true
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Saved = -1; $presentation.Close() }
        catch { Write-Warning ('Could not close the private staging presentation: ' + $_.Exception.Message) }
    }
    if ($ownsApplication -and $null -ne $pptApplication) {
        # Do not close a presentation the user might have opened during preparation.
        try { if ($pptApplication.Presentations.Count -eq 0) { $pptApplication.Quit() } }
        catch { Write-Warning ('Could not close the PowerPoint instance created by this script: ' + $_.Exception.Message) }
    }
    if ($completed) {
        # Exact files only. No recursive deletion or computed tree removal.
        Remove-Item -LiteralPath $stagingMaster, $stagingAddin -Force
        try { [IO.Directory]::Delete($stagingDirectory, $false) }
        catch { Write-Warning ('Staging directory retained: ' + $stagingDirectory) }
    }
    else { Write-Warning ('Failed preparation retained for inspection: ' + $stagingDirectory) }
}

if ($Build) {
    & (Join-Path $PSScriptRoot 'build.ps1') -InputPpam $addinOutput -Version $Version
}

[pscustomobject]@{
    Version = $Version
    SourceMaster = $SourcePptm
    PreparedMaster = $masterOutput
    PreparedAddin = $addinOutput
    ImportedModules = $sources.Count
    PreservedForms = $preservedForms
    PreservedOtherStandardModules = $preservedOtherStandardModules
    BuiltDistributable = [bool]$Build
    RuntimeValidated = $false
}
