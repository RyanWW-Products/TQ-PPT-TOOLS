<#
.SYNOPSIS
    Builds the distributable Trial Quest add-in by injecting the tracked ribbon
    (ribbon/customUI.xml + ribbon/images/*) into a macro-only .ppam.

    This REPLACES the manual "open in the Office RibbonX Editor and apply the
    ribbon" step. Author your macros in the .pptm, Save As .ppam, then run this.

.DESCRIPTION
    A .ppam is just an OPC (zip) package. PowerPoint's "Save As .ppam" produces a
    package that contains your VBA project but no ribbon. This script copies that
    package to dist\TrialQuest.ppam and injects:
        /customUI/customUI.xml
        /customUI/_rels/customUI.xml.rels
        /customUI/images/*.png
    then ensures [Content_Types].xml declares the png default and that the package
    root relationships (_rels/.rels) point at the customUI part.

    The operation is idempotent: running it on a .ppam that already has a ribbon
    simply overwrites the ribbon parts with the tracked source.

.PARAMETER InputPpam
    The macro-only .ppam you saved from the .pptm. Defaults to the newest
    "TrialQuest Addin Master v*.ppam" in the repo root.

.PARAMETER OutputPpam
    Where to write the ribboned add-in. Defaults to dist\TrialQuest.ppam.

.PARAMETER Version
    Optional. If supplied (e.g. 5.4.4), version.json's addinVersion is updated to
    match. You must ALSO set Const ADDIN_VERSION in Updater.bas to the same value
    before exporting the .ppam, or the Update button will misreport the version.

.EXAMPLE
    .\build\build.ps1 -InputPpam ".\TrialQuest Addin Master v5-4-4.ppam" -Version 5.4.4
#>
[CmdletBinding()]
param(
    [string]$InputPpam,
    [string]$OutputPpam,
    [string]$Version
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# --- Resolve paths relative to the repo root (parent of this script's folder) ---
$RepoRoot   = Split-Path -Parent $PSScriptRoot
$RibbonDir  = Join-Path $RepoRoot 'ribbon'
$ImagesDir  = Join-Path $RibbonDir 'images'
$DistDir    = Join-Path $RepoRoot 'dist'
$VersionFile= Join-Path $RepoRoot 'version.json'

if (-not $InputPpam) {
    $candidate = Get-ChildItem -Path $RepoRoot -Filter '*.ppam' -File |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $candidate) { throw "No input .ppam found in $RepoRoot. Pass -InputPpam explicitly." }
    $InputPpam = $candidate.FullName
}
if (-not (Test-Path $InputPpam)) { throw "Input .ppam not found: $InputPpam" }
if (-not (Test-Path (Join-Path $RibbonDir 'customUI.xml'))) { throw "ribbon\customUI.xml not found." }

if (-not $OutputPpam) {
    if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }
    $OutputPpam = Join-Path $DistDir 'TrialQuest.ppam'
}

Write-Host "Input  : $InputPpam"
Write-Host "Output : $OutputPpam"

# --- Optionally stamp version.json -----------------------------------------
if ($Version) {
    if (Test-Path $VersionFile) {
        $vj = Get-Content $VersionFile -Raw | ConvertFrom-Json
        $vj.addinVersion = $Version
        ($vj | ConvertTo-Json -Depth 10) | Set-Content -Path $VersionFile -Encoding utf8
        Write-Host "version.json addinVersion -> $Version"
    } else {
        Write-Warning "version.json not found; skipping version stamp."
    }
    Write-Warning "Make sure Const ADDIN_VERSION = `"$Version`" in Updater.bas matched the .ppam you just exported."
}

# --- Rebuild the package, PRESERVING entry order, and inject the ribbon ------
# Do NOT use ZipArchive 'Update' mode: it reorders the package (rewritten parts
# get moved to the end), and PowerPoint's add-in loader rejects the result. We
# instead copy every input part in its original order -- patching
# [Content_Types].xml and _rels/.rels in place and replacing any existing ribbon
# parts -- then append any new ribbon parts at the end (like the RibbonX Editor).

function Convert-FromBytes {
    param([byte[]]$Bytes)
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return @{ Text = [System.Text.Encoding]::UTF8.GetString($Bytes, 3, $Bytes.Length - 3); Bom = $true }
    }
    return @{ Text = [System.Text.Encoding]::UTF8.GetString($Bytes); Bom = $false }
}
function Convert-ToBytes { param([string]$Text, [bool]$Bom)
    (New-Object System.Text.UTF8Encoding($Bom)).GetBytes($Text)
}
function Edit-ContentTypes { param([byte[]]$Bytes)
    $d = Convert-FromBytes $Bytes
    if ($d.Text -notmatch '(?i)Extension="png"') {
        $d.Text = $d.Text -replace '(?i)(<Default\s+Extension="xml"[^>]*/>)', ('$1<Default Extension="png" ContentType="image/png" />')
        Write-Host "Patched [Content_Types].xml (added png default)."
    }
    Convert-ToBytes $d.Text $d.Bom
}
function Edit-RootRels { param([byte[]]$Bytes)
    $d = Convert-FromBytes $Bytes
    if ($d.Text -notmatch '(?i)ui/extensibility') {
        $rel = '<Relationship Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="/customUI/customUI.xml" Id="rIdTQCustomUI" />'
        $d.Text = $d.Text -replace '(?i)</Relationships>', ($rel + '</Relationships>')
        Write-Host "Patched _rels/.rels (added customUI relationship)."
    }
    Convert-ToBytes $d.Text $d.Bom
}

# Ribbon parts to inject (name -> bytes)
$ribbon = [ordered]@{}
$ribbon['customUI/customUI.xml']            = [System.IO.File]::ReadAllBytes((Join-Path $RibbonDir 'customUI.xml'))
$ribbon['customUI/_rels/customUI.xml.rels'] = [System.IO.File]::ReadAllBytes((Join-Path $RibbonDir 'customUI.xml.rels'))
Get-ChildItem -Path $ImagesDir -Filter '*.png' -File | ForEach-Object {
    $ribbon["customUI/images/$($_.Name)"] = [System.IO.File]::ReadAllBytes($_.FullName)
}

# Read input entries in order; patch content-types/rels, replace existing ribbon parts.
$inZip   = [System.IO.Compression.ZipFile]::OpenRead($InputPpam)
$parts   = New-Object System.Collections.ArrayList
$present = @{}
foreach ($e in $inZip.Entries) {
    $ms = New-Object System.IO.MemoryStream
    $s  = $e.Open(); $s.CopyTo($ms); $s.Dispose()
    $bytes = $ms.ToArray()
    $name  = $e.FullName
    if     ($name -eq '[Content_Types].xml') { $bytes = Edit-ContentTypes $bytes }
    elseif ($name -eq '_rels/.rels')         { $bytes = Edit-RootRels $bytes }
    elseif ($ribbon.Contains($name))         { $bytes = $ribbon[$name] }   # replace in place
    [void]$parts.Add(@{ Name = $name; Bytes = $bytes })
    $present[$name] = $true
}
$inZip.Dispose()

if (-not $present.ContainsKey('[Content_Types].xml')) { throw "[Content_Types].xml missing -- is this a valid .ppam?" }
if (-not $present.ContainsKey('_rels/.rels'))         { throw "_rels/.rels missing -- is this a valid .ppam?" }

# Append any ribbon parts the input didn't already have (a fresh .ppam has none).
foreach ($k in $ribbon.Keys) {
    if (-not $present.ContainsKey($k)) { [void]$parts.Add(@{ Name = $k; Bytes = $ribbon[$k] }) }
}
$imgCount = ($ribbon.Keys | Where-Object { $_ -like 'customUI/images/*' }).Count

# Write the output package in this exact order.
if (Test-Path $OutputPpam) { Remove-Item -LiteralPath $OutputPpam -Force }
$outStream = [System.IO.File]::Open($OutputPpam, [System.IO.FileMode]::Create)
$outZip    = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($p in $parts) {
        $entry = $outZip.CreateEntry($p.Name, [System.IO.Compression.CompressionLevel]::Optimal)
        $st = $entry.Open()
        try { $st.Write($p.Bytes, 0, $p.Bytes.Length) } finally { $st.Dispose() }
    }
}
finally {
    $outZip.Dispose()
    $outStream.Dispose()
}
Write-Host "Injected ribbon + $imgCount image(s)."

Write-Host ""
Write-Host "BUILD COMPLETE -> $OutputPpam" -ForegroundColor Green
Write-Host "Next: commit dist\TrialQuest.ppam + version.json, then push (clients pull this via the Update button)."
