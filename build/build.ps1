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

# --- Copy the source package to the output, then patch it in place ----------
Copy-Item -LiteralPath $InputPpam -Destination $OutputPpam -Force

function Set-ZipEntryBytes {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName, [byte[]]$Bytes)
    $existing = $Zip.GetEntry($EntryName)
    if ($existing) { $existing.Delete() }
    $entry  = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    try { $stream.Write($Bytes, 0, $Bytes.Length) } finally { $stream.Dispose() }
}

function Get-ZipEntryText {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

$zip = [System.IO.Compression.ZipFile]::Open($OutputPpam, 'Update')
try {
    # 1) customUI.xml + its rels
    Set-ZipEntryBytes $zip 'customUI/customUI.xml' `
        ([System.IO.File]::ReadAllBytes((Join-Path $RibbonDir 'customUI.xml')))
    Set-ZipEntryBytes $zip 'customUI/_rels/customUI.xml.rels' `
        ([System.IO.File]::ReadAllBytes((Join-Path $RibbonDir 'customUI.xml.rels')))

    # 2) images
    $imgCount = 0
    Get-ChildItem -Path $ImagesDir -Filter '*.png' -File | ForEach-Object {
        Set-ZipEntryBytes $zip ("customUI/images/" + $_.Name) ([System.IO.File]::ReadAllBytes($_.FullName))
        $imgCount++
    }

    # 3) [Content_Types].xml -> ensure png default exists
    $ct = Get-ZipEntryText $zip '[Content_Types].xml'
    if (-not $ct) { throw "[Content_Types].xml missing from package -- is this a valid .ppam?" }
    if ($ct -notmatch '(?i)Extension="png"') {
        $ct = $ct -replace '(?i)(<Default\s+Extension="xml"[^>]*/>)', ('$1<Default Extension="png" ContentType="image/png" />')
        Set-ZipEntryBytes $zip '[Content_Types].xml' ([System.Text.Encoding]::UTF8.GetBytes($ct))
        Write-Host "Patched [Content_Types].xml (added png default)."
    }

    # 4) _rels/.rels -> ensure the customUI extensibility relationship exists
    $rels = Get-ZipEntryText $zip '_rels/.rels'
    if (-not $rels) { throw "_rels/.rels missing from package -- is this a valid .ppam?" }
    if ($rels -notmatch '(?i)ui/extensibility') {
        $relXml = '<Relationship Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="/customUI/customUI.xml" Id="rIdTQCustomUI" />'
        $rels = $rels -replace '(?i)</Relationships>', ($relXml + '</Relationships>')
        Set-ZipEntryBytes $zip '_rels/.rels' ([System.Text.Encoding]::UTF8.GetBytes($rels))
        Write-Host "Patched _rels/.rels (added customUI relationship)."
    }

    Write-Host "Injected ribbon + $imgCount image(s)."
}
finally {
    $zip.Dispose()
}

Write-Host ""
Write-Host "BUILD COMPLETE -> $OutputPpam" -ForegroundColor Green
Write-Host "Next: commit dist\TrialQuest.ppam + version.json, then push (clients pull this via the Update button)."
