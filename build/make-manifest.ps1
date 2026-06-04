<#
.SYNOPSIS
    Regenerates manifest.json from the contents of the assets\ folder.

.DESCRIPTION
    Walks every file under assets\, records its repo-relative path (forward
    slashes) and byte size, and writes manifest.json at the repo root in the
    exact shape the add-in's Updater.bas expects:

        { "assetsVersion": N, "files": [ {"path":"...","size":123}, ... ] }

    The Update button downloads any asset whose local size differs from the
    size recorded here (or is missing). So: whenever you change a template,
    bump "assetsVersion" in version.json, rerun this, then commit + push.

.PARAMETER AssetsVersion
    Override the assetsVersion written into the manifest. Defaults to the value
    in version.json so the two files stay in sync.

.EXAMPLE
    .\build\make-manifest.ps1
#>
[CmdletBinding()]
param([int]$AssetsVersion = -1)

$ErrorActionPreference = 'Stop'
$RepoRoot     = Split-Path -Parent $PSScriptRoot
$AssetsDir    = Join-Path $RepoRoot 'assets'
$VersionFile  = Join-Path $RepoRoot 'version.json'
$ManifestFile = Join-Path $RepoRoot 'manifest.json'

if (-not (Test-Path $AssetsDir)) { throw "assets\ folder not found at $AssetsDir" }

if ($AssetsVersion -lt 0) {
    if (Test-Path $VersionFile) {
        $AssetsVersion = [int](Get-Content $VersionFile -Raw | ConvertFrom-Json).assetsVersion
    } else {
        $AssetsVersion = 1
    }
}

$prefixLen = $AssetsDir.Length + 1
$files = Get-ChildItem -Path $AssetsDir -Recurse -File | Sort-Object FullName | ForEach-Object {
    $rel = $_.FullName.Substring($prefixLen) -replace '\\', '/'
    '    {{"path":"{0}","size":{1}}}' -f $rel, $_.Length
}

$json = @()
$json += '{'
$json += '  "assetsVersion": {0},' -f $AssetsVersion
$json += '  "files": ['
$json += ($files -join ",`r`n")
$json += '  ]'
$json += '}'

# Write UTF-8 without BOM so the GitHub raw bytes parse cleanly everywhere.
[System.IO.File]::WriteAllText($ManifestFile, ($json -join "`r`n") + "`r`n", `
    (New-Object System.Text.UTF8Encoding($false)))

Write-Host ("manifest.json written: {0} file(s), assetsVersion {1}" -f $files.Count, $AssetsVersion) -ForegroundColor Green
