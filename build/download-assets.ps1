<#
.SYNOPSIS
    Downloads the Trial Quest template assets from the (private) GitHub repo into
    the local add-in folder. Used by the installer, and reusable standalone.

.DESCRIPTION
    Reads the GitHub read-only token from HKCU\Software\TrialQuest\Addin\GitHubToken
    (the installer writes it there), fetches version.json + manifest.json via the
    GitHub *contents* API, then downloads every asset whose local copy is missing
    or a different size. Finally records the synced assets version in the registry.

    Exit codes: 0 = success, 2 = no token configured, 1 = error.

.PARAMETER Owner   GitHub user/org (must match GH_OWNER in Updater.bas).
.PARAMETER Repo    Repo name        (must match GH_REPO).
.PARAMETER Branch  Branch/ref       (default main).
.PARAMETER Dest    The %APPDATA%\Microsoft\AddIns folder (manifest paths hang off this).
#>
param(
    [Parameter(Mandatory = $true)][string]$Owner,
    [Parameter(Mandatory = $true)][string]$Repo,
    [string]$Branch = 'main',
    [Parameter(Mandatory = $true)][string]$Dest
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$RegPath = 'HKCU:\Software\TrialQuest\Addin'

try {
    $token = (Get-ItemProperty -Path $RegPath -Name 'GitHubToken' -ErrorAction SilentlyContinue).GitHubToken
    if ([string]::IsNullOrWhiteSpace($token)) {
        Write-Host 'No update token configured; skipping asset download.'
        exit 2
    }

    $headers = @{
        Authorization          = "Bearer $token"
        Accept                 = 'application/vnd.github.raw'
        'User-Agent'           = 'TrialQuest-Installer'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    function Get-ContentsUrl([string]$path) {
        # Percent-encode each segment (handles spaces, &, +, ...) but keep the slashes.
        $enc = ($path -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
        "https://api.github.com/repos/$Owner/$Repo/contents/$enc`?ref=$Branch"
    }

    $manifest = (Invoke-WebRequest -Uri (Get-ContentsUrl 'manifest.json') -Headers $headers -UseBasicParsing).Content | ConvertFrom-Json

    $count = 0
    foreach ($f in $manifest.files) {
        $target = Join-Path $Dest ($f.path -replace '/', '\')
        $need = $true
        if (Test-Path -LiteralPath $target) {
            if ((Get-Item -LiteralPath $target).Length -eq [int64]$f.size) { $need = $false }
        }
        if ($need) {
            $dir = Split-Path $target -Parent
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Invoke-WebRequest -Uri (Get-ContentsUrl ('assets/' + $f.path)) -Headers $headers -OutFile $target -UseBasicParsing
            $count++
            Write-Host ("  + {0}" -f $f.path)
        }
    }

    if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
    Set-ItemProperty -Path $RegPath -Name 'AssetsVersion' -Value ([string]$manifest.assetsVersion)
    Write-Host ("Done: {0} file(s) downloaded; assets v{1}." -f $count, $manifest.assetsVersion)
    exit 0
}
catch {
    Write-Host ("ERROR: {0}" -f $_.Exception.Message)
    exit 1
}
