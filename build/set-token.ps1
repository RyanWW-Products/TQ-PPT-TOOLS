<#
.SYNOPSIS
    Securely store the GitHub update token on this machine and verify it works.

.DESCRIPTION
    Reads the read-only github_pat token from your CLIPBOARD (copy it from GitHub
    first), strips any stray whitespace/newlines, tests that it can read the repo,
    and only then saves it to HKCU\Software\TrialQuest\Addin\GitHubToken -- the
    same place the installer writes it and the add-in's Update button reads it.

    Usage:
        1. On GitHub, click the copy icon next to your new token.
        2. Run:  .\build\set-token.ps1
#>
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Owner   = 'RyanWW-Products'
$Repo    = 'TQ-PPT-TOOLS'
$Branch  = 'main'
$RegPath = 'HKCU:\Software\TrialQuest\Addin'

# Pull the token from the clipboard and remove ALL whitespace (spaces, tabs,
# CR/LF) -- GitHub PATs contain none, so this safely kills paste artifacts.
$raw = Get-Clipboard -Raw
$token = ($raw -replace '\s', '')
if ([string]::IsNullOrWhiteSpace($token)) {
    Write-Host 'Clipboard is empty. Copy your github_pat token from GitHub first, then re-run.' -ForegroundColor Yellow
    exit 1
}

# Sanity feedback (length + prefix only -- never the value).
$prefix = if ($token.Length -ge 11) { $token.Substring(0, 11) } else { $token }
Write-Host ("Read {0} characters from clipboard, starting '{1}...'" -f $token.Length, $prefix) -ForegroundColor Cyan
if ($token -notmatch '^(github_pat_|ghp_|ghs_)') {
    Write-Host "Warning: that doesn't look like a GitHub token. Continuing anyway to test it." -ForegroundColor Yellow
}

$headers = @{
    Authorization          = "Bearer $token"
    Accept                 = 'application/vnd.github.raw'
    'User-Agent'           = 'TQ-set-token'
    'X-GitHub-Api-Version' = '2022-11-28'
}

Write-Host "Testing token against $Owner/$Repo ..." -ForegroundColor Cyan
try {
    $vj = (Invoke-WebRequest "https://api.github.com/repos/$Owner/$Repo/contents/version.json?ref=$Branch" `
            -Headers $headers -UseBasicParsing).Content
    Write-Host 'SUCCESS - the repo returned:' -ForegroundColor Green
    Write-Host $vj
}
catch {
    Write-Host ("FAILED: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host 'Token NOT saved. Check: token resource owner = RyanWW-Products,' -ForegroundColor Red
    Write-Host 'repository access includes TQ-PPT-TOOLS, and Contents = Read-only.' -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
Set-ItemProperty -Path $RegPath -Name 'GitHubToken' -Value $token
Write-Host "Token saved to $RegPath\GitHubToken" -ForegroundColor Green
Write-Host 'Done. This machine can now download updates.' -ForegroundColor Green
