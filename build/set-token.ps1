<#
.SYNOPSIS
    Securely store the GitHub update token on this machine and verify it works.

.DESCRIPTION
    Prompts for the read-only github_pat token (input hidden, never echoed and not
    stored in command history), tests that it can read the repo, and only then
    saves it to HKCU\Software\TrialQuest\Addin\GitHubToken -- the same place the
    installer writes it and the add-in's Update button reads it.

    Run it any time you need to set or rotate the token on a machine:
        .\build\set-token.ps1
#>
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Owner   = 'RyanWW-Products'
$Repo    = 'TQ-PPT-TOOLS'
$Branch  = 'main'
$RegPath = 'HKCU:\Software\TrialQuest\Addin'

$sec = Read-Host 'Paste your github_pat token (it will stay hidden)' -AsSecureString
$token = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
if ([string]::IsNullOrWhiteSpace($token)) {
    Write-Host 'No token entered. Nothing saved.' -ForegroundColor Yellow
    exit 1
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
