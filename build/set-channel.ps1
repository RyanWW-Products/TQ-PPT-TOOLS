<#
.SYNOPSIS
    Switch this machine's Trial Quest update channel between stable and beta.

.DESCRIPTION
    Sets HKCU\Software\TrialQuest\Addin\Channel. The add-in's Update button then
    pulls from the matching branch on the next check: stable -> main, beta -> beta.
    Switching does not download anything by itself -- click Trial Quest >
    Check for Updates afterwards to pull from the new channel.

.PARAMETER Channel
    'stable' (live users) or 'beta' (you / beta testers).

.EXAMPLE
    .\build\set-channel.ps1 beta
#>
param(
    [Parameter(Mandatory = $true)][ValidateSet('stable', 'beta')][string]$Channel
)
$ErrorActionPreference = 'Stop'
$RegPath = 'HKCU:\Software\TrialQuest\Addin'
if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
Set-ItemProperty -Path $RegPath -Name 'Channel' -Value $Channel
Write-Host ("Channel set to '{0}'." -f $Channel) -ForegroundColor Green
Write-Host "Open PowerPoint > Trial Quest > Check for Updates to pull from it." -ForegroundColor Green
