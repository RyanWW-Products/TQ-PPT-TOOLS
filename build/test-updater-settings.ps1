<# Reproduces the old registry-reset bug and exercises the registry commands
   emitted by Updater.bas using only disposable keys and a dummy token. #>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$testKey = 'HKCU:\Software\TrialQuestTests\' + [Guid]::NewGuid().ToString('N')
$sentinels = @{
    GitHubToken = 'dummy-token-for-regression-only'
    Channel = 'beta'
    AssetsVersion = '2'
    AutoCheck = '1'
    CustomSetting = 'preserve-me'
}
function Seed-Settings {
    if (-not (Test-Path -LiteralPath $testKey)) { New-Item -Path $testKey -Force | Out-Null }
    foreach ($name in $sentinels.Keys) { Set-ItemProperty -LiteralPath $testKey -Name $name -Value $sentinels[$name] }
    Set-ItemProperty -LiteralPath $testKey -Name InstalledVersion -Value '5.6.7'
}
try {
    Seed-Settings
    # This is the old swapper command, confined to the unique disposable key.
    New-Item -Path $testKey -Force | Out-Null
    if ((Get-Item -LiteralPath $testKey).GetValue('GitHubToken')) { throw 'Old bug did not reproduce.' }
    Write-Output 'PASS | old New-Item -Force command erases the dummy token'
    Seed-Settings

    $source = [IO.File]::ReadAllText((Join-Path $repoRoot 'Modules/Updater.bas'))
    $start = $source.IndexOf('    ts.WriteLine "if ($copied) {"')
    $end = $source.IndexOf('    ts.WriteLine "Remove-Item', $start)
    if ($start -lt 0 -or $end -lt 0) { throw 'Cannot locate generated swapper registry block.' }
    $lines = foreach ($line in ($source.Substring($start, $end - $start) -split '\r?\n')) {
        if ($line -match '^\s*ts\.WriteLine "(.*)"\s*$') {
            $Matches[1].Replace('" & newVersion & "', '5.6.8').Replace('HKCU:\Software\TrialQuest\Addin', $testKey)
        }
    }
    $script = [scriptblock]::Create($lines -join "`n")
    $copied = $true
    & $script
    $settings = Get-Item -LiteralPath $testKey
    foreach ($name in $sentinels.Keys) {
        if ($settings.GetValue($name) -cne $sentinels[$name]) { throw "Update changed $name." }
    }
    if ($settings.GetValue('InstalledVersion') -ne '5.6.8') { throw 'Version was not updated.' }
    Write-Output 'PASS | update preserves token, channel, asset version, auto-check and unrelated settings'

    $copied = $false
    Set-ItemProperty -LiteralPath $testKey -Name InstalledVersion -Value '5.6.7'
    & $script
    if ((Get-Item -LiteralPath $testKey).GetValue('InstalledVersion') -ne '5.6.7') { throw 'Unapplied update changed version.' }
    Write-Output 'PASS | unapplied update leaves the installed version alone'

    Remove-Item -LiteralPath $testKey
    $copied = $true
    & $script
    if ((Get-Item -LiteralPath $testKey).GetValue('InstalledVersion') -ne '5.6.8') { throw 'Missing settings key was not initialized.' }
    Write-Output 'PASS | update initializes a missing settings key'
} finally {
    if (Test-Path -LiteralPath $testKey) { Remove-Item -LiteralPath $testKey }
}
