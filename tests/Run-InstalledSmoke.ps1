# Runs under the disposable standard user's profile, in the target architecture.
[CmdletBinding()]
param([Parameter(Mandatory = $true)][ValidateSet('x86', 'x64')][string] $Architecture)
$ErrorActionPreference = 'Stop'
$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Smoke tests must run without administrator rights.' }
$env:LOCALAPPDATA = [Environment]::GetFolderPath('LocalApplicationData')
$package = Join-Path $PSScriptRoot 'package'
$failed = $false
try {
    & "$package\Install.ps1" -OutlookBitness $Architecture
    & "$package\Install.ps1" -OutlookBitness $Architecture
    # Keep activation out of this long-lived process so it cannot lock the DLL during uninstall.
    $powershell = (Get-Process -Id $PID).Path
    & $powershell -NoProfile -ExecutionPolicy Bypass -File "$package\Diagnose.ps1" -OutlookBitness $Architecture -TestComActivation -OutputPath "$PSScriptRoot\COM-diagnosis.txt"
    if ($LASTEXITCODE -ne 0) { throw 'Isolated COM activation failed.' }
    & "$PSScriptRoot\harness\Meetling.OutlookAddIn.SmokeTests.exe"
    if ($LASTEXITCODE -ne 0) { throw 'IDispatch/Ribbon smoke failed.' }
    # This profile is disposable: verify damaged settings/secrets are never read or changed.
    $settingsRoot = Join-Path $env:LOCALAPPDATA 'Meetling\OutlookAddIn'
    $settings = Join-Path $settingsRoot 'settings.xml'
    $secret = Join-Path $settingsRoot 'api-key.bin'
    Set-Content -LiteralPath $settings -Value '<broken' -Encoding UTF8
    [IO.File]::WriteAllBytes($secret, [byte[]]@(1, 2, 3))
    $settingsHash = (Get-FileHash $settings).Hash
    $secretHash = (Get-FileHash $secret).Hash
    & "$PSScriptRoot\harness\Meetling.OutlookAddIn.SmokeTests.exe"
    if ($LASTEXITCODE -ne 0) { throw 'Corrupted-settings startup smoke failed.' }
    if ((Get-FileHash $settings).Hash -ne $settingsHash -or (Get-FileHash $secret).Hash -ne $secretHash) {
        throw 'Settings or secrets were changed.'
    }
} catch {
    $failed = $true
    Write-Output ($_ | Out-String)
} finally {
    & "$package\Uninstall.ps1" -OutlookBitness $Architecture
}
if (-not $failed) {
    if ((Get-FileHash $settings).Hash -ne $settingsHash -or (Get-FileHash $secret).Hash -ne $secretHash) {
        throw 'Uninstall changed settings or secrets.'
    }
    Write-Output 'PASS: missing/corrupt settings, secret preservation and uninstall.'
}
if ($failed) { exit 1 }
exit 0
