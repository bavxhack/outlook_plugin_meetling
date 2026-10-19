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
    & "$package\Diagnose.ps1" -OutlookBitness $Architecture -TestComActivation -OutputPath "$PSScriptRoot\COM-diagnosis.txt"
    & "$PSScriptRoot\harness\Meetling.OutlookAddIn.SmokeTests.exe"
    if ($LASTEXITCODE -ne 0) { throw 'IDispatch/Ribbon smoke failed.' }
} catch {
    $failed = $true
    Write-Output ($_ | Out-String)
} finally {
    & "$package\Uninstall.ps1" -OutlookBitness $Architecture
}
if ($failed) { exit 1 }
exit 0
