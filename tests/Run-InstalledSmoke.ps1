# Runs under the disposable standard user's profile, in the target architecture.
[CmdletBinding()]
param([Parameter(Mandatory = $true)][ValidateSet('x86', 'x64')][string] $Architecture)
$ErrorActionPreference = 'Stop'
$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Smoke tests must run without administrator rights.' }
$env:LOCALAPPDATA = [Environment]::GetFolderPath('LocalApplicationData')
$package = Join-Path $PSScriptRoot 'package (Outlook test)'
$failed = $false
$appPathsKey = 'Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
$registryView = if ($Architecture -eq 'x86') { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$fixtureRoot = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
$fixtureCreated = $false
try {
    & "$package\Install.ps1" -OutlookBitness $Architecture
    # Modify only this disposable profile; never overwrite an existing Outlook entry.
    $existingKey = $fixtureRoot.OpenSubKey($appPathsKey)
    if ($null -ne $existingKey) { $existingKey.Dispose(); throw 'Unexpected Outlook fixture key already exists.' }
    $fixture = $fixtureRoot.CreateSubKey($appPathsKey)
    $fixtureCreated = $true
    try {
        foreach ($emptyValue in @('', '   ', '""')) {
            $fixture.SetValue('', $emptyValue)
            & "$package\Install.ps1" -OutlookBitness $Architecture
        }
        . (Join-Path $package 'Package-Tools.ps1')
        # A valid quoted PE path must still identify the process architecture.
        $fixture.SetValue('', '"' + (Get-Process -Id $PID).Path + '"')
        if ((Get-OutlookArchitecture) -ne $Architecture) { throw 'Valid App Paths architecture detection regressed.' }
    } finally { $fixture.Dispose() }
    $fixtureRoot.DeleteSubKeyTree($appPathsKey)
    $fixtureCreated = $false
    Write-Output 'PASS: empty Outlook App Paths values and package path with spaces/parentheses.'

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
    try {
        if ($fixtureCreated) { $fixtureRoot.DeleteSubKeyTree($appPathsKey) }
    } finally { $fixtureRoot.Dispose() }
    & "$package\Uninstall.ps1" -OutlookBitness $Architecture
}
if (-not $failed) {
    if ((Get-FileHash $settings).Hash -ne $settingsHash -or (Get-FileHash $secret).Hash -ne $secretHash) {
        throw 'Uninstall changed settings or secrets.'
    }
    Write-Output 'PASS: missing/corrupt settings, secret preservation and uninstall.'
}
if ($failed) { exit 1 }
# Exercise the actual MSI in the same unprivileged profile after ZIP cleanup.
function Invoke-Msi([string[]] $MsiArguments) {
    $process = Start-Process "$env:WINDIR\System32\msiexec.exe" -ArgumentList $MsiArguments -Wait -PassThru
    if ($process.ExitCode -notin @(0, 3010)) {
        Get-Content "$PSScriptRoot\msi.log" -ErrorAction SilentlyContinue
        throw "MSI failed: $($process.ExitCode)"
    }
}
$msi = "`"$PSScriptRoot\Meetling.msi`""
$log = "`"$PSScriptRoot\msi.log`""
$installedDll = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$Architecture\Meetling.OutlookAddIn.dll"
$msiInstalled = $false
try {
    & "$PSScriptRoot\Test-MsiUserInterface.ps1" -MsiPath "$PSScriptRoot\Meetling.msi"
    $msiInstalled = $true
    if (-not (Test-Path $installedDll)) { throw 'MSI payload missing.' }
    # Repair must restore a deleted file and preserve user data.
    Remove-Item $installedDll -Force
    Invoke-Msi @('/fa', $msi, '/qn', '/norestart', '/l*v', $log)
    if (-not (Test-Path $installedDll)) { throw 'MSI repair failed.' }
    & $powershell -NoProfile -ExecutionPolicy Bypass -File "$package\Diagnose.ps1" -OutlookBitness $Architecture -TestComActivation -OutputPath "$PSScriptRoot\COM-diagnosis.txt"
    if ($LASTEXITCODE -ne 0) { throw 'MSI COM activation failed.' }
    & "$PSScriptRoot\harness\Meetling.OutlookAddIn.SmokeTests.exe"
    if ($LASTEXITCODE -ne 0) { throw 'MSI IDispatch/Ribbon smoke failed.' }
} finally {
    if ($msiInstalled) { Invoke-Msi @('/x', $msi, '/qn', '/norestart', '/l*v', $log) }
}
if (Test-Path $installedDll) { throw 'MSI uninstall left payload behind.' }
$root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
try {
    foreach ($path in @('Software\Microsoft\Office\Outlook\Addins\Meetling.OutlookAddIn', 'Software\Classes\CLSID\{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}')) {
        $key = $root.OpenSubKey($path)
        if ($key) { $key.Dispose(); throw "MSI uninstall left registration: $path" }
    }
} finally { $root.Dispose() }
if ((Get-FileHash $settings).Hash -ne $settingsHash -or (Get-FileHash $secret).Hash -ne $secretHash) {
    throw 'MSI changed settings or secrets.'
}
Write-Output 'PASS: standard-user MSI installation, repair, COM/Ribbon activation, uninstall and user-data preservation.'
exit 0
