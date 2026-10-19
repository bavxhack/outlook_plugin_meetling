# CI driver only. Never included in installation packages.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('x86', 'x64')][string] $Architecture,
    [Parameter(Mandatory = $true)][string] $PackageDirectory,
    [Parameter(Mandatory = $true)][string] $HarnessDirectory
)
$ErrorActionPreference = 'Stop'
if ($env:GITHUB_ACTIONS -ne 'true') { throw 'This disposable-account driver is restricted to GitHub Actions.' }
$userName = 'MtSmoke' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$password = ConvertTo-SecureString ([Guid]::NewGuid().ToString('N') + '!aA1') -AsPlainText -Force
$root = Join-Path $env:PUBLIC $userName
try {
    New-LocalUser -Name $userName -Password $password -AccountNeverExpires | Out-Null
    New-Item -ItemType Directory -Path "$root\package", "$root\harness" -Force | Out-Null
    Copy-Item "$PackageDirectory\*" "$root\package" -Recurse -Force
    Copy-Item "$HarnessDirectory\*" "$root\harness" -Recurse -Force
    Copy-Item (Join-Path $PSScriptRoot 'Run-InstalledSmoke.ps1') "$root\Run.ps1"
    # Grant access only to this disposable test directory and this test account.
    & icacls.exe $root /grant ("${userName}:(OI)(CI)M") /T /Q | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Test directory ACL setup failed.' }
    $credential = [Management.Automation.PSCredential]::new("$env:COMPUTERNAME\$userName", $password)
    $ps = if ($Architecture -eq 'x86') { "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" } else { "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" }
    $child = Start-Process -FilePath $ps -Credential $credential -LoadUserProfile -Wait -PassThru -WorkingDirectory $root -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$root\Run.ps1`"", '-Architecture', $Architecture
    ) -RedirectStandardOutput "$root\stdout.log" -RedirectStandardError "$root\stderr.log"
    Get-Content "$root\stdout.log"
    Get-Content "$root\stderr.log"
    if (Test-Path "$root\COM-diagnosis.txt") { Get-Content "$root\COM-diagnosis.txt" }
    if ($child.ExitCode -ne 0) { throw "Standard-user COM smoke test failed: $($child.ExitCode)" }
} finally {
    Remove-LocalUser -Name $userName -ErrorAction SilentlyContinue
    if (Test-Path $root) { Remove-Item $root -Recurse -Force }
}
