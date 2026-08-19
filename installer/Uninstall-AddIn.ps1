[CmdletBinding()]
param([Parameter(Mandatory=$true)][ValidateSet('x86','x64')][string]$OutlookBitness)
$ErrorActionPreference = 'Stop'
$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
$dll = Join-Path $target 'Meetling.OutlookAddIn.dll'
$regasm = Join-Path $env:WINDIR $(if ($OutlookBitness -eq 'x86') {'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'} else {'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'})
if (Test-Path $dll) { & $regasm $dll /unregister | Out-Null }
Remove-Item 'HKCU:\Software\Microsoft\Office\Outlook\Addins\Meetling.OutlookAddIn' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'Meetling-Add-in wurde deinstalliert.'
