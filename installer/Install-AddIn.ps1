[CmdletBinding()]
param([Parameter(Mandatory=$true)][ValidateSet('x86','x64')][string]$OutlookBitness,
      [string]$BuildRoot = (Join-Path $PSScriptRoot '..\src\Meetling.OutlookAddIn\bin\Release'))
$ErrorActionPreference = 'Stop'
$platform = if ($OutlookBitness -eq 'x86') { 'x86' } else { 'x64' }
$source = Join-Path $BuildRoot "$platform\net48"
$dll = Join-Path $source 'Meetling.OutlookAddIn.dll'
if (-not (Test-Path $dll)) { throw "Add-in nicht gefunden: $dll. Zuerst Release|$platform bauen." }
$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$platform"
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item (Join-Path $source '*') $target -Recurse -Force
$regasm = Join-Path $env:WINDIR $(if ($OutlookBitness -eq 'x86') {'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'} else {'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'})
& $regasm (Join-Path $target 'Meetling.OutlookAddIn.dll') /codebase /tlb | Out-Null
$key = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\Meetling.OutlookAddIn'
New-Item $key -Force | Out-Null
New-ItemProperty $key -Name FriendlyName -Value 'Meetling-Konferenzen' -PropertyType String -Force | Out-Null
New-ItemProperty $key -Name Description -Value 'Erstellt Meetling-Konferenzen aus Outlook-Terminen.' -PropertyType String -Force | Out-Null
New-ItemProperty $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
Write-Host 'Meetling-Add-in wurde für den aktuellen Benutzer installiert. Outlook jetzt starten.'
