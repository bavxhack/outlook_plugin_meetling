[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness,

    [string] $PackageDirectory = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$addInPath = Join-Path $PackageDirectory 'Meetling.OutlookAddIn.dll'
if (-not (Test-Path $addInPath)) {
    throw "Meetling.OutlookAddIn.dll wurde im Paketverzeichnis nicht gefunden: $PackageDirectory"
}

$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item (Join-Path $PackageDirectory '*.dll') $target -Force

$frameworkDirectory = if ($OutlookBitness -eq 'x86') {
    'Microsoft.NET\Framework\v4.0.30319'
} else {
    'Microsoft.NET\Framework64\v4.0.30319'
}
$regasm = Join-Path $env:WINDIR "$frameworkDirectory\RegAsm.exe"
& $regasm (Join-Path $target 'Meetling.OutlookAddIn.dll') /codebase /tlb | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Die COM-Registrierung ist mit Exitcode $LASTEXITCODE fehlgeschlagen."
}

$key = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\Meetling.OutlookAddIn'
New-Item $key -Force | Out-Null
New-ItemProperty $key -Name FriendlyName -Value 'Meetling-Konferenzen' -PropertyType String -Force | Out-Null
New-ItemProperty $key -Name Description -Value 'Erstellt Meetling-Konferenzen aus Outlook-Terminen.' -PropertyType String -Force | Out-Null
New-ItemProperty $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host 'Meetling wurde installiert. Outlook kann jetzt gestartet werden.'
