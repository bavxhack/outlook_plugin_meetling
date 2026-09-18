[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness
)

$ErrorActionPreference = 'Stop'
$registryView = if ($OutlookBitness -eq 'x86') { '32' } else { '64' }
$viewArgument = "/reg:$registryView"
$progId = 'Meetling.OutlookAddIn'
$classId = '{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}'
$keys = @(
    "HKCU\Software\Microsoft\Office\Outlook\Addins\$progId",
    "HKCU\Software\Classes\$progId",
    "HKCU\Software\Classes\CLSID\$classId"
)

foreach ($key in $keys) {
    & reg.exe DELETE $key /f $viewArgument 2>$null | Out-Null
    if ($LASTEXITCODE -notin @(0, 1)) {
        throw "Der Registrierungsschlüssel konnte nicht entfernt werden: $key"
    }
}

$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'Meetling wurde für den aktuellen Benutzer deinstalliert.'
