[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness,

    [string] $PackageDirectory = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$sourceAddIn = Join-Path $PackageDirectory 'Meetling.OutlookAddIn.dll'
if (-not (Test-Path $sourceAddIn)) {
    throw "Meetling.OutlookAddIn.dll wurde im Paketverzeichnis nicht gefunden: $PackageDirectory"
}

function Invoke-RegistryCommand {
    param([Parameter(Mandatory = $true)][string[]] $Arguments)

    & reg.exe @Arguments | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Die Registrierung ist mit Exitcode $LASTEXITCODE fehlgeschlagen."
    }
}

$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item (Join-Path $PackageDirectory '*.dll') $target -Force

$addInPath = (Resolve-Path (Join-Path $target 'Meetling.OutlookAddIn.dll')).Path
$codeBase = [Uri]::new($addInPath).AbsoluteUri
$registryView = if ($OutlookBitness -eq 'x86') { '32' } else { '64' }
$progId = 'Meetling.OutlookAddIn'
$classId = '{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}'
$className = 'Meetling.OutlookAddIn.MeetlingAddIn'
$assemblyName = 'Meetling.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$classesRoot = 'HKCU\Software\Classes'
$classKey = "$classesRoot\CLSID\$classId"
$inprocKey = "$classKey\InprocServer32"
$addInKey = "HKCU\Software\Microsoft\Office\Outlook\Addins\$progId"
$viewArgument = "/reg:$registryView"

Invoke-RegistryCommand @('ADD', "$classesRoot\$progId", '/ve', '/d', 'Meetling Outlook Add-in', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', "$classesRoot\$progId\CLSID", '/ve', '/d', $classId, '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $classKey, '/ve', '/d', $className, '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', "$classKey\ProgId", '/ve', '/d', $progId, '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $inprocKey, '/ve', '/d', 'mscoree.dll', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $inprocKey, '/v', 'ThreadingModel', '/t', 'REG_SZ', '/d', 'Both', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $inprocKey, '/v', 'Class', '/t', 'REG_SZ', '/d', $className, '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $inprocKey, '/v', 'Assembly', '/t', 'REG_SZ', '/d', $assemblyName, '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $inprocKey, '/v', 'RuntimeVersion', '/t', 'REG_SZ', '/d', 'v4.0.30319', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $inprocKey, '/v', 'CodeBase', '/t', 'REG_SZ', '/d', $codeBase, '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $addInKey, '/v', 'FriendlyName', '/t', 'REG_SZ', '/d', 'Meetling-Konferenzen', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $addInKey, '/v', 'Description', '/t', 'REG_SZ', '/d', 'Erstellt Meetling-Konferenzen aus Outlook-Terminen.', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $addInKey, '/v', 'LoadBehavior', '/t', 'REG_DWORD', '/d', '3', '/f', $viewArgument)

Write-Host 'Meetling wurde für den aktuellen Benutzer installiert. Outlook kann jetzt gestartet werden.'
