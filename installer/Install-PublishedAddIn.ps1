[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness,

    [string] $PackageDirectory = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$requiredFiles = @(
    'Meetling.OutlookAddIn.dll',
    'Meetling.Core.dll',
    'Microsoft.Office.Interop.Outlook.dll'
)
foreach ($file in $requiredFiles) {
    if (-not (Test-Path (Join-Path $PackageDirectory $file))) {
        throw "Erforderliche Datei fehlt im Paket: $file. Bitte das vollständige ZIP erneut herunterladen und entpacken."
    }
}

function Invoke-RegistryCommand {
    param([Parameter(Mandatory = $true)][string[]] $Arguments)

    & reg.exe @Arguments | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Die Registrierung ist mit Exitcode $LASTEXITCODE fehlgeschlagen."
    }
}

function Get-InstalledOutlookBitness {
    $locations = @(
        @{ Path = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Value = 'Platform' },
        @{ Path = 'SOFTWARE\Microsoft\Office\16.0\Outlook'; Value = 'Bitness' }
    )
    $views = @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )

    foreach ($view in $views) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        try {
            foreach ($location in $locations) {
                $key = $baseKey.OpenSubKey($location.Path)
                try {
                    if ($null -eq $key) {
                        continue
                    }

                    $registryValue = $key.GetValue($location.Value)
                    $value = if ($null -eq $registryValue) { $null } else { $registryValue.ToString().ToLowerInvariant() }
                    if ($value -in @('x86', 'x64')) {
                        return $value
                    }
                }
                finally {
                    if ($null -ne $key) {
                        $key.Dispose()
                    }
                }
            }
        }
        finally {
            $baseKey.Dispose()
        }
    }

    return $null
}

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    throw 'Outlook läuft noch. Bitte Outlook vollständig schließen und die Installation erneut starten.'
}

$installedBitness = Get-InstalledOutlookBitness
if ($null -ne $installedBitness -and $installedBitness -ne $OutlookBitness) {
    throw "Installationspaket und Outlook passen nicht zusammen. Installiert ist Outlook $installedBitness, ausgewählt wurde $OutlookBitness."
}

$frameworkRelease = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue).Release
if ($null -eq $frameworkRelease -or $frameworkRelease -lt 528040) {
    throw '.NET Framework 4.8 ist nicht installiert. Bitte .NET Framework 4.8 installieren und die Installation wiederholen.'
}

$sourceDlls = Get-ChildItem $PackageDirectory -Filter '*.dll' -File
$sourceDlls | Unblock-File

$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
New-Item -ItemType Directory -Force $target | Out-Null
$sourceDlls | Copy-Item -Destination $target -Force
Get-ChildItem $target -Filter '*.dll' -File | Unblock-File

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

Invoke-RegistryCommand @('QUERY', $inprocKey, '/v', 'CodeBase', $viewArgument)
Invoke-RegistryCommand @('QUERY', $addInKey, '/v', 'LoadBehavior', $viewArgument)

Write-Host 'Meetling wurde für den aktuellen Benutzer installiert. Outlook kann jetzt gestartet werden.'
