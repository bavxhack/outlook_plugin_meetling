[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('x86', 'x64')][string] $OutlookBitness,
    [string] $PackageDirectory = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Package-Tools.ps1')
Assert-InstallPrerequisites $OutlookBitness
$identity = Get-PackageIdentity $PackageDirectory $OutlookBitness
function Invoke-RegistryCommand {
    param([Parameter(Mandatory = $true)][string[]] $Arguments)

    & reg.exe @Arguments | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Die Registrierung ist mit Exitcode $LASTEXITCODE fehlgeschlagen."
    }
}


# Every runtime DLL is copied. No settings or secrets are read or removed.
$sourceDlls = @(Get-ChildItem -LiteralPath $PackageDirectory -Filter '*.dll' -File)
$sourceDlls | Unblock-File
# Versioned directories preserve the previous registration if a new copy fails.
$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness\$($identity.Version)"
New-Item -ItemType Directory -Force $target | Out-Null
foreach ($file in $sourceDlls) {
    $destination = Join-Path $target $file.Name
    if ([IO.Path]::GetFullPath($file.FullName) -ne [IO.Path]::GetFullPath($destination)) {
        Copy-Item -LiteralPath $file.FullName -Destination $destination -Force
    }
}
$installedIdentity = Get-PackageIdentity $target $OutlookBitness
if ($installedIdentity.FullName -ne $identity.FullName) { throw 'Installed assembly identity differs from package.' }
foreach ($file in $sourceDlls) {
    if ((Get-FileHash $file.FullName).Hash -ne (Get-FileHash (Join-Path $target $file.Name)).Hash) {
        throw "Copy verification failed: $($file.Name)"
    }
}
Get-ChildItem $target -Filter '*.dll' -File | Unblock-File
$addInPath = (Resolve-Path (Join-Path $target 'Meetling.OutlookAddIn.dll')).Path
$codeBase = [Uri]::new($addInPath).AbsoluteUri
$registryView = if ($OutlookBitness -eq 'x86') { '32' } else { '64' }
$progId = 'Meetling.OutlookAddIn'
$classId = '{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}'
$className = 'Meetling.OutlookAddIn.MeetlingAddIn'
$assemblyName = $identity.FullName
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
Invoke-RegistryCommand @('ADD', $addInKey, '/v', 'FriendlyName', '/t', 'REG_SZ', '/d', 'Meetling', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $addInKey, '/v', 'Description', '/t', 'REG_SZ', '/d', 'Lokale Meetling-Basisversion.', '/f', $viewArgument)
Invoke-RegistryCommand @('ADD', $addInKey, '/v', 'LoadBehavior', '/t', 'REG_DWORD', '/d', '3', '/f', $viewArgument)

Invoke-RegistryCommand @('QUERY', $inprocKey, '/v', 'CodeBase', $viewArgument)
Invoke-RegistryCommand @('QUERY', $addInKey, '/v', 'LoadBehavior', $viewArgument)


Write-Host "Registration completed for current user: $($identity.FullName), $OutlookBitness."
Write-Host 'Loading in Outlook has not been verified. Continue with docs/smoke-test.md.'
