[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness
)

$ErrorActionPreference = 'Stop'
$registryView = if ($OutlookBitness -eq 'x86') {
    [Microsoft.Win32.RegistryView]::Registry32
} else {
    [Microsoft.Win32.RegistryView]::Registry64
}
$progId = 'Meetling.OutlookAddIn'
$classId = '{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}'
$keys = @(
    "Software\Microsoft\Office\Outlook\Addins\$progId",
    "Software\Classes\$progId",
    "Software\Classes\CLSID\$classId"
)

$currentUser = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    $registryView
)
try {
    foreach ($key in $keys) {
        $currentUser.DeleteSubKeyTree($key, $false)
    }
}
finally {
    $currentUser.Dispose()
}

$target = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
Write-Host 'Meetling wurde für den aktuellen Benutzer deinstalliert.'
