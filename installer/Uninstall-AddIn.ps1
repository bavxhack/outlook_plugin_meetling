[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness
)

$ErrorActionPreference = 'Stop'
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close Outlook completely first.' }
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
# Remove only known shipped binaries, preserving any settings and unrelated files.
if (Test-Path $target) {
    Get-ChildItem $target -Recurse -File |
        Where-Object { $_.Name -in @('Meetling.OutlookAddIn.dll', 'Meetling.Core.dll', 'Microsoft.Office.Interop.Outlook.dll') } |
        Remove-Item -Force
}
Write-Host 'Meetling wurde für den aktuellen Benutzer deinstalliert.'

