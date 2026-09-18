[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness,

    [string] $BuildRoot = (Join-Path $PSScriptRoot '..\src\Meetling.OutlookAddIn\bin')
)

$ErrorActionPreference = 'Stop'
$source = Join-Path $BuildRoot "$OutlookBitness\Release\net48"
$addIn = Join-Path $source 'Meetling.OutlookAddIn.dll'
if (-not (Test-Path $addIn)) {
    throw "Add-in nicht gefunden: $addIn. Zuerst Release|$OutlookBitness bauen."
}

$installerParameters = @{
    OutlookBitness = $OutlookBitness
    PackageDirectory = $source
}
& (Join-Path $PSScriptRoot 'Install-PublishedAddIn.ps1') @installerParameters
