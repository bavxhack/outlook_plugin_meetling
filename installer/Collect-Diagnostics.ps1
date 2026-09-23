[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x86', 'x64')]
    [string] $OutlookBitness,

    [string] $OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Meetling-Outlook-Diagnose.txt')
)

$ErrorActionPreference = 'Stop'
$registryView = if ($OutlookBitness -eq 'x86') {
    [Microsoft.Win32.RegistryView]::Registry32
} else {
    [Microsoft.Win32.RegistryView]::Registry64
}
$installDirectory = Join-Path $env:LOCALAPPDATA "Meetling\OutlookAddIn\$OutlookBitness"
$lines = [System.Collections.Generic.List[string]]::new()

function Add-Section {
    param([Parameter(Mandatory = $true)][string] $Title)
    $lines.Add('')
    $lines.Add("=== $Title ===")
}

function Add-RegistryKey {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Win32.RegistryKey] $Root,
        [Parameter(Mandatory = $true)][string] $Path
    )

    $lines.Add("[$Path]")
    $key = $Root.OpenSubKey($Path)
    if ($null -eq $key) {
        $lines.Add('<nicht vorhanden>')
        return
    }

    try {
        foreach ($name in $key.GetValueNames()) {
            $displayName = if ([string]::IsNullOrEmpty($name)) { '(Standard)' } else { $name }
            $value = $key.GetValue($name)
            if ($value -is [byte[]]) {
                $lines.Add("$displayName = <Binärwert, $($value.Length) Bytes>")
            } else {
                $lines.Add("$displayName = $value")
            }
        }
    }
    finally {
        $key.Dispose()
    }
}

$lines.Add("Meetling Outlook Diagnose – $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss K')")
$lines.Add("Computer: $env:COMPUTERNAME")
$lines.Add("Windows: $([Environment]::OSVersion.VersionString)")
$lines.Add("PowerShell: $($PSVersionTable.PSVersion)")
$lines.Add("Gewählte Outlook-Bitness: $OutlookBitness")
$frameworkRelease = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue).Release
$lines.Add(".NET-Framework-Release: $frameworkRelease")

Add-Section 'Installierte Dateien'
if (Test-Path $installDirectory) {
    foreach ($file in Get-ChildItem $installDirectory -File | Sort-Object Name) {
        $hash = if ($file.Extension -ieq '.dll') { (Get-FileHash $file.FullName -Algorithm SHA256).Hash } else { '-' }
        $zoneBlocked = $null -ne (Get-Item $file.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue)
        $lines.Add("$($file.Name) | $($file.Length) Bytes | SHA256=$hash | Windows-blockiert=$zoneBlocked")
    }
} else {
    $lines.Add("Installationsordner fehlt: $installDirectory")
}

Add-Section 'Meetling-Registry'
$currentUser = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
try {
    Add-RegistryKey $currentUser 'Software\Microsoft\Office\Outlook\Addins\Meetling.OutlookAddIn'
    Add-RegistryKey $currentUser 'Software\Classes\CLSID\{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}\InprocServer32'

    $lines.Add('[Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems]')
    $disabledItems = $currentUser.OpenSubKey('Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems')
    try {
        $disabledCount = if ($null -eq $disabledItems) { 0 } else { $disabledItems.ValueCount }
        $lines.Add("Anzahl deaktivierter Einträge: $disabledCount (Inhalte werden aus Datenschutzgründen nicht ausgegeben)")
    }
    finally {
        if ($null -ne $disabledItems) { $disabledItems.Dispose() }
    }
}
finally {
    $currentUser.Dispose()
}

Add-Section 'Meetling-Diagnoseprotokoll'
$diagnosticLog = Join-Path $env:LOCALAPPDATA 'Meetling\OutlookAddIn\diagnose.log'
if (Test-Path $diagnosticLog) {
    $lines.AddRange([string[]](Get-Content $diagnosticLog -Tail 100))
} else {
    $lines.Add('<noch nicht angelegt; das Add-in wurde vermutlich nicht bis OnConnection geladen>')
}

Add-Section 'Relevante Windows-Ereignisse der letzten drei Tage'
try {
    $events = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = (Get-Date).AddDays(-3) } -MaxEvents 500 -ErrorAction Stop |
        Where-Object {
            $_.ProviderName -in @('.NET Runtime', 'Application Error', 'Application Hang', 'SideBySide') -and
            ($_.Message -match 'OUTLOOK|Meetling|mscoree')
        } |
        Select-Object -First 50
    if ($null -eq $events) {
        $lines.Add('<keine passenden Ereignisse gefunden>')
    } else {
        foreach ($event in $events) {
            $message = $event.Message -replace "`r?`n", ' '
            $lines.Add("$($event.TimeCreated.ToString('o')) | $($event.ProviderName) | ID $($event.Id) | $message")
        }
    }
}
catch {
    $lines.Add("Ereignisprotokoll konnte nicht gelesen werden: $($_.Exception.Message)")
}

$parent = Split-Path $OutputPath -Parent
if (-not [string]::IsNullOrEmpty($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$lines | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Host "Diagnose wurde erstellt: $OutputPath"
