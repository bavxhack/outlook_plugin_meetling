# Shared, read-only package and architecture checks. Windows PowerShell 5.1 compatible.
Set-StrictMode -Version 2.0

function Get-PeArchitecture([string] $Path) {
    $stream = [IO.File]::OpenRead($Path)
    $reader = [IO.BinaryReader]::new($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw "Invalid PE: $Path" }
        $stream.Position = 0x3c
        $stream.Position = $reader.ReadInt32()
        if ($reader.ReadUInt32() -ne 0x4550) { throw "Invalid PE signature: $Path" }
        switch ($reader.ReadUInt16()) {
            0x14c { return 'x86' }
            0x8664 { return 'x64' }
            default { throw "Unsupported PE architecture: $Path" }
        }
    } finally { $reader.Dispose(); $stream.Dispose() }
}

function Get-PackageIdentity([string] $Directory, [string] $Architecture) {
    foreach ($name in @('Meetling.OutlookAddIn.dll', 'Meetling.Core.dll', 'Microsoft.Office.Interop.Outlook.dll')) {
        $file = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required runtime file missing: $file" }
        # GetAssemblyName reads metadata without loading dependencies into this process.
        $identity = [Reflection.AssemblyName]::GetAssemblyName($file)
        if ($identity.Name -ne [IO.Path]::GetFileNameWithoutExtension($name)) { throw "Unexpected assembly: $file" }
    }
    $dll = Join-Path $Directory 'Meetling.OutlookAddIn.dll'
    if ((Get-PeArchitecture $dll) -ne $Architecture) { throw "Package architecture does not match $Architecture" }
    $identity = [Reflection.AssemblyName]::GetAssemblyName($dll)
    # x86 PE alone could be AnyCPU. Require the compiler's explicit target flag.
    $expected = if ($Architecture -eq 'x86') { 'X86' } else { 'Amd64' }
    if ($identity.ProcessorArchitecture.ToString() -ne $expected) { throw "Assembly is not explicitly built for $Architecture" }
    return $identity
}

function Get-OutlookArchitecture {
    # Prefer the actual registered OUTLOOK.EXE over a generic Office platform value.
    $candidates = @()
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
            $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            try {
                $key = $root.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE')
                if ($null -ne $key) {
                    try {
                        $path = ([string]$key.GetValue('')).Trim().Trim('"').Trim()
                        # App Paths may exist without a usable default value.
                        # Test-Path rejects empty strings instead of returning false.
                        if (-not [string]::IsNullOrWhiteSpace($path) -and
                            (Test-Path -LiteralPath $path -PathType Leaf)) {
                            $candidates += Get-PeArchitecture $path
                        }
                    } finally { $key.Dispose() }
                }
            } finally { $root.Dispose() }
        }
    }
    $unique = @($candidates | Select-Object -Unique)
    if ($unique.Count -gt 1) { throw 'Conflicting Outlook installations. Resolve App Paths before installation.' }
    if ($unique.Count -eq 1) { return $unique[0] }
    # Outlook-specific Bitness values only; generic ClickToRun could describe Office without Outlook.
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        try {
            foreach ($version in @('16.0', '15.0', '14.0')) {
                $key = $root.OpenSubKey("Software\Microsoft\Office\$version\Outlook")
                if ($null -ne $key) {
                    try {
                        $value = [string]$key.GetValue('Bitness')
                        if ($value -in @('x86', 'x64')) { $candidates += $value }
                    } finally { $key.Dispose() }
                }
            }
        } finally { $root.Dispose() }
    }
    $unique = @($candidates | Select-Object -Unique)
    if ($unique.Count -gt 1) { throw 'Conflicting Outlook bitness registry values.' }
    if ($unique.Count -eq 1) { return $unique[0] }
    return $null
}

function Assert-InstallPrerequisites([string] $Architecture) {
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close Outlook completely first.' }
    if ($Architecture -eq 'x64' -and -not [Environment]::Is64BitOperatingSystem) { throw 'x64 requires 64-bit Windows.' }
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
    try {
        $key = $root.OpenSubKey('SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full')
        try {
            if ($null -eq $key -or [int]$key.GetValue('Release', 0) -lt 528040) { throw '.NET Framework 4.8 required.' }
        } finally { if ($null -ne $key) { $key.Dispose() } }
    } finally { $root.Dispose() }
    $actual = Get-OutlookArchitecture
    if ($actual -and $actual -ne $Architecture) { throw "Outlook is $actual, package is $Architecture." }
    if (-not $actual) { Write-Warning 'Outlook bitness could not be determined. Verify in Outlook: File > Office Account > About Outlook.' }
}
