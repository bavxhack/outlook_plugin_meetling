[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('x86', 'x64')][string] $OutlookBitness,
    [string] $OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Meetling-Outlook-Diagnose.txt'),
    [switch] $TestComActivation
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Package-Tools.ps1')
$view = if ($OutlookBitness -eq 'x86') { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
$lines = [Collections.Generic.List[string]]::new()
$lines.Add("Meetling diagnosis $(Get-Date -Format o)")
$lines.Add("Windows: $([Environment]::OSVersion.VersionString)")
$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
$lines.Add("Administrator token: $($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))")
$lines.Add("PowerShell: $($PSVersionTable.PSVersion); process x64: $([Environment]::Is64BitProcess); target: $OutlookBitness")
try { $lines.Add("Detected Outlook: $(Get-OutlookArchitecture)") } catch { $lines.Add("Outlook architecture: $($_.Exception.GetType().FullName)") }
$path = $null
$registeredAssembly = $null
$activationFailed = $false
foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
    try {
        foreach ($name in @(
            'Software\Microsoft\Office\Outlook\Addins\Meetling.OutlookAddIn',
            'Software\Classes\Meetling.OutlookAddIn\CLSID',
            'Software\Classes\CLSID\{4BB84EF2-9D83-4710-BDBA-7CBFD6939B48}\InprocServer32'
        )) {
            $key = $root.OpenSubKey($name)
            $lines.Add("[$hive / $view / $name]")
            if ($null -eq $key) { $lines.Add('<missing>'); continue }
            try {
                foreach ($value in $key.GetValueNames()) { $lines.Add("$value = $($key.GetValue($value))") }
                if ($name.EndsWith('InprocServer32') -and -not $path) {
                    $codeBase = [string]$key.GetValue('CodeBase')
                    if ($codeBase) {
                        try { $path = ([Uri]$codeBase).LocalPath }
                        catch { $lines.Add('Invalid CodeBase URI') }
                    }
                    $registeredAssembly = [string]$key.GetValue('Assembly')
                }
            } finally { $key.Dispose() }
        }
    } finally { $root.Dispose() }
}
if (-not $path) { $lines.Add('COM registration missing or CodeBase unset.') }
elseif (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $lines.Add("Registered assembly missing: $path") }
else {
    try {
        $identity = Get-PackageIdentity (Split-Path $path -Parent) $OutlookBitness
        $lines.Add("Package readable and architecture correct: $($identity.FullName)")
        if ($identity.FullName -ne $registeredAssembly) { $lines.Add('ERROR: Registry Assembly identity differs from installed DLL.') }
    } catch { $lines.Add("Package/architecture/metadata check failed: $($_.Exception.Message)") }
}
if ($TestComActivation) {
    $processArchitecture = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
    if ($PSVersionTable.PSEdition -eq 'Core' -or $processArchitecture -ne $OutlookBitness) {
        throw 'COM test requires Windows PowerShell 5.1 with target bitness. See README.'
    }
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close Outlook before the optional COM test.' }
    $instance = $null
    try {
        $type = [Type]::GetTypeFromProgID('Meetling.OutlookAddIn', $true)
        $instance = [Activator]::CreateInstance($type)
        $lines.Add('COM activation succeeded. No lifecycle or button methods invoked. This does not prove loading in Outlook.')
    } catch {
        $activationFailed = $true
        $failure = $_.Exception
        while ($null -ne $failure) {
            $lines.Add("COM activation failed: $($failure.GetType().FullName) HRESULT=$('{0:X8}' -f $failure.HResult)")
            if ($failure -is [IO.FileNotFoundException] -or $failure -is [IO.FileLoadException]) {
                $lines.Add("AssemblyFile = $($failure.FileName)")
            }
            $failure = $failure.InnerException
        }
    } finally {
        if ($null -ne $instance -and [Runtime.InteropServices.Marshal]::IsComObject($instance)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($instance)
        }
    }
}
$log = Join-Path $env:LOCALAPPDATA 'Meetling\OutlookAddIn\diagnose.log'
$lines.Add('[Recent lifecycle log; compare timestamps with your Outlook run]')
if (Test-Path $log) { $lines.AddRange([string[]](Get-Content $log -Tail 100)) }
else { $lines.Add('No log: does not prove activation failure; directory may be unwritable.') }
$lines.Add('Interpretation: no OnConnection => investigate Outlook loading, registration and disabled add-ins.')
$lines.Add('OnConnection but no GetCustomUI => inspect supported window context and Outlook Ribbon UI errors.')
$lines.Add('GetCustomUI but no visible tab => enable Office UI error display and inspect rejected XML.')
$lines.Add('Visible tab but no callback entry => inspect IDispatch resolution; callback failure => HRESULT in log.')
$lines.Add('CI/diagnostic activation is not an Outlook integration test. No settings, secrets, message bodies or event messages collected.')
$parent = Split-Path $OutputPath -Parent
if ($parent) { New-Item -ItemType Directory -Force $parent | Out-Null }
$lines | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Host "Diagnosis written: $OutputPath"

if ($activationFailed) { $lines | Write-Output; throw 'COM activation failed; see the diagnosis report.' }
