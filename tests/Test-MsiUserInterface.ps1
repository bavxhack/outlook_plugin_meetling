# Runs in the disposable standard-user profile. Exercise a real double-click UI flow.
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string] $MsiPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Wait-SetupWindow([string] $Title) {
    $deadline = (Get-Date).AddSeconds(90)
    $condition = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $Title)
    while ((Get-Date) -lt $deadline) {
        $window = [Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children, $condition)
        if ($null -ne $window) { return $window }
        if ($script:setup.HasExited) { throw "Setup exited before showing '$Title': $($script:setup.ExitCode)" }
        Start-Sleep -Milliseconds 250
    }
    throw "Setup window did not appear: $Title"
}

function Invoke-SetupButton($Window, [string[]] $Names) {
    foreach ($name in $Names) {
        $condition = [Windows.Automation.AndCondition]::new(
            [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $name),
            [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Button)
        )
        $button = $Window.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
        if ($null -ne $button -and $button.Current.IsEnabled) {
            $pattern = $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
            $pattern.Invoke()
            return
        }
    }
    throw "Setup button missing: $($Names -join ', ')"
}

$setup = Start-Process "$env:WINDIR\System32\msiexec.exe" -ArgumentList @('/i', "`"$MsiPath`"", '/norestart') -PassThru
$finished = $false
try {
    $welcome = Wait-SetupWindow 'Meetling für Outlook – Installation'
    Invoke-SetupButton $welcome @('Installieren')
    # Standard WiX dialogs use ProductName as their window title. Wait for Finish;
    # progress and error windows can have the same title and must not count as success.
    $deadline = (Get-Date).AddSeconds(90)
    while ((Get-Date) -lt $deadline) {
        $window = Wait-SetupWindow 'Meetling Outlook Add-in-Setup'
        try {
            Invoke-SetupButton $window @('Fertig stellen', '&Fertig stellen')
            $finished = $true
            break
        } catch {
            if ($setup.HasExited) { throw }
            Start-Sleep -Milliseconds 250
        }
    }
    if (-not $finished) { throw 'Setup did not show a completion dialog.' }
    if (-not $setup.WaitForExit(30000) -or $setup.ExitCode -notin @(0, 3010)) { throw 'Interactive MSI installation failed.' }
    Write-Output 'PASS: visible MSI welcome, Install button and completion dialog under standard user.'
} finally {
    if (-not $setup.HasExited) { Stop-Process -Id $setup.Id -Force -ErrorAction SilentlyContinue }
    $setup.Dispose()
}
