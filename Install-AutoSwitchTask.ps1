# Install-AutoSwitchTask.ps1
# Run as Administrator on each PC.
# Creates a hidden logon task that launches the watcher via WScript.

param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Left', 'Right')]
    [string]$Side,

    [string]$TaskName = 'KeyboardFollowVirtualHere',
    [string]$BaseFolder = 'C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow'
)

$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $BaseFolder 'FlowVh_MouseClaim_ByOwnership.ps1'
$vbsPath    = Join-Path $BaseFolder 'RunHidden.vbs'

if (-not (Test-Path $scriptPath)) {
    throw "Watcher script not found: $scriptPath"
}

New-Item -ItemType Directory -Path $BaseFolder -Force | Out-Null

$vbs = @"
Set WshShell = CreateObject(""WScript.Shell"")
WshShell.Run ""powershell.exe -NoProfile -ExecutionPolicy Bypass -File """"$scriptPath"""" -Side $Side"", 0, False
"@
Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII

$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$vbsPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances Ignore -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Auto-switch VirtualHere USB keyboard ownership based on Logitech Flow mouse movement." `
    -Force | Out-Null

Start-ScheduledTask -TaskName $TaskName
Write-Host "Installed task '$TaskName' for side '$Side'."
