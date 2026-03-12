# Disable-UsbConnectSounds.ps1
# Disables only the Windows device connect/disconnect sounds for the current user.

Set-ItemProperty -Path 'HKCU:\AppEvents\Schemes\Apps\.Default\DeviceConnect\.Current'    -Name '(Default)' -Value ''
Set-ItemProperty -Path 'HKCU:\AppEvents\Schemes\Apps\.Default\DeviceDisconnect\.Current' -Name '(Default)' -Value ''

Write-Host 'Disabled Device Connect and Device Disconnect sounds for the current user.'
