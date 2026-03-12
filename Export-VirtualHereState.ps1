# Export-VirtualHereState.ps1
# Dumps GET CLIENT STATE XML to %TEMP%\vh_state.xml and opens it in Notepad.

$pipe = New-Object System.IO.Pipes.NamedPipeClientStream('vhclient')
$pipe.Connect()
$pipe.ReadMode = [System.IO.Pipes.PipeTransmissionMode]::Message

$writer = New-Object System.IO.StreamWriter($pipe)
$writer.AutoFlush = $true
$writer.Write('GET CLIENT STATE')

$reader = New-Object System.IO.StreamReader($pipe)
$bytes = New-Object System.Collections.Generic.List[byte]

while ($reader.Peek() -ne -1) {
    $bytes.Add($reader.Read())
}

$xml = [System.Text.Encoding]::UTF8.GetString($bytes)
$path = Join-Path $env:TEMP 'vh_state.xml'
$xml | Out-File $path -Encoding utf8
$pipe.Dispose()

notepad $path
