# FlowVh_MouseClaim_ByOwnership.ps1
# PowerShell 5.1 compatible
#
# Watches VirtualHere ownership state first. If this PC does not currently own the
# keyboard, it watches for local mouse movement and attempts a takeover.
#
# Designed for a two-PC Logitech Flow setup where:
#   - LEFT PC releases to the RIGHT at the right-most edge
#   - RIGHT PC releases to the LEFT at the left-most edge
#
#
# Requires:
#   - VirtualHere Windows client installed and running
#   - VirtualHere service/background mode configured
#   - Device visible in GET CLIENT STATE XML

Add-Type -AssemblyName System.Windows.Forms

# =========================
# CONFIG
# =========================
$Side                = "Left"      # "Left" or "Right"
$DeviceAddress       = "REPLACE WITH YOUR DEVICE ADDRESS" #EX: pve.15 - Run Export-VirtualHereState.ps1 to learn this value
$DeviceProduct       = "REPLACE WITH YOUR DEVICE NAME" #EX: Logitech G710 Keyboard - Run Export-VirtualHereState.ps1 to learn this value

$PollIntervalMs      = 75
$StatusPollMs        = 250
$MovementThresholdPx = 1
$ReclaimCooldownMs   = 1200
$StopUseDelayMs      = 125
$Debug               = $false
# =========================

function Write-DebugLog {
    param([string]$Message)
    if ($Debug) {
        Write-Host "[$(Get-Date -Format HH:mm:ss.fff)] $Message"
    }
}

function Invoke-VhPipeCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command
    )

    $pipe = $null

    try {
        $pipe = New-Object System.IO.Pipes.NamedPipeClientStream("vhclient")
        $pipe.Connect()
        $pipe.ReadMode = [System.IO.Pipes.PipeTransmissionMode]::Message

        $writer = New-Object System.IO.StreamWriter($pipe)
        $writer.AutoFlush = $true
        $writer.Write($Command)

        $reader = New-Object System.IO.StreamReader($pipe)
        $bytes = New-Object System.Collections.Generic.List[byte]

        while ($reader.Peek() -ne -1) {
            $bytes.Add($reader.Read())
        }

        return [System.Text.Encoding]::UTF8.GetString($bytes).Trim()
    }
    finally {
        if ($pipe) {
            try { $pipe.Dispose() } catch {}
        }
    }
}

function Get-VhClientStateXml {
    $result = Invoke-VhPipeCommand -Command "GET CLIENT STATE"

    if ([string]::IsNullOrWhiteSpace($result)) {
        return $null
    }

    try {
        return [xml]$result
    }
    catch {
        Write-DebugLog "Failed to parse GET CLIENT STATE XML"
        return $null
    }
}

function Get-KeyboardDeviceNode {
    param([xml]$XmlDoc)

    if ($null -eq $XmlDoc) { return $null }
    if ($null -eq $XmlDoc.state) { return $null }
    if ($null -eq $XmlDoc.state.server) { return $null }

    foreach ($node in $XmlDoc.state.server.device) {
        if ($node.product -eq $DeviceProduct) {
            return $node
        }

        if (($node.address -eq "15") -and ($node.vendor -eq "Logitech")) {
            return $node
        }
    }

    return $null
}

function Test-KeyboardOwnedByThisPc {
    param([xml]$XmlDoc)

    $device = Get-KeyboardDeviceNode -XmlDoc $XmlDoc
    if ($null -eq $device) {
        Write-DebugLog "Keyboard device node not found"
        return $false
    }

    $boundHost = [string]$device.boundClientHostname
    $localHost = [string]$env:COMPUTERNAME

    Write-DebugLog "boundClientHostname=$boundHost localHost=$localHost state=$($device.state)"

    return ($boundHost -ieq $localHost)
}

function Get-CursorPosition {
    return [System.Windows.Forms.Cursor]::Position
}

function Is-LeaveEdge {
    param($Pos)

    if ($Side -eq "Right") {
        return ($Pos.X -le 0)
    }

    $screenWidth = [System.Windows.Forms.SystemInformation]::VirtualScreen.Width
    $rightEdge = $screenWidth - 1
    return ($Pos.X -ge $rightEdge)
}

function Attempt-Takeover {
    Write-DebugLog "Sending STOP USING,$DeviceAddress"
    $stopResult = Invoke-VhPipeCommand -Command "STOP USING,$DeviceAddress"
    Write-DebugLog "STOP result: $stopResult"

    Start-Sleep -Milliseconds $StopUseDelayMs

    Write-DebugLog "Sending USE,$DeviceAddress"
    $useResult = Invoke-VhPipeCommand -Command "USE,$DeviceAddress"
    Write-DebugLog "USE result: $useResult"

    return $useResult
}

$lastPos = Get-CursorPosition
$haveKeyboard = $false
$lastStatusCheck = [datetime]::MinValue
$lastReclaimAt = [datetime]::MinValue

while ($true) {
    $now = Get-Date

    if (($now - $lastStatusCheck).TotalMilliseconds -ge $StatusPollMs) {
        $xml = Get-VhClientStateXml
        $haveKeyboard = Test-KeyboardOwnedByThisPc -XmlDoc $xml
        Write-DebugLog "haveKeyboard=$haveKeyboard"
        $lastStatusCheck = $now
    }

    if (-not $haveKeyboard) {
        $pos = Get-CursorPosition
        $dx = [math]::Abs($pos.X - $lastPos.X)
        $dy = [math]::Abs($pos.Y - $lastPos.Y)
        $moved = ($dx -ge $MovementThresholdPx -or $dy -ge $MovementThresholdPx)

        if ($moved) {
            Write-DebugLog "Mouse moved. X=$($pos.X) Y=$($pos.Y) dx=$dx dy=$dy"

            if (-not (Is-LeaveEdge -Pos $pos)) {
                $cooldownOk = (($now - $lastReclaimAt).TotalMilliseconds -ge $ReclaimCooldownMs)

                if ($cooldownOk) {
                    $result = Attempt-Takeover
                    $lastReclaimAt = $now

                    if ($result -match '^OK') {
                        $haveKeyboard = $true
                    }
                }
            }
            else {
                Write-DebugLog "At leave edge, not taking over"
            }
        }

        $lastPos = $pos
    }
    else {
        $lastPos = Get-CursorPosition
    }

    Start-Sleep -Milliseconds $PollIntervalMs
}
