# Auto Switch USB Devices with Logitech Flow

Automatically moves a **USB device shared through VirtualHere** to the currently active Windows PC by watching **Logitech Flow mouse movement**.

This was built for a two-PC setup where:

- the **left PC** hands off when the cursor reaches its **right edge**
- the **right PC** hands off when the cursor reaches its **left edge**
- the device is plugged into a **VirtualHere USB server**
- both PCs run the **VirtualHere Windows client**

In this repo, the target device example is:

- **Product:** `Logitech G710 Keyboard`
- **VirtualHere address:** `pve.15`

Because apparently buying a Logitech keyboard was too obvious.

---

## What this does

The watcher script does **not** constantly hammer VirtualHere with process launches.

Instead it:

1. Polls `GET CLIENT STATE` from the running VirtualHere client through the documented named pipe API.
2. Checks whether this PC already owns the keyboard by comparing:
   - `boundClientHostname`
   - local `$env:COMPUTERNAME`
3. Only if this PC does **not** own the keyboard, it watches for **local mouse movement**.
4. On movement, if the cursor is not sitting on the handoff edge, it does:
   - `STOP USING,pve.15`
   - short delay
   - `USE,pve.15`

That makes the active PC reclaim the keyboard without you manually doing anything.

---

## Files in this repo

- `FlowVh_MouseClaim_ByOwnership.ps1` - main watcher script
- `Install-AutoSwitchTask.ps1` - creates a hidden logon scheduled task
- `Uninstall-AutoSwitchTask.ps1` - removes the scheduled task
- `Disable-UsbConnectSounds.ps1` - disables Windows USB connect/disconnect sounds for the current user
- `Export-VirtualHereState.ps1` - dumps raw VirtualHere XML state for troubleshooting

---

## Requirements

### Hardware / topology

- 2 Windows PCs
- Logitech Flow working for the mouse
- keyboard connected to a VirtualHere server
- reachable VirtualHere server hosting the keyboard

### Software

- VirtualHere Windows client installed on both PCs
- VirtualHere server properly licensed/configured
- PowerShell 5.1 or newer
- Task Scheduler enabled

---

## VirtualHere assumptions

The script is currently built around this exact device identity:

- `DeviceAddress = "pve.15"`
- `DeviceProduct = "Logitech G710 Keyboard"`

If either changes, update the config block at the top of `FlowVh_MouseClaim_ByOwnership.ps1`.

---

## How ownership detection works

`GET CLIENT STATE` returns XML like this for the keyboard:

```xml
<device vendor="Logitech"
        product="Logitech G710 Keyboard"
        address="15"
        ...
        boundClientHostname="KGRANT11A" />
```

The script treats the keyboard as **owned by this PC** when:

- `boundClientHostname` matches `$env:COMPUTERNAME`

That is the clean signal used to decide whether the script should even care about mouse movement.

---

## Setup

### 1. Copy files to both PCs

Recommended path:

```text
C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow
```

Put all repo files there on both PCs.

---

### 2. Edit the side on each PC

Open `FlowVh_MouseClaim_ByOwnership.ps1` and set:

On the **left PC**:

```powershell
$Side = "Left"
```

On the **right PC**:

```powershell
$Side = "Right"
```

That is the only per-PC config required in the script as packaged.

---

### 3. Make sure VirtualHere client is already running

This script talks to the VirtualHere **named pipe API**:

```text
\\.\pipe\vhclient
```

So the VirtualHere client must already be installed and running on each PC.

---

### 4. Test manually first

Run the script in PowerShell before installing it as a hidden task:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow\FlowVh_MouseClaim_ByOwnership.ps1"
```

Verify that:

- keyboard follows active PC
- no weird fighting occurs
- edge handoff feels correct

If you want console logging during testing, temporarily set:

```powershell
$Debug = $true
```

Then set it back to `false` when done.

---

### 5. Install hidden startup task

Run **PowerShell as Administrator** and install the task.

On the **left PC**:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow\Install-AutoSwitchTask.ps1" -Side Left
```

On the **right PC**:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow\Install-AutoSwitchTask.ps1" -Side Right
```

This creates a Scheduled Task named:

```text
KeyboardFollowVirtualHere
```

It launches the watcher hidden at logon using a small `RunHidden.vbs` wrapper so you do not get a stupid visible PowerShell window every time you sign in.

---

### 6. Disable USB connect/disconnect sounds

If Windows keeps making the USB ding/dong noise every time ownership changes, run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow\Disable-UsbConnectSounds.ps1"
```

This disables only:

- `Device Connect`
- `Device Disconnect`

for the current user.

---

## Uninstall

Remove the startup task:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow\Uninstall-AutoSwitchTask.ps1"
```

---

## Troubleshooting

### Export raw VirtualHere state XML

If device detection or ownership logic stops working, dump the raw state:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\AutoSwitchUsbDevicesWithLogitechFlow\Export-VirtualHereState.ps1"
```

This writes the XML to:

```text
%TEMP%\vh_state.xml
```

and opens it in Notepad.

---

### Common issues

#### The script runs but nothing happens

Check that VirtualHere client is running and exposing `\\.\pipe\vhclient`.

#### The wrong PC claims the keyboard

Make sure `$Side` is correct on each machine.

#### Keyboard identity changed

If VirtualHere address or product name changed, update:

```powershell
$DeviceAddress
$DeviceProduct
```

#### PowerShell window still appears at logon

Re-run `Install-AutoSwitchTask.ps1`. The task should launch `wscript.exe` against `RunHidden.vbs`, not `powershell.exe` directly.

---

## Tuning

Useful knobs in `FlowVh_MouseClaim_ByOwnership.ps1`:

```powershell
$PollIntervalMs
$StatusPollMs
$MovementThresholdPx
$ReclaimCooldownMs
$StopUseDelayMs
```

Good starting behavior is already baked in.

If it feels too eager, increase:

```powershell
$ReclaimCooldownMs
$StopUseDelayMs
```

If it feels sluggish, reduce them slightly.

---

## Notes

This solution is intentionally built around:

- Windows
- Logitech Flow for mouse handoff
- VirtualHere named pipe API
- a two-PC left/right topology

It is not trying to be a universal framework for every USB device under the sun because that way lies madness.
