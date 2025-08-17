param(
  [string]$Source   = "C:\POC\Source",
  [string]$UsbLabel = "DUCKY",
  [int]   $MaxMB    = 200
)

$ErrorActionPreference = 'Stop'

function Get-UsbByLabel([string]$Label) {
  Get-Volume -FileSystemLabel $Label -ErrorAction SilentlyContinue |
    Where-Object { $_.DriveType -eq 'Removable' } |
    Select-Object -First 1
}

function Write-Log([string]$Msg) {
  try {
    if ($script:LogPath) {
      ("{0:o} {1}" -f (Get-Date).ToUniversalTime(), $Msg) | Add-Content -Path $script:LogPath -Encoding UTF8
    }
  } catch { }
}

try {
  # Locate USB first
  $usb = Get-UsbByLabel $UsbLabel
  if (-not $usb) { throw "USB with label '$UsbLabel' not found." }
  $usbRoot = ($usb.DriveLetter + ":\")
  $script:LogPath = Join-Path $usbRoot "poc_copy.log"
  Write-Log "[*] Starting (Source=$Source, UsbLabel=$UsbLabel, MaxMB=$MaxMB)"

  # Validate source
  if (-not (Test-Path -LiteralPath $Source)) { throw "Source path not found: $Source" }

  # Guardrail: require at least one POC_OK.txt
  $marker = Get-ChildItem -LiteralPath $Source -Recurse -Filter "POC_OK.txt" -ErrorAction SilentlyContinue
  if (-not $marker) { throw "No POC_OK.txt marker found under $Source." }

  # Size check
  $totalBytes = (Get-ChildItem -LiteralPath $Source -Recurse -File | Measure-Object Length -Sum).Sum
  $sizeMB = [math]::Round($totalBytes / 1MB, 2)
  if ($sizeMB -gt $MaxMB) { throw "Source size ${sizeMB}MB exceeds limit ${MaxMB}MB." }

  # Destination
  $destRoot = Join-Path $usbRoot "POC_Extracted"
  if (-not (Test-Path -LiteralPath $destRoot)) { New-Item -ItemType Directory -Path $destRoot | Out-Null }
  $dest = $destRoot

  Write-Log "[*] Copying ${sizeMB}MB from '$Source' to '$dest'"
  Copy-Item -Path (Join-Path $Source '*') -Destination $dest -Recurse -Force -ErrorAction Stop

  # Manifest
  $files = Get-ChildItem -LiteralPath $dest -Recurse -File | ForEach-Object {
    [pscustomobject]@{
      path       = $_.FullName.Substring($usbRoot.Length)
      bytes      = $_.Length
      mtime_utc  = $_.LastWriteTimeUtc
    }
  }
  $manifest = [pscustomobject]@{
    source       = $Source
    usbLabel     = $UsbLabel
    generatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    fileCount    = $files.Count
    totalBytes   = ($files | Measure-Object bytes -Sum).Sum
    markersFound = ($marker | Select-Object -ExpandProperty FullName)
    files        = $files
  }
  $manifestPath = Join-Path $dest "manifest.json"
  $manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding UTF8

  Write-Log "[+] Copy complete. Files=$($files.Count) Bytes=$($manifest.totalBytes) Manifest=$(Resolve-Path $manifestPath)"
}
catch {
  Write-Log "[!] ERROR: $($_.Exception.Message)"
  exit 1
}
