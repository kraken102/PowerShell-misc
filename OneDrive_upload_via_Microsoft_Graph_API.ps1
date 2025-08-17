# ---------------------------------------
# OneDrive upload via Microsoft Graph (ID-based, resilient, HttpClient PUT)
# - Works for very large files (e.g., 15+ GB)
# - Chunks snap to 320 KiB grid (Graph requirement)
# - Retries + resume from server's nextExpectedRanges
# - Forces conflictBehavior=replace
# - Prints final webUrl and verifies by path
# ---------------------------------------

# ==== EDIT THESE ====
$LocalPath    = "C:\*\Backup.rar"       # local file
$OneDrivePath = "Backups\backup.rar"    # OneDrive target path (slashes normalized)

# ==== Safety / logging ====
$ErrorActionPreference = 'Stop'
$VerbosePreference     = 'Continue'

# ==== Hardening for HTTPS ====
[System.Net.ServicePointManager]::SecurityProtocol  = [System.Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::Expect100Continue = $false

# ---- Modules (ensure fresh Invoke-MgGraphRequest) ----
Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
try { Install-Module Microsoft.Graph.Authentication -Force -Scope CurrentUser -ErrorAction Stop } catch {}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# ---- HttpClient (for stable PUTs + Connection: close) ----
if (-not $script:HttpClient) {
  $script:HttpHandler = [System.Net.Http.HttpClientHandler]::new()
  $script:HttpClient  = [System.Net.Http.HttpClient]::new($script:HttpHandler)
  $script:HttpClient.Timeout = [TimeSpan]::FromMinutes(30)
  $script:HttpClient.DefaultRequestHeaders.ExpectContinue = $false
  $script:HttpClient.DefaultRequestHeaders.UserAgent.ParseAdd("PS-OneDriveUploader/1.3")
}

function Ensure-GraphConnection {
  param([string[]]$Scopes)
  try { $ctx = Get-MgContext } catch { $ctx=$null }
  if ($ctx -and $ctx.Scopes -and ($Scopes |? { $ctx.Scopes -contains $_ }).Count -eq $Scopes.Count) {
    Write-Verbose "Graph already connected with required scopes."
    return
  }
  Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
  try {
    Connect-MgGraph -Scopes $Scopes -NoWelcome
  } catch {
    Write-Warning "Browser sign-in failed. Using device code…"
    Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome
  }
  $ctx = Get-MgContext
  if (-not $ctx) { throw "Failed to connect to Microsoft Graph." }
  Write-Host "✅ Connected as: $($ctx.Account)" -ForegroundColor Green
}

function Normalize-OneDrivePath([string]$p){ ($p -replace '\\','/').TrimStart('/') }

function Ensure-OneDriveFolderPath {
  param([Parameter(Mandatory)][Object]$FolderPath)  # accept array/anything; coerce to string
  $FolderPath = [string](Normalize-OneDrivePath ($FolderPath -join '/'))
  if ([string]::IsNullOrWhiteSpace($FolderPath)) { return "root" }
  $parentId = "root"
  foreach($seg in ($FolderPath.Split('/') | Where-Object { $_ })) {
    $url = "https://graph.microsoft.com/v1.0/me/drive/items/$($parentId)/children?`$select=id,name,folder"
    $res = Invoke-MgGraphRequest -Method GET -Uri $url
    $found = $res.value | Where-Object { $_.name -eq $seg -and $_.folder }
    if ($found) { $parentId = $found.id; continue }
    Write-Verbose "Creating OneDrive folder '$seg' under $parentId"
    $body = @{ name=$seg; folder=@{}; "@microsoft.graph.conflictBehavior"="fail" } | ConvertTo-Json
    $created = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/me/drive/items/$($parentId)/children" -Body $body -ContentType "application/json"
    if (-not $created.id){ throw "Failed to create folder '$seg'." }
    $parentId = $created.id
  }
  return $parentId
}

# Snap a size to the nearest LOWER multiple of 320 KiB (327,680 bytes)
function Snap-ToGraphGrid([Int64]$size){
  $grid = 327680
  return [int64]([math]::Floor($size / $grid) * $grid)
}

function Get-SessionResumeStart {
  param([Parameter(Mandatory)][string]$UploadUrl)
  try {
    $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $UploadUrl)
    $req.Headers.ConnectionClose = $true
    $resp = $script:HttpClient.SendAsync($req).GetAwaiter().GetResult()
    $content = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    $resp.Dispose()
    if ($content) {
      $j = $content | ConvertFrom-Json
      if ($j.nextExpectedRanges -and $j.nextExpectedRanges.Count -gt 0) {
        $first = $j.nextExpectedRanges[0]  # e.g. "7083733760-"
        $startStr = ($first -split '-')[0]
        if ($startStr -match '^\d+$') { return [int64]$startStr }
      }
    }
  } catch {
    Write-Verbose "Could not query session state: $($_.Exception.Message)"
  }
  return $null
}

function Upload-LargeFile {
  param(
    [Parameter(Mandatory)][string]$LocalFile,
    [Parameter(Mandatory)][string]$SessionUrl
  )

  Write-Verbose "Creating upload session: $SessionUrl"
  # Force predictable conflict behavior: replace target if it exists
  $sessionBody = @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } } | ConvertTo-Json
  $session = Invoke-MgGraphRequest -Method POST -Uri $SessionUrl -Body $sessionBody -ContentType "application/json"
  $uploadUrl = $session.uploadUrl
  if (-not $uploadUrl) { throw "Failed to create upload session." }

  $fs = [System.IO.File]::OpenRead($LocalFile)
  try {
    # Base chunk = 20 MB (64 * 320 KiB). Min chunk = 8 * 320 KiB.
    [Int64]$chunkSize = Snap-ToGraphGrid 20MB
    [Int64]$minChunk  = 2560KB
    [Int64]$total     = $fs.Length
    [Int64]$start     = 0

    while ($start -lt $total) {
      $fs.Position = $start
      [Int64]$remaining = $total - $start

      # For all but last chunk, keep size on the 320 KiB grid
      [Int64]$thisSize  = if ($remaining -lt $chunkSize) { $remaining } else { $chunkSize }
      if ($remaining -gt $chunkSize) { $thisSize = Snap-ToGraphGrid $thisSize }

      # byte[] must be Int32 length
      $buffer = New-Object byte[] ([int]$thisSize)
      $read   = $fs.Read($buffer, 0, $buffer.Length)
      if ($read -le 0) { break }

      [Int64]$end = $start + $read - 1

      # Progress
      $percent = [int](($end + 1) * 100 / $total)
      Write-Progress -Activity "Uploading to OneDrive" -Status "$percent% ($start-$end)" -PercentComplete $percent
      Write-Verbose ("PUT chunk {0:n0}-{1:n0} of {2:n0} (size {3:n0})" -f $start, $end, $total, $read)

      $maxRetries = 6
      $attempt = 0
      while ($true) {
        try {
          $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $uploadUrl)
          $req.Headers.ConnectionClose = $true

          $content = [System.Net.Http.ByteArrayContent]::new($buffer, 0, $read)
          $null = $content.Headers.TryAddWithoutValidation("Content-Range", "bytes $start-$end/$total")
          $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")
          $req.Content = $content

          $resp   = $script:HttpClient.SendAsync($req).GetAwaiter().GetResult()
          $status = [int]$resp.StatusCode
          $body   = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
          $resp.Dispose()

          if ($status -in 200,201) {
            Write-Verbose "Upload complete."
            if ($body) {
              try {
                $item = $body | ConvertFrom-Json
                if ($item.webUrl) {
                  Write-Host "✅ Final item: $($item.name)" -ForegroundColor Green
                  Write-Host "🔗 $($item.webUrl)" -ForegroundColor Cyan
                }
              } catch {}
            }
            $start = $end + 1
            break
          }
          elseif ($status -eq 202) {
            # In-progress response; honor server's nextExpectedRanges if present
            if ($body) {
              try {
                $j = $body | ConvertFrom-Json
                if ($j.nextExpectedRanges -and $j.nextExpectedRanges.Count -gt 0) {
                  $nextStart = [int64](($j.nextExpectedRanges[0] -split '-')[0])
                  if ($nextStart -gt $start) {
                    Write-Verbose "Server suggests next start at $nextStart; resyncing."
                    $start = $nextStart
                    break # break retry loop to rebuild buffer
                  }
                }
              } catch {}
            }
            # If no hint, proceed sequentially
            $start = $end + 1
            break
          }
          else {
            throw "Unexpected status $status`n$body"
          }
        }
        catch [System.IO.IOException],[System.Net.WebException],[System.Net.Http.HttpRequestException],[System.Threading.Tasks.TaskCanceledException],[System.Exception] {
          $attempt++
          Write-Warning ("Chunk {0}-{1} failed (attempt {2}/{3}): {4}" -f $start,$end,$attempt,$maxRetries,$_.Exception.Message)
          if ($attempt -ge $maxRetries) { throw }

          # Ask the session where to resume
          $resumeStart = Get-SessionResumeStart -UploadUrl $uploadUrl
          if ($resumeStart -ne $null -and $resumeStart -ge 0) {
            Write-Verbose "Server nextExpectedRanges start = $resumeStart; resyncing."
            $start = $resumeStart
            break  # break retry loop to rebuild buffer for new $start
          }

          # After a few misses, shrink chunk size for stability (down to minChunk) and snap to grid
          if ($attempt -ge 3 -and $chunkSize -gt $minChunk) {
            $chunkSize = [Int64]([Math]::Max($minChunk, [int64]($chunkSize / 2)))
            $chunkSize = Snap-ToGraphGrid $chunkSize
            Write-Verbose "Reducing chunk size to $chunkSize bytes (320 KiB aligned)."
          }

          # Exponential backoff (1,2,4,8,16,32s; cap at 60)
          $delay = [int][Math]::Min(60, [Math]::Pow(2, $attempt))
          Start-Sleep -Seconds $delay
          # loop to retry same (or resumed) chunk
        }
      } # retry loop
    } # while
  }
  finally {
    $fs.Dispose()
  }
}

function Upload-SmallFile {
  param(
    [Parameter(Mandatory)][string]$LocalFile,
    [Parameter(Mandatory)][string]$ContentUrl
  )
  Write-Verbose "Uploading small file to $ContentUrl"
  $bytes = [System.IO.File]::ReadAllBytes($LocalFile)

  $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $ContentUrl)
  $req.Headers.ConnectionClose = $true
  $content = [System.Net.Http.ByteArrayContent]::new($bytes)
  $content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")
  $req.Content = $content
  $resp = $script:HttpClient.SendAsync($req).GetAwaiter().GetResult()
  $status = [int]$resp.StatusCode
  $body   = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
  $resp.Dispose()
  if ($status -notin 200,201) { throw "Small upload failed with status $status`n$body" }
}

# ===== Main =====
if (-not (Test-Path -LiteralPath $LocalPath)) { throw "Local file not found: $LocalPath" }
Ensure-GraphConnection -Scopes @("Files.ReadWrite")   # least-priv for personal/work OneDrive

# Normalize and split path
$normPath   = Normalize-OneDrivePath $OneDrivePath
$lastSlash  = $normPath.LastIndexOf('/')
if ($lastSlash -ge 0) {
  $parentPath = $normPath.Substring(0, $lastSlash)
  $fileName   = $normPath.Substring($lastSlash + 1)
} else {
  $parentPath = ''
  $fileName   = $normPath
}

if ([string]::IsNullOrWhiteSpace($fileName)) { throw "Target file name resolved empty." }

# Ensure folders exist and get parent item ID
$parentId    = Ensure-OneDriveFolderPath -FolderPath $parentPath
$encodedName = [System.Uri]::EscapeDataString($fileName)

# Build ID-based endpoints
$sessionUrl  = "https://graph.microsoft.com/v1.0/me/drive/items/$($parentId):/$($encodedName):/createUploadSession"
$contentUrl  = "https://graph.microsoft.com/v1.0/me/drive/items/$($parentId):/$($encodedName):/content"

# Upload
$length = (Get-Item -LiteralPath $LocalPath).Length
$ok = $false
try {
  if ($length -lt 4MB) {
    Write-Host "Uploading small file via simple PUT..." -ForegroundColor Cyan
    Upload-SmallFile -LocalFile $LocalPath -ContentUrl $contentUrl
  } else {
    Write-Host "Uploading large file via resumable session..." -ForegroundColor Cyan
    Upload-LargeFile -LocalFile $LocalPath -SessionUrl $sessionUrl
  }
  $ok = $true
}
catch {
  Write-Host "❌ Upload failed." -ForegroundColor Red
  Write-Warning $_
  if ($_.Exception.Response -and $_.Exception.Response.Content) {
    Write-Host "Server said:" -ForegroundColor Yellow
    $_.Exception.Response.Content | Out-String | Write-Host
  }
  throw
}
finally {
  if ($ok) {
    # Verify by path and print a clickable URL
    $verifyUrl = "https://graph.microsoft.com/v1.0/me/drive/items/$($parentId):/$($encodedName)"
    try {
      $finalItem = Invoke-MgGraphRequest -Method GET -Uri $verifyUrl
      if ($finalItem.webUrl) {
        Write-Host "✅ Upload complete & verified: $OneDrivePath" -ForegroundColor Green
        Write-Host "🔗 Open: $($finalItem.webUrl)" -ForegroundColor Cyan
      } else {
        Write-Warning "Upload reported success but item wasn't found at expected path (may have been renamed)."
      }
    } catch {
      Write-Warning "Verification GET failed: $($_.Exception.Message)"
    }
  }
  Disconnect-MgGraph | Out-Null
  if ($script:HttpClient) { $script:HttpClient.Dispose() ; $script:HttpClient = $null }
}
