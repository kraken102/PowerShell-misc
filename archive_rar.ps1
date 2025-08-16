# ----------------------------
# User-configurable variables
# ----------------------------
$RarExe        = "C:\Program Files\WinRAR\rar.exe"   # Path to rar.exe
$ArchiveName   = "backup.rar"                        # Name of the archive
$SourceFiles   = "$PWD\*"                            # Files/folders in current directory
$Password      = "SuperSecret123"                    # Password
$Recurse       = $true                               # recurse subfolders
$Overwrite     = $true                               # overwrite existing archive
$EncryptNames  = $true                               # true = -hp, false = -p
# ----------------------------

# Build the output path in the current directory
$OutputArchive = Join-Path -Path $PWD -ChildPath $ArchiveName

# Validate rar.exe
if (-not (Test-Path -LiteralPath $RarExe)) {
    throw "rar.exe not found at '$RarExe'. Update `$RarExe."
}

# Handle overwrite
if (Test-Path -LiteralPath $OutputArchive) {
    if ($Overwrite) {
        Remove-Item -LiteralPath $OutputArchive -Force
    }
    else {
        throw "Archive already exists: $OutputArchive"
    }
}

# Pick password switch
$passSwitch = if ($EncryptNames) { "-hp$Password" } else { "-p$Password" }

# Run rar.exe directly (no need to pre-build args)
if ($Recurse) {
    & $RarExe a $passSwitch -ep1 -y -r $OutputArchive $SourceFiles
} else {
    & $RarExe a $passSwitch -ep1 -y $OutputArchive $SourceFiles
}

if ($LASTEXITCODE -eq 0) {
    Write-Host "Archive created: $OutputArchive" -ForegroundColor Green
} else {
    Write-Host "rar.exe exited with code $LASTEXITCODE" -ForegroundColor Red
}
