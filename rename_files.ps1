# Get current directory info
$CurrentDir = Get-Location
$DirName = Split-Path $CurrentDir -Leaf

# Get all files in current dir
$Files = Get-ChildItem -File | Sort-Object LastWriteTime

# Phase 1: Rename all files to a temporary unique name
foreach ($File in $Files) {
    try {
        $RandomID = [guid]::NewGuid().ToString("N")
        $Ext = $File.Extension
        $TempName = "temp_$RandomID$Ext"
        Rename-Item -LiteralPath $File.FullName -NewName $TempName -ErrorAction Stop
    } catch {
        Write-Warning "Phase 1 failed on '$($File.Name)': $($_.Exception.Message)"
    }
}

# Refresh file list (now renamed to temp)
$TempFiles = Get-ChildItem -File | Sort-Object LastWriteTime
$Counter = 1

# Phase 2: Rename temp files to final format
foreach ($File in $TempFiles) {
    try {
        $Ext = $File.Extension
        $NewName = "{0}_{1:D4}{2}" -f $DirName, $Counter, $Ext
        Rename-Item -LiteralPath $File.FullName -NewName $NewName -ErrorAction Stop
        $Counter++
    } catch {
        Write-Warning "Phase 2 failed on '$($File.Name)': $($_.Exception.Message)"
    }
}
