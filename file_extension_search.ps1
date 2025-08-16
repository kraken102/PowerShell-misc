# Set the initial path to the current working directory
$initialPath = (Get-Location).ProviderPath

Write-Host "📁 Scanning $initialPath for file extensions..." -ForegroundColor Cyan

# Step 1: Find all unique file extensions in the current directory and subdirectories
$extensions = Get-ChildItem -Path $initialPath -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -ne "" } |
    Select-Object -ExpandProperty Extension -Unique

if ($extensions.Count -eq 0) {
    Write-Host "❌ No file extensions found in the current directory. Exiting." -ForegroundColor Red
    exit
}

Write-Host "✅ Found $($extensions.Count) unique extensions:`n$($extensions -join ", ")" -ForegroundColor Green

# Step 2: Search entire C:\ drive for files with those extensions,
# but exclude files from the initial path
Write-Host "🔍 Searching entire C:\ drive for matching files (excluding $initialPath)..." -ForegroundColor Yellow

$results = @()

foreach ($ext in $extensions) {
    $pattern = "*$ext"
    Write-Host "🔎 Searching for files matching $pattern..." -ForegroundColor Magenta

    try {
        $found = Get-ChildItem -Path "C:\" -Recurse -Include $pattern -File -ErrorAction SilentlyContinue |
            Where-Object { -not ($_.FullName -like "$initialPath*") }

        $results += $found
    } catch {
        Write-Host "⚠️ Error while searching for $($pattern): $($_)" -ForegroundColor Red
    }
}

# Output summary
Write-Host "`n✅ Search complete. Found $($results.Count) matching files (excluding current dir)." -ForegroundColor Green

# Export results to CSV
$csvPath = Join-Path -Path $initialPath -ChildPath "found_files.csv"
$results | Select-Object FullName | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "💾 Results exported to $csvPath" -ForegroundColor Cyan

# Display first few results
if ($results.Count -gt 0) {
    Write-Host "`n📝 Showing first 10 results:" -ForegroundColor Cyan
    $results | Select-Object FullName -First 10
} else {
    Write-Host "💤 No matching files found on C:\ (excluding $initialPath)." -ForegroundColor DarkGray
}
