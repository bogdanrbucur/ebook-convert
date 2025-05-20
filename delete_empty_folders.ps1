# Set your root folder path here
$baseFolder = "C:\Users\bogda\iCloudDrive\Books\English"

# Get all directories, sorted deepest first
Get-ChildItem -Path $baseFolder -Recurse -Directory |
    Sort-Object FullName -Descending |
    ForEach-Object {
        $items = Get-ChildItem -Path $_.FullName -Force -ErrorAction SilentlyContinue

        if ($items.Count -eq 0) {
            # $response = Read-Host "Empty folder found: $($_.FullName) — Delete? (y/n)"
            # if ($response -eq 'y') {
                try {
                    Remove-Item $_.FullName -Force -Recurse
                    Write-Host "✅ Deleted: $($_.FullName)"
                } catch {
                    Write-Warning "❌ Failed to delete: $($_.FullName) — $($_.Exception.Message)"
                }
            # } else {
            #     Write-Host "❎ Skipped: $($_.FullName)"
            # }
        }
    }