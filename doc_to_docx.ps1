$baseFolder = "C:\Users\bogda\iCloudDrive\Books\Romanian"

# Start Word once
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

# Get all .doc files (excluding .docx)
$docFiles = Get-ChildItem -Path $baseFolder -Recurse -Include *.doc -File |
    Where-Object { $_.Extension -ieq ".doc" -and -not $_.Name.EndsWith(".docx") }

foreach ($file in $docFiles) {
    $docPath = $file.FullName
    $docxPath = Join-Path $file.DirectoryName "$($file.BaseName).docx"

    try {
        Write-Host "Converting: $docPath → $docxPath"

        # Open the document (readOnly = false, confirmConversions = false, addToRecentFiles = false)
        $document = $word.Documents.Open($docPath, $false, $false, $false)

        # Save as DOCX (format 16 = wdFormatDocumentDefault)
        $document.SaveAs($docxPath, 16)

        # Close document
        $document.Close($false)

        # Rename original to .DELETE only if .docx exists
        if (Test-Path $docxPath) {
            Rename-Item -Path $docPath -NewName "$($file.Name).DELETE"
            Write-Host "✅ Converted and renamed: $($file.Name)"
        } else {
            Write-Warning "❌ Conversion failed or skipped: $docPath"
        }

    } catch {
        Write-Warning "❌ Error during conversion: $docPath - $_"
    }
}

# Quit Word
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()