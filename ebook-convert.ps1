# Set your base folder here
$baseFolder = "C:\Users\bogda\iCloudDrive\Books\Romanian"

# Path to Calibre's ebook-convert executable
$ebookConvertPath = "C:\CalibrePortable\Calibre Portable\Calibre\ebook-convert.exe"

# Get all .docx files recursively
Get-ChildItem -Path $baseFolder -Recurse -Filter *.docx -File | ForEach-Object {
    $docxFile = $_.FullName
    $epubFile = [System.IO.Path]::ChangeExtension($docxFile, ".epub")
    $deleteFile = "$docxFile.DELETE"

    # Run conversion
    & "$ebookConvertPath" "$docxFile" "$epubFile"

    # If conversion succeeded, rename original
    if (Test-Path $epubFile) {
        Rename-Item -Path $docxFile -NewName "$($_.Name).DELETE"
    } else {
        Write-Warning "Conversion failed for: $docxFile"
    }
}
