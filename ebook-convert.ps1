# Set your base folder here
$baseFolder = "C:\Users\bogda\iCloudDrive\Books\English"

# Path to Calibre's ebook-convert executable
$ebookConvertPath = "C:\CalibrePortable\Calibre Portable\Calibre\ebook-convert.exe"

# Get all .mobi files recursively
Get-ChildItem -Path $baseFolder -Recurse -Filter *.mobi -File | ForEach-Object {
    $mobiFile = $_.FullName
    $epubFile = [System.IO.Path]::ChangeExtension($mobiFile, ".epub")
    $deleteFile = "$mobiFile.DELETE"

    # Run conversion
    & "$ebookConvertPath" "$mobiFile" "$epubFile"

    # If conversion succeeded, rename original
    if (Test-Path $epubFile) {
        Rename-Item -Path $mobiFile -NewName "$($_.Name).DELETE"
    } else {
        Write-Warning "Conversion failed for: $mobiFile"
    }
}
