Add-Type -AssemblyName System.IO.Compression.FileSystem

﻿$sourceFolder = "C:\Path\To\Source\Folder"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$date = Get-Date -Format "yyyy-MM-dd"
$zipFileName = "Backup_$date-$timestamp.zip"
$zipFile = "C:\Path\To\Output\$zipFileName"

# Create a new zip file
[System.IO.Compression.ZipFile]::CreateFromDirectory($sourceFolder, $zipFile)

Write-Host "Folder '$sourceFolder' backed up to '$zipFile' successfully."
