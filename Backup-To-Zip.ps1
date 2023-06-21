$sourceFolder = "C:\Path\To\Source\Folder"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$date = Get-Date -Format "yyyy-MM-dd"
$zipFileName = "Backup_$date-$timestamp.zip"
$zipFile = "C:\Path\To\Output\$zipFileName"

# Create a new empty zip file
$shell = New-Object -ComObject Shell.Application
$zipPackage = $shell.NameSpace($zipFile)
$null = $zipPackage.CopyHere($null)

# Wait for the zip file to be created
Start-Sleep -Seconds 2

# Get the full list of files and subdirectories in the source folder
$files = Get-ChildItem -Path $sourceFolder -Recurse

# Add each file and subdirectory to the zip file
foreach ($file in $files) {
    $relativePath = $file.FullName.Substring($sourceFolder.Length + 1)
    $zipPackage.CopyHere($file.FullName, 16)
    # Wait for each item to be added before proceeding
    Start-Sleep -Milliseconds 200
}

Write-Host "Folder '$sourceFolder' backed up to '$zipFile' successfully."