# Import required modules
Import-Module PnP.PowerShell
Import-Module ImportExcel

# SharePoint site URL
$siteUrl = "www.sharepointsiteurl.com"

# Connect to SharePoint
try {
    Connect-PnPOnline -Url $siteUrl -Interactive
    Write-Host "Connected to SharePoint site: $siteUrl" -ForegroundColor Green
}
catch {
    Write-Host "Error connecting to SharePoint: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Function to convert bytes to MB
function ConvertTo-MB {
    param([long]$bytes)
    return [math]::Round($bytes / 1MB, 2)
}

$libraryName = "Shared Documents"

Write-Host "Scanning files in the '$libraryName' library..." -ForegroundColor Cyan

# Function to get all items in a folder using REST API
function Get-FolderItems {
    param ([string]$folderUrl)
    
    $restUrl = "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$folderUrl')/Files?`$expand=ListItemAllFields,Versions"
    $items = Invoke-PnPSPRestMethod -Url $restUrl -Method Get
    return $items.value
}

# Function to recursively process folders and files
function Process-Folder {
    param ([string]$folderUrl)
    
    Write-Host "Processing folder: $folderUrl" -ForegroundColor Yellow
    
    try {
        $items = Get-FolderItems -folderUrl $folderUrl

        foreach ($item in $items) {
            $fileName = $item.Name
            $fileType = $fileName.Split('.')[-1].ToLower()
            $versions = $item.Versions.Count + 1  # Current version + historical versions
            $currentSize = $item.Length
            $totalSize = $currentSize + ($item.Versions | Measure-Object -Property Size -Sum).Sum

            Write-Host "File: $fileName, Type: $fileType, Versions: $versions" -ForegroundColor Gray

            if ($versions -gt 1 -or $fileType -eq "psd") {
                $global:fileInfo += [PSCustomObject]@{
                    Name = $fileName
                    Type = $fileType
                    CurrentSizeMB = ConvertTo-MB -bytes $currentSize
                    TotalSizeMB = ConvertTo-MB -bytes $totalSize
                    Versions = $versions
                    URL = $item.ServerRelativeUrl
                    FolderPath = $folderUrl
                }
                Write-Host "Added to report: $fileName, Versions: $versions, Total Size: $(ConvertTo-MB -bytes $totalSize) MB" -ForegroundColor Green
            }
        }

        # Process subfolders
        $subFolders = Invoke-PnPSPRestMethod -Url "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$folderUrl')/Folders" -Method Get
        foreach ($subFolder in $subFolders.value) {
            Process-Folder -folderUrl $subFolder.ServerRelativeUrl
        }
    }
    catch {
        Write-Host "Error processing folder $folderUrl`: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Initialize file info array
$global:fileInfo = @()

# Start processing from the root folder
Process-Folder -folderUrl $libraryName

Write-Host "Total files processed: $($global:fileInfo.Count)" -ForegroundColor Green

# Sort files by total size (descending)
$sortedFiles = $global:fileInfo | Sort-Object TotalSizeMB -Descending

# Export to Excel
$excelPath = "C:\Users\Pablo.garner\Desktop\SharePointFilesWithVersionDetails.xlsx"
$sortedFiles | Export-Excel -Path $excelPath -AutoSize -TableName "FilesWithVersionDetails" -WorksheetName "Files" -FreezeTopRow

Write-Host "Results exported to: $excelPath" -ForegroundColor Cyan

# Disconnect from SharePoint
Disconnect-PnPOnline
Write-Host "Disconnected from SharePoint." -ForegroundColor Green
