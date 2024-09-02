# P4BLO DEV Add User to Distribution Lists
# Version: 1.0
# Author: P4BLO DEV
# Description: This script adds a user to multiple distribution lists

function Show-BrandedBanner {
    Write-Host @"
██████╗ ██╗  ██╗██████╗ ██╗      ██████╗ 
██╔══██╗██║  ██║██╔══██╗██║     ██╔═══██╗
██████╔╝███████║██████╔╝██║     ██║   ██║
██╔═══╝ ╚════██║██╔══██╗██║     ██║   ██║
██║          ██║██████╔╝███████╗╚██████╔╝
╚═╝          ╚═╝╚═════╝ ╚══════╝ ╚═════╝ 
██████╗ ███████╗██╗   ██╗
██╔══██╗██╔════╝██║   ██║
██║  ██║█████╗  ██║   ██║
██║  ██║██╔══╝  ╚██╗ ██╔╝
██████╔╝███████╗ ╚████╔╝ 
╚═════╝ ╚══════╝  ╚═══╝  
"@
}

# Display the branded banner
Show-BrandedBanner

# Define the user to be added
$userToAdd = "email address of the user who needs to be added to dls"
# Define the path for the log file
$logFile = "C:\Powershell\Scripts\user+dl\AddUserToDistributionLists_Log.txt"
# Define the path for the file containing the list of Distribution lists
$distributionListFile = "C:\Powershell\Scripts\user+dl\DistributionLists.txt"
# Function to write to log file
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append
}
# Initialize counter for successful additions
$successCount = 0
# Check if the distribution list file exists
if (-not (Test-Path $distributionListFile)) {
    Write-Log "Error: Distribution list file not found at $distributionListFile"
    Write-Host "Distribution list file not found. Check the log file at $logFile for details."
    exit
}
# Read the list of Distribution lists from the file
$distributionLists = Get-Content $distributionListFile
Write-Log "Starting script execution"
Write-Log "User to be added: $userToAdd"
Write-Log "Total Distribution lists in the file: $($distributionLists.Count)"
# Connect to Exchange Online
try {
    Write-Log "Connecting to Exchange Online"
    Connect-ExchangeOnline -ErrorAction Stop
    Write-Log "Successfully connected to Exchange Online"
}
catch {
    Write-Log "Error connecting to Exchange Online: $_"
    Write-Host "Failed to connect to Exchange Online. Check the log file at $logFile for details."
    exit
}
foreach ($list in $distributionLists) {
    try {
        # Trim any whitespace from the list name
        $list = $list.Trim()
        
        # Add the user to the Distribution list
        Add-DistributionGroupMember -Identity $list -Member $userToAdd -ErrorAction Stop
        
        Write-Log "Successfully added $userToAdd to $list"
        $successCount++
    }
    catch {
        Write-Log "Error adding $userToAdd to $list. Error: $_"
    }
}
Write-Log "Script execution completed"
Write-Log "Successfully added to $successCount out of $($distributionLists.Count) Distribution lists"
# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Log "Disconnected from Exchange Online"
Write-Host "Script execution completed. Check the log file at $logFile for details."
