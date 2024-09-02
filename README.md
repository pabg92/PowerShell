```
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
```

# PowerShell Scripts Collection

This repository contains a collection of PowerShell scripts for various administrative tasks. Below is a detailed description of each script and its functionality.

## Table of Contents

1. [AddUserMemberToMultipleDLS.ps1](#addusermembertomultipledlsps1)
2. [get current user logged in.ps1](#get-current-user-logged-inps1)
3. [GetUserMemberships.ps1](#getusermembershipsps1)
4. [powershell-List multiple file versions on sharepoint.ps1](#powershell-list-multiple-file-versions-on-sharepointps1)

## AddUserMemberToMultipleDLS.ps1

### Description
This script adds a user to multiple distribution lists in Exchange Online.

### Features
- Displays a branded banner at the start of execution
- Connects to Exchange Online
- Reads distribution lists from a file
- Adds a specified user to each distribution list
- Logs all actions and errors
- Provides a summary of successful additions

### Usage
1. Ensure you have the necessary permissions to connect to Exchange Online and modify distribution lists.
2. Update the `$userToAdd` variable with the email address of the user to be added.
3. Prepare a text file (`DistributionLists.txt`) with the list of distribution lists, one per line.
4. Run the script in PowerShell.

### Requirements
- Exchange Online PowerShell module
- Appropriate permissions in Exchange Online

## get current user logged in.ps1

### Description
This script retrieves information about the currently logged-in user or the last logged-in user on a specified computer.

### Features
- Prompts for a computer name
- Checks for currently logged-in user
- If no current user, retrieves information about the last logged-in user
- Displays username, logon time, and status

### Usage
1. Run the script in PowerShell.
2. Enter the name of the computer you want to check when prompted.

### Requirements
- PowerShell remoting enabled on the target computer
- Appropriate permissions to query WMI on the target computer

## GetUserMemberships.ps1

### Description
This script retrieves all groups and distribution lists that a specified user is a member of in Active Directory.

### Features
- Imports the Active Directory module
- Retrieves all groups a user is a member of
- Separates security groups and distribution lists
- Displays the results in a formatted output

### Usage
1. Ensure you have the Active Directory module installed.
2. Run the script with the following command:
   ```powershell
   .\GetUserMemberships.ps1
   Get-UserMemberships -Username "username"
   ```
   Replace "username" with the actual username you want to check.

### Requirements
- Active Directory PowerShell module
- Appropriate permissions to query Active Directory

## powershell-List multiple file versions on sharepoint.ps1

### Description
This script scans a SharePoint document library for files with multiple versions or specific file types, and exports the results to an Excel file.

### Features
- Connects to a SharePoint site
- Recursively scans folders in a specified document library
- Identifies files with multiple versions or specific file types (e.g., .psd)
- Calculates current and total size of files (including all versions)
- Exports results to an Excel file, sorted by total size

### Usage
1. Update the `$siteUrl` variable with your SharePoint site URL.
2. Ensure you have the necessary SharePoint permissions.
3. Run the script in PowerShell.

### Requirements
- PnP.PowerShell module
- ImportExcel module
- Appropriate permissions to access the SharePoint site and document library

## General Notes

- Ensure you have the necessary permissions and modules installed before running these scripts.
- Review and update any hardcoded paths or URLs in the scripts to match your environment.

For any questions or issues, please contact github@p4blo.dev
