# P4BLO DEV Get User Memberships
# Version: 1.0
# Author: P4BLO DEV
# Description: This script retrieves all groups and distribution lists that a specified user is a member of in Active Directory.

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

# Import the Active Directory module
Import-Module ActiveDirectory

# Function to get all groups and distribution lists a user is a member of
function Get-UserMemberships {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserName
    )

    try {
        # Get the user object
        $user = Get-ADUser -Identity $UserName -Properties MemberOf -ErrorAction Stop

        # Get all groups the user is a member of
        $groups = $user.MemberOf | ForEach-Object {
            Get-ADGroup -Identity $_ -Properties GroupCategory
        }

        # Separate security groups and distribution lists
        $securityGroups = $groups | Where-Object {$_.GroupCategory -eq 'Security'}
        $distributionLists = $groups | Where-Object {$_.GroupCategory -eq 'Distribution'}

        # Output results
        Write-Host "User $UserName is a member of the following groups and distribution lists:"
        
        Write-Host "`nSecurity Groups:"
        $securityGroups | ForEach-Object { Write-Host "- $($_.Name)" }

        Write-Host "`nDistribution Lists:"
        $distributionLists | ForEach-Object { Write-Host "- $($_.Name)" }

        # Return the groups (can be used for further processing if needed)
        return $groups
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "Error: User not found. Please check the provided username."
    }
    catch {
        Write-Host "An error occurred: $_"
    }
}

# Example usage:
# Get-UserMemberships -Username "p4blo.dev"
