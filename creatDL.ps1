# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline

# Path to the local text file containing DL email addresses
$dlFilePath = "\\\Users\FolderRedirect\\Desktop\DLUsers.txt"

# Read DL email addresses from the file
$dlAddresses = Get-Content $dlFilePath

# Define the members to be added to each DL
$membersToAdd = @(
    "member1@p4blo.dev",
    "member2@p4blo.dev"
)

# Define delay in seconds between creating each DL
$delaySeconds = 5

# Counter for created DLs
$createdDLs = 0

# Loop through each DL address and create the DL
foreach ($dlAddress in $dlAddresses) {
    try {
        # Create the DL using the email address from the file
        New-DistributionGroup -Name $dlAddress -PrimarySmtpAddress $dlAddress -MemberJoinRestriction Closed
        
        # Add the specified members to the DL
        foreach ($member in $membersToAdd) {
            Add-DistributionGroupMember -Identity $dlAddress -Member $member
        }
        
        Write-Host "Created DL: $dlAddress and added members"
        $createdDLs++
        
        # Delay before next operation
        Start-Sleep -Seconds $delaySeconds
    }
    catch {
        Write-Host "Error creating DL $dlAddress : $_"
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Script completed. $createdDLs DLs have been created and members added."