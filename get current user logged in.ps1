# Function to get the current or last logged-in user
function Get-LoggedInUser {
    param (
        [string]$ComputerName
    )
    
    try {
        $output = @{
            ComputerName = $ComputerName
            Username = $null
            LogonTime = $null
            Status = $null
        }

        # Check for currently logged-in user
        $explorer = Get-WmiObject -Class Win32_Process -ComputerName $ComputerName -Filter "Name = 'explorer.exe'" -ErrorAction Stop
        
        if ($explorer) {
            $output.Username = ($explorer.GetOwner()).User
            $output.LogonTime = (Get-WmiObject -Class Win32_LogonSession -ComputerName $ComputerName | Where-Object {$_.LogonType -eq 2} | Sort-Object StartTime -Descending | Select-Object -First 1).StartTime
            $output.Status = "Currently logged in"
        } else {
            # If no current user, get the last logged-in user
            $lastLogon = Get-WmiObject -Class Win32_NetworkLoginProfile -ComputerName $ComputerName | 
                         Where-Object {$_.LastLogon -ne $null} | 
                         Sort-Object LastLogon -Descending | 
                         Select-Object -First 1

            if ($lastLogon) {
                $output.Username = $lastLogon.Name
                $output.LogonTime = $lastLogon.LastLogon
                $output.Status = "Last logged in"
            } else {
                $output.Status = "No login information found"
            }
        }

        return New-Object PSObject -Property $output
    }
    catch {
        Write-Error 'Error accessing $ComputerName: $_'
    }
}

# Prompt for the computer name
$computerName = Read-Host "Enter the name of the computer you want to check"

# Check login information for the specified computer
$result = Get-LoggedInUser -ComputerName $computerName

# Display the results
Write-Output "Computer Name: $($result.ComputerName)"
Write-Output "Username: $($result.Username)"
Write-Output "Logon Time: $($result.LogonTime)"
Write-Output "Status: $($result.Status)"