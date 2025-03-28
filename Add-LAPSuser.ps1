<#
.SYNOPSIS
Creates a local user account with a randomly generated password and adds it to the local Administrators group to be targeted by LAPS.

.DESCRIPTION
This script was made for use with Intune but can be adapted. This script creates a new local user account with a 10-character random password. The account is added to the local Administrators group. The password is set to never expire for convenience, making this script useful in scenarios like implementing Local Administrator Password Solution (LAPS) configurations.
Deploy this script from Intune as a Platform script. Run as system.

.PARAMETERS
None.

.PREREQUISITES
- The script must be run with administrative privileges.
- Modify the `$Username` variable to set the desired account name before running the script.

.NOTES
- The script sets a random password using alphanumeric characters (A-Z, a-z, 0-9).
- By default, the execution policy is bypassed for the session to ensure the script runs without restrictions.
- Script checks if account already exists and exits without changes if it does.
- Designed for automated deployment through Intune with no user interaction.
- Ensure the account creation complies with your organization's policies for local administrator accounts.

.EXAMPLE
# Run the script to create a LAPS user
.\Add-LAPSuser.ps1

#>

$LogFilePath = Join-Path -Path $env:TEMP -ChildPath "ScriptLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    # Write to the console
    switch ($Type) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        default { Write-Host $LogMessage }
    }
    
    Add-Content -Path $LogFilePath -Value $LogMessage
}

Write-Log "Add-LAPSuser.ps1 started..."

try {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    Write-Log "Execution policy set to Bypass for current process"
} 
catch {
    Write-Log "Failed to set execution policy: $_" -Type "ERROR"
    exit 1
}

$Username = "local.scs"
$password = -join ((65..90) + (97..122) + (48..57) | Get-Random -Count 10 | ForEach-Object {[char]$_})

Write-Log "Checking if user $Username already exists..."
if (Get-LocalUser -Name $Username -ErrorAction SilentlyContinue) {
    Write-Log "User $Username already exists! No changes made." -Type "INFO"
    Write-Log "Script execution completed - account already exists"
    exit 0
} 
else {
    Write-Log "Creating new LAPS account..."
    try {
        $User = New-LocalUser -Name $Username -Password (ConvertTo-SecureString -String $password -AsPlainText -Force) -PasswordNeverExpires -ErrorAction Stop
        Write-Log "User account $Username created successfully!"
        Write-Log "Password: $password" -Type "INFO"
        
        try {
            $Group = Get-LocalGroup -Name "Administrators"
            Add-LocalGroupMember -Group $Group -Member $User -ErrorAction Stop
            Write-Log "User $Username added to Administrators group"
            Write-Log "LAPS account creation completed successfully!"
        }
        catch {
            Write-Log "Failed to add user to Administrators group: $_" -Type "ERROR"
            Write-Log "Attempting to clean up by removing the created user account..." -Type "WARNING"
            
            try {
                Remove-LocalUser -Name $Username -ErrorAction Stop
                Write-Log "User account $Username removed successfully" -Type "INFO"
            }
            catch {
                Write-Log "Failed to remove user account: $_" -Type "ERROR"
            }
            
            exit 1
        }
    }
    catch {
        Write-Log "Failed to create user account: $_" -Type "ERROR"
        exit 1
    }
}

Write-Log "Script execution completed"