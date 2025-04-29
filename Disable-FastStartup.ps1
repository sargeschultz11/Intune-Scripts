#### Disable-FastStartup.ps1 ####

<#
.SYNOPSIS
Disables Fast Startup by setting the HiberbootEnabled registry value to 0.

.DESCRIPTION
This script modifies the Windows registry to disable the Fast Startup feature. 
Fast Startup can cause issues with dual boot setups, network configurations, or certain hardware. 
By setting the HiberbootEnabled value to 0, the system ensures that Fast Startup is turned off.

.PARAMETERS
None

.PREREQUISITES
- The script must be run in the system context or as an administrator.

.NOTES
- Modifies the registry; ensure you back up the registry before running this script.
- Disabling Fast Startup may slightly increase boot time but improves compatibility with some system configurations.

.EXAMPLE
.\Disable-FastStartup.ps1
Runs the script to disable Fast Startup.
#>
# Define the log file path in the Temp folder
$LogFilePath = Join-Path -Path $env:TEMP -ChildPath "ScriptLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to the log file
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO" # Supports INFO, WARNING, ERROR
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    # Write to the console
    switch ($Type) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        default { Write-Host $LogMessage }
    }
    
    # Write to log file
    Add-Content -Path $LogFilePath -Value $LogMessage
}

$RegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
$ValueName = "HiberbootEnabled"

if (Test-Path $RegistryPath) {
    try {
        # Set the registry value
        Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value 0 -Force
        Write-Log "Successfully set $ValueName to 0 at $RegistryPath."
    } catch {
        Write-Log "Error: Unable to set $ValueName. $_"
    }
} else {
    Write-Log "Error: Registry path $RegistryPath does not exist."
}