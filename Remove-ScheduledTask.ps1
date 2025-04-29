#### Remove-ScheduledTask.ps1 ####

<#
.SYNOPSIS
Removes a scheduled task by its name.

.DESCRIPTION
This script removes a specified scheduled task from a Windows system by its hardcoded name.

.PARAMETERS
None.

.PREREQUISITES
- The script must be run in the system context or ran as administrator.
- The Task Scheduler service must be running.

.NOTES
- This script uses the `Unregister-ScheduledTask` cmdlet, which is available in Windows PowerShell.
- If the task is not found, the script will log a warning but won't report an error.
- The task name is hardcoded into the script and should be updated as needed.

.EXAMPLE
.\Remove-ScheduledTask.ps1
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

# Enter your task name here
$TaskName = "Your Task Name"

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "This script must be run as an administrator." -Type "ERROR"
    exit 1
}

Write-Log "Starting removal of scheduled task: $TaskName"

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully removed scheduled task: $TaskName"
    } catch {
        Write-Log "Failed to remove task '$TaskName': $_" -Type "ERROR"
        exit 1
    }
} else {
    Write-Log "Task '$TaskName' was not found on this system." -Type "WARNING"
    exit 0
}

Write-Log "Script execution completed."