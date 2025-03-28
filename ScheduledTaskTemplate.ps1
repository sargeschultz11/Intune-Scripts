#### ScheduledTaskTemplate.ps1 ####

<#
.SYNOPSIS
Template for creating a scheduled task using an XML definition.

.DESCRIPTION
This script creates a scheduled task using an embedded XML definition. 
Customize the XML content and task name to suit your needs.
Manaully create the task you want to deploy on your local machine and export it.

.PARAMETERS
None. The script does not accept parameters.

.PREREQUISITES
- Must be run with administrator privileges.
- Windows PowerShell 5.1 or later.
- The account running the script must have permissions to create scheduled tasks.

.NOTES
- Task settings should be defined in the embedded XML.
- Designed for environments requiring automated task creation.

.EXAMPLE
.\ScheduledTaskTemplate.ps1

#>

###################
## Embed the XML ##
###################
# Replace this XML with your exported task XML.
$TaskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2024-04-22T13:19:22.5877406</Date>
    <Author>Administrator</Author>
    <URI>\Your Task Name</URI>
  </RegistrationInfo>
  <Triggers>
    <!-- Define triggers here -->
  </Triggers>
  <Principals>
    <!-- Define principals here -->
  </Principals>
  <Settings>
    <!-- Define settings here -->
  </Settings>
  <Actions Context="Author">
    <!-- Define actions here -->
  </Actions>
</Task>
"@

##############################
## Task Management Section ##
##############################

# Define the task name
$TaskName = "Your Task Name"

$LogFilePath = Join-Path -Path $env:TEMP -ChildPath "ScriptLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        default { Write-Host $LogMessage }
    }
    
    Add-Content -Path $LogFilePath -Value $LogMessage
}

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Log "Task '$TaskName' already exists. Removing it before creating a new one..." "WARNING"
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-Log "Task '$TaskName' successfully removed."
    } catch {
        Write-Log "Failed to remove the task '$TaskName'. Error details: $_" "ERROR"
        exit 1
    }
}

Write-Log "Creating task '$TaskName'..."
try {
    Register-ScheduledTask -Xml $TaskXml -TaskName $TaskName -ErrorAction Stop
    Write-Log "Task '$TaskName' successfully created."
} catch {
    Write-Log "Failed to create the task '$TaskName'. Error details: $_" "ERROR"
    exit 1
}

Write-Log "Script execution completed."
