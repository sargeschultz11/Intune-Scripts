#### Remove-OneStart.ps1 ####
<#
.SYNOPSIS
Removes OneStart.ai software and its components from Windows systems.

.DESCRIPTION
This script performs a comprehensive cleanup of OneStart.ai software, which is classified as a potentially unwanted program (PUP) 
or adware by multiple security vendors. The script removes all traces of OneStart including running processes, 
installed files, registry entries, scheduled tasks, and browser extensions.

The script includes detailed logging functionality that records all actions and results to both 
the console (with color-coded status indicators) and a log file for later review.

.PARAMETERS
None - This script does not accept parameters but runs with full removal functionality by default.

.PREREQUISITES
- Windows PowerShell 5.1 or later
- Administrative privileges (must be run as Administrator)
- Access to the Windows Registry
- Access to the Windows Task Scheduler

.NOTES
Author: Ryan Schultz
Version: 1.0
Last Updated: April 29, 2025

The script performs the following actions:
- Terminates OneStart related processes
- Removes OneStart directories from all user profiles
- Deletes OneStart registry entries
- Removes OneStart scheduled tasks
- Checks for and removes browser extensions related to OneStart
- Creates a detailed log file and summary report

The script creates a timestamped log file in the TEMP directory with detailed information about 
the removal process, including successes, failures, and a summary of actions taken.

A system restart is recommended after running this script to ensure complete removal.

.EXAMPLE
.\Remove-OneStart.ps1

Run the script with administrative privileges to remove all OneStart components from the system.

.EXAMPLE
.\Remove-OneStart.ps1

powershell.exe -ExecutionPolicy Bypass -File .\Remove-OneStart.ps1

Run the script with execution policy bypass (useful for deployment via Group Policy or Intune).
#>

# Create a log file
$logFile = Join-Path -Path $env:TEMP -ChildPath "OneStartRemoval_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-Log {
    param(
        [string]$message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "$timestamp - [$level] $message"
    $logMessage | Out-File -FilePath $logFile -Append
    
    switch ($level) {
        'INFO'    { Write-Host $logMessage -ForegroundColor White }
        'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
        'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
    }
}

# Track statistics
$script:stats = @{
    ProcessesStopped = 0
    DirectoriesRemoved = 0
    RegistryKeysRemoved = 0
    ScheduledTasksRemoved = 0
    ExtensionsRemoved = 0
    ErrorsEncountered = 0
}

Write-Log "Starting OneStart.ai removal process..." 'INFO'
Write-Log "Log file created at: $logFile" 'INFO'

# 1. Terminate related processes
Write-Log "Terminating OneStart related processes..." 'INFO'
$processNames = @("DBar", "OneStart", "OneStartBar", "OneStartTray", "Chromium", "Quick Updater")
$processesFound = $false

foreach ($proc in $processNames) {
    $processes = Get-Process -Name $proc -ErrorAction SilentlyContinue
    if ($processes) {
        $processesFound = $true
        try {
            $processes | Stop-Process -Force
            Write-Log "Stopped process: $proc (Count: $($processes.Count))" 'SUCCESS'
            $script:stats.ProcessesStopped += $processes.Count
        } catch {
            Write-Log "Failed to stop process: $proc - $_" 'ERROR'
            $script:stats.ErrorsEncountered++
        }
    }
}

if (-not $processesFound) {
    Write-Log "No OneStart related processes found running." 'INFO'
}

# 2. Remove OneStart directories for all users
Write-Log "Removing OneStart directories..." 'INFO'
$filePaths = @(
    "\AppData\Roaming\OneStart\",
    "\AppData\Local\OneStart.ai",
    "\AppData\Local\OneStart*\"
)
$directoriesFound = $false

foreach ($userFolder in Get-ChildItem C:\Users -Directory) {
    $userDirectoriesFound = $false
    
    foreach ($path in $filePaths) {
        $fullPath = Join-Path $userFolder.FullName $path
        if ($fullPath -like "*`**") {
            $basePath = Split-Path $fullPath
            $pattern = Split-Path $fullPath -Leaf
            $matchingPaths = Get-ChildItem -Path $basePath -Directory -Filter $pattern -ErrorAction SilentlyContinue
            
            if ($matchingPaths) {
                $directoriesFound = $true
                $userDirectoriesFound = $true
                
                foreach ($matchPath in $matchingPaths) {
                    try {
                        Remove-Item -Path $matchPath.FullName -Recurse -Force -ErrorAction Stop
                        Write-Log "Deleted: $($matchPath.FullName)" 'SUCCESS'
                        $script:stats.DirectoriesRemoved++
                    } catch {
                        Write-Log "Failed to delete: $($matchPath.FullName) - $_" 'ERROR'
                        $script:stats.ErrorsEncountered++
                    }
                }
            }
        } 
        elseif (Test-Path $fullPath) {
            $directoriesFound = $true
            $userDirectoriesFound = $true
            
            try {
                Remove-Item -Path $fullPath -Recurse -Force -ErrorAction Stop
                Write-Log "Deleted: $fullPath" 'SUCCESS'
                $script:stats.DirectoriesRemoved++
            } catch {
                Write-Log "Failed to delete: $fullPath - $_" 'ERROR'
                $script:stats.ErrorsEncountered++
            }
        }
    }
    
    if ($userDirectoriesFound) {
        Write-Log "Removed OneStart directories for user: $($userFolder.Name)" 'INFO'
    }
}

if (-not $directoriesFound) {
    Write-Log "No OneStart directories found." 'INFO'
}

# 3. Remove OneStart registry keys
Write-Log "Removing OneStart registry entries..." 'INFO'
$registryKeysFound = $false

foreach ($registryHive in Get-ChildItem Registry::HKEY_USERS) {
    $userSID = Split-Path $registryHive.Name -Leaf
    $userKeysFound = $false
    
    $regPath = "$($registryHive.PSPath)\Software\OneStart.ai"
    if (Test-Path $regPath) {
        $registryKeysFound = $true
        $userKeysFound = $true
        try {
            Remove-Item -Path $regPath -Recurse -Force
            Write-Log "Removed registry key: $regPath" 'SUCCESS'
            $script:stats.RegistryKeysRemoved++
        } catch {
            Write-Log "Failed to remove registry key: $regPath - $_" 'ERROR'
            $script:stats.ErrorsEncountered++
        }
    }
    
    $runKeyPath = "$($registryHive.PSPath)\Software\Microsoft\Windows\CurrentVersion\Run"
    $regProperties = @("OneStartBar", "OneStartBarUpdate", "OneStartUpdate")
    $runKeysFound = $false
    
    if (Test-Path $runKeyPath) {
        foreach ($property in $regProperties) {
            if (Get-ItemProperty -Path $runKeyPath -Name $property -ErrorAction SilentlyContinue) {
                $registryKeysFound = $true
                $userKeysFound = $true
                $runKeysFound = $true
                try {
                    Remove-ItemProperty -Path $runKeyPath -Name $property
                    Write-Log "Removed registry value: $property from $runKeyPath" 'SUCCESS'
                    $script:stats.RegistryKeysRemoved++
                } catch {
                    Write-Log "Failed to remove registry value: $property from $runKeyPath - $_" 'ERROR'
                    $script:stats.ErrorsEncountered++
                }
            }
        }
        
        if (-not $runKeysFound) {
            Write-Log "No OneStart Run keys found for user SID: $userSID" 'INFO'
        }
    }
    
    if ($userKeysFound) {
        Write-Log "Removed OneStart registry keys for user SID: $userSID" 'INFO'
    } else {
        Write-Log "No OneStart registry keys found for user SID: $userSID" 'INFO'
    }
}

$additionalRegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OneStart*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OneStart*"
)
$additionalKeysFound = $false

foreach ($regPath in $additionalRegPaths) {
    if ($regPath -like "*`**") {
        $basePath = Split-Path $regPath
        $pattern = (Split-Path $regPath -Leaf).Replace('*', '*')
        $matchingPaths = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue | 
                         Where-Object { $_.PSChildName -like $pattern }
        
        if ($matchingPaths) {
            $registryKeysFound = $true
            $additionalKeysFound = $true
            
            foreach ($matchPath in $matchingPaths) {
                try {
                    Remove-Item -Path $matchPath.PSPath -Recurse -Force
                    Write-Log "Removed registry key: $($matchPath.PSPath)" 'SUCCESS'
                    $script:stats.RegistryKeysRemoved++
                } catch {
                    Write-Log "Failed to remove registry key: $($matchPath.PSPath) - $_" 'ERROR'
                    $script:stats.ErrorsEncountered++
                }
            }
        }
    }
    elseif (Test-Path $regPath) {
        $registryKeysFound = $true
        $additionalKeysFound = $true
        
        try {
            Remove-Item -Path $regPath -Recurse -Force
            Write-Log "Removed registry key: $regPath" 'SUCCESS'
            $script:stats.RegistryKeysRemoved++
        } catch {
            Write-Log "Failed to remove registry key: $regPath - $_" 'ERROR'
            $script:stats.ErrorsEncountered++
        }
    }
}

if (-not $additionalKeysFound) {
    Write-Log "No additional OneStart registry keys found in HKLM." 'INFO'
}

if (-not $registryKeysFound) {
    Write-Log "No OneStart registry entries found." 'INFO'
}

# 4. Remove scheduled tasks
Write-Log "Removing OneStart scheduled tasks..." 'INFO'
$scheduledTasks = @("OneStart Chromium", "OneStart Updater")
$tasksFound = $false

foreach ($task in $scheduledTasks) {
    try {
        $exists = Get-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue
        if ($exists) {
            $tasksFound = $true
            Unregister-ScheduledTask -TaskName $task -Confirm:$false
            Write-Log "Removed scheduled task: $task" 'SUCCESS'
            $script:stats.ScheduledTasksRemoved++
        }
    } catch {
        Write-Log "Failed to remove scheduled task: $task - $_" 'ERROR'
        $script:stats.ErrorsEncountered++
    }
}

try {
    $additionalTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*OneStart*" -or $_.TaskPath -like "*OneStart*" } -ErrorAction SilentlyContinue
    
    if ($additionalTasks) {
        $tasksFound = $true
        foreach ($task in $additionalTasks) {
            try {
                Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false
                Write-Log "Removed additional scheduled task: $($task.TaskName) (Path: $($task.TaskPath))" 'SUCCESS'
                $script:stats.ScheduledTasksRemoved++
            } catch {
                Write-Log "Failed to remove additional scheduled task: $($task.TaskName) - $_" 'ERROR'
                $script:stats.ErrorsEncountered++
            }
        }
    }
} catch {
    Write-Log "Error searching for additional OneStart scheduled tasks: $_" 'ERROR'
    $script:stats.ErrorsEncountered++
}

if (-not $tasksFound) {
    Write-Log "No OneStart scheduled tasks found." 'INFO'
}

# 5. Remove browser extensions (Chrome, Edge, Firefox)
Write-Log "Checking for browser extensions..." 'INFO'
$extensionsFound = $false

# Chrome and Edge
$chromeBasedBrowsers = @{
    "Chrome" = "$env:LOCALAPPDATA\Google\Chrome\User Data";
    "Edge" = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
}

foreach ($browser in $chromeBasedBrowsers.Keys) {
    $browserPath = $chromeBasedBrowsers[$browser]
    $browserExtensionsFound = $false
    
    if (Test-Path $browserPath) {
        $profileDirs = Get-ChildItem -Path $browserPath -Directory | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile*" }
        
        foreach ($profile in $profileDirs) {
            $extensionsPath = Join-Path $profile.FullName "Extensions"
            $profileExtensionsFound = $false
            
            if (Test-Path $extensionsPath) {
                Get-ChildItem -Path $extensionsPath -Directory | ForEach-Object {
                    try {
                        $manifestPath = Join-Path $_.FullName "manifest.json"
                        if (Test-Path $manifestPath) {
                            $manifestContent = Get-Content -Path $manifestPath -Raw -ErrorAction SilentlyContinue
                            if ($manifestContent -match "OneStart") {
                                $extensionsFound = $true
                                $browserExtensionsFound = $true
                                $profileExtensionsFound = $true
                                
                                try {
                                    Remove-Item -Path $_.FullName -Recurse -Force
                                    Write-Log "Removed $browser extension: $($_.Name) in profile $($profile.Name)" 'SUCCESS'
                                    $script:stats.ExtensionsRemoved++
                                } catch {
                                    Write-Log "Failed to remove $browser extension: $($_.Name) in profile $($profile.Name) - $_" 'ERROR'
                                    $script:stats.ErrorsEncountered++
                                }
                            }
                        }
                    } catch {
                        Write-Log "Error checking $browser extension manifest in profile $($profile.Name): $_" 'ERROR'
                        $script:stats.ErrorsEncountered++
                    }
                }
            }
            
            if (-not $profileExtensionsFound) {
                Write-Log "No OneStart extensions found in $browser profile: $($profile.Name)" 'INFO'
            }
        }
        
        if (-not $browserExtensionsFound) {
            Write-Log "No OneStart extensions found in $browser." 'INFO'
        }
    } else {
        Write-Log "$browser not installed or not found at the default location." 'INFO'
    }
}

# Firefox
$firefoxPath = "$env:APPDATA\Mozilla\Firefox\Profiles"
$firefoxExtensionsFound = $false

if (Test-Path $firefoxPath) {
    Get-ChildItem -Path $firefoxPath -Directory | ForEach-Object {
        $profileName = $_.Name
        $extensionsPath = Join-Path $_.FullName "extensions"
        $profileExtensionsFound = $false
        
        if (Test-Path $extensionsPath) {
            $oneStartExtensions = Get-ChildItem -Path $extensionsPath -File | Where-Object { $_.Name -like "*onestart*" }
            
            if ($oneStartExtensions) {
                $extensionsFound = $true
                $firefoxExtensionsFound = $true
                $profileExtensionsFound = $true
                
                foreach ($extension in $oneStartExtensions) {
                    try {
                        Remove-Item -Path $extension.FullName -Force
                        Write-Log "Removed Firefox extension: $($extension.Name) in profile $profileName" 'SUCCESS'
                        $script:stats.ExtensionsRemoved++
                    } catch {
                        Write-Log "Failed to remove Firefox extension: $($extension.Name) in profile $profileName - $_" 'ERROR'
                        $script:stats.ErrorsEncountered++
                    }
                }
            }
        }
        
        if (-not $profileExtensionsFound) {
            Write-Log "No OneStart extensions found in Firefox profile: $profileName" 'INFO'
        }
    }
    
    if (-not $firefoxExtensionsFound) {
        Write-Log "No OneStart extensions found in Firefox." 'INFO'
    }
} else {
    Write-Log "Firefox not installed or not found at the default location." 'INFO'
}

if (-not $extensionsFound) {
    Write-Log "No OneStart browser extensions found." 'INFO'
}

# 6. Display summary report
Write-Log "OneStart removal process completed." 'SUCCESS'
Write-Log "-------- Summary Report --------" 'INFO'
Write-Log "Processes stopped: $($script:stats.ProcessesStopped)" 'INFO'
Write-Log "Directories removed: $($script:stats.DirectoriesRemoved)" 'INFO'
Write-Log "Registry keys removed: $($script:stats.RegistryKeysRemoved)" 'INFO'
Write-Log "Scheduled tasks removed: $($script:stats.ScheduledTasksRemoved)" 'INFO'
Write-Log "Browser extensions removed: $($script:stats.ExtensionsRemoved)" 'INFO'
Write-Log "Errors encountered: $($script:stats.ErrorsEncountered)" $(if($script:stats.ErrorsEncountered -gt 0){'WARNING'}else{'INFO'})
Write-Log "--------------------------------" 'INFO'
Write-Log "A system restart is recommended to complete the removal process." 'INFO'
Write-Log "Log file created at: $logFile" 'INFO'