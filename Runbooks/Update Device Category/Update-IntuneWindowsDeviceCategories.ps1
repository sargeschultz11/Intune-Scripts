<#
.SYNOPSIS
    Updates Intune Windows device categories based on primary user's department when category is missing.
.DESCRIPTION
    This Azure Runbook script authenticates to Microsoft Graph API using a client ID and secret,
    retrieves all Intune Windows devices, and for devices with no category, sets the category
    to match the primary user's department. Includes a -WhatIf parameter for testing without making changes.
.PARAMETER WhatIf
    If specified, shows what changes would occur without actually making any updates.
.NOTES
    File Name: Update-IntuneWindowsDeviceCategories.ps1
    Author: Ryan Schultz
    Version: 1.0
#>

# Parameters for Azure Automation
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Setup logging function using Azure Automation's logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO" # Supports INFO, WARNING, ERROR, WHATIF
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Type] $Message"
    
    switch ($Type) {
        "ERROR" { 
            Write-Error $Message
            Write-Verbose $LogMessage -Verbose
        }
        "WARNING" { 
            Write-Warning $Message 
            Write-Verbose $LogMessage -Verbose
        }
        "WHATIF" { 
            Write-Verbose "[WHATIF] $Message" -Verbose
        }
        default { 
            Write-Verbose $LogMessage -Verbose
        }
    }
}

function Get-MsGraphToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    
    try {
        Write-Log "Attempting to acquire Microsoft Graph API token..."
        
        if ([string]::IsNullOrEmpty($TenantId) -or [string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret)) {
            Write-Log "Using Azure Automation variables for authentication"
            $TenantId = Get-AutomationVariable -Name 'TenantId'
            $ClientId = Get-AutomationVariable -Name 'ClientId'
            $ClientSecret = Get-AutomationVariable -Name 'ClientSecret'
        }
        
        $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        $body = @{
            client_id     = $ClientId
            scope         = "https://graph.microsoft.com/.default"
            client_secret = $ClientSecret
            grant_type    = "client_credentials"
        }
        
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
        Write-Log "Successfully acquired token" 
        return $response.access_token
    }
    catch {
        Write-Log "Failed to acquire token: $_" -Type "ERROR"
        throw "Authentication failed: $_"
    }
}

function Invoke-MsGraphRequest {
    param (
        [string]$Token,
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null,
        [string]$ContentType = "application/json"
    )
    
    try {
        $headers = @{
            Authorization = "Bearer $Token"
        }
        
        $params = @{
            Uri         = $Uri
            Headers     = $headers
            Method      = $Method
            ContentType = $ContentType
        }
        
        if ($null -ne $Body -and $Method -ne "GET") {
            $params.Add("Body", ($Body | ConvertTo-Json -Depth 10))
        }
        
        return Invoke-RestMethod @params
    }
    catch {
        Write-Log "Graph API request failed: $_" -Type "ERROR"
        throw "Graph API request failed: $_"
    }
}

function Get-IntuneWindowsDevices {
    param (
        [string]$Token
    )
    
    try {
        Write-Log "Retrieving Intune Windows devices..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'"
        $devices = @()
        $response = Invoke-MsGraphRequest -Token $Token -Uri $uri
        
        $devices += $response.value
        
        while ($null -ne $response.'@odata.nextLink') {
            Write-Log "Retrieving next page of devices..."
            $response = Invoke-MsGraphRequest -Token $Token -Uri $response.'@odata.nextLink'
            $devices += $response.value
        }
        
        Write-Log "Retrieved $($devices.Count) Windows devices from Intune"
        return $devices
    }
    catch {
        Write-Log "Failed to retrieve Intune devices: $_" -Type "ERROR"
        throw "Failed to retrieve Intune devices: $_"
    }
}

function Get-IntuneDeviceCategories {
    param (
        [string]$Token
    )
    
    try {
        Write-Log "Retrieving device categories..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories"
        $categories = Invoke-MsGraphRequest -Token $Token -Uri $uri
        Write-Log "Retrieved $($categories.value.Count) device categories"
        
        $categoryLookup = @{}
        foreach ($category in $categories.value) {
            $categoryLookup[$category.displayName] = $category.id
        }
        
        return $categoryLookup
    }
    catch {
        Write-Log "Failed to retrieve device categories: $_" -Type "ERROR"
        throw "Failed to retrieve device categories: $_"
    }
}

function Get-DevicePrimaryUser {
    param (
        [string]$Token,
        [string]$DeviceId
    )
    
    try {
        Write-Log "Retrieving primary user for device $DeviceId..."
        $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/users"
        $response = Invoke-MsGraphRequest -Token $Token -Uri $uri
        
        if ($response.value.Count -gt 0) {
            return $response.value[0]
        }
        else {
            Write-Log "No primary user found for device $DeviceId" -Type "WARNING"
            return $null
        }
    }
    catch {
        Write-Log "Failed to retrieve primary user for device $DeviceId`: $_" -Type "ERROR"
        return $null
    }
}

function Get-UserDetails {
    param (
        [string]$Token,
        [string]$UserId
    )
    
    try {
        Write-Log "Retrieving details for user $UserId..."
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId`?`$select=id,displayName,department"
        $user = Invoke-MsGraphRequest -Token $Token -Uri $uri
        
        return $user
    }
    catch {
        Write-Log "Failed to retrieve details for user $UserId`: $_" -Type "ERROR"
        return $null
    }
}

function Update-DeviceCategory {
    param (
        [string]$Token,
        [string]$DeviceId,
        [string]$CategoryId,
        [string]$CategoryName,
        [switch]$WhatIf
    )
    
    try {
        if ($WhatIf) {
            Write-Log "Would update device category for device $DeviceId to: $CategoryName (ID: $CategoryId)" -Type "WHATIF"
            return $true
        }
        else {
            Write-Log "Updating device category for device $DeviceId to: $CategoryName (ID: $CategoryId)"
            
            $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceCategory/`$ref"
            $body = @{
                "@odata.id" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$CategoryId"
            }
            
            $jsonBody = $body | ConvertTo-Json -Depth 10
            Write-Log "Request body: $jsonBody"
            Write-Log "Request URI: $uri"
            
            Invoke-MsGraphRequest -Token $Token -Uri $uri -Method "PUT" -Body $body
            Write-Log "Successfully updated device category"
            return $true
        }
    }
    catch {
        Write-Log "Failed to update device category for device $DeviceId`: $_" -Type "ERROR"
        return $false
    }
}

try {
    if ($WhatIf) {
        Write-Log "=== WHATIF MODE ENABLED - NO CHANGES WILL BE MADE ===" -Type "WHATIF"
    }
    
    Write-Log "=== Intune Device Category Update Started ==="
    
    $token = Get-MsGraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    
    $categoryLookup = Get-IntuneDeviceCategories -Token $token
    
    Write-Log "Available device categories:"
    foreach ($cat in $categoryLookup.Keys) {
        Write-Log "- $cat (ID: $($categoryLookup[$cat]))"
    }
    
    $windowsDevices = Get-IntuneWindowsDevices -Token $token
    
    $updatedCount = 0
    $errorCount = 0
    $skippedCount = 0
    $matchCount = 0
    
    foreach ($device in $windowsDevices) {
        try {
            $deviceName = $device.deviceName
            $deviceId = $device.id
            $category = $device.deviceCategoryDisplayName
            
            Write-Log "Processing device: $deviceName (ID: $deviceId)"
            Write-Log "Current Category: '$category'"
            
            $primaryUser = Get-DevicePrimaryUser -Token $token -DeviceId $deviceId
            
            if ($null -ne $primaryUser) {
                $userDetails = Get-UserDetails -Token $token -UserId $primaryUser.id
                
                if ($null -ne $userDetails -and -not [string]::IsNullOrEmpty($userDetails.department)) {
                    $userDepartment = $userDetails.department
                    Write-Log "Found department '$userDepartment' for user $($userDetails.displayName)"
                    
                    if ($categoryLookup.ContainsKey($userDepartment)) {
                        $categoryId = $categoryLookup[$userDepartment]
                        
                        if ([string]::IsNullOrEmpty($category) -or 
                            $category -eq "Unassigned" -or 
                            $category -eq "Unknown" -or 
                            $category -ne $userDepartment) {
                            
                            if (![string]::IsNullOrEmpty($category) -and 
                                $category -ne "Unassigned" -and 
                                $category -ne "Unknown" -and 
                                $category -ne $userDepartment) {
                                Write-Log "Device $deviceName has category '$category' which doesn't match user department '$userDepartment'. Updating..." -Type "WARNING"
                            } else {
                                Write-Log "Device $deviceName has no valid category assigned. Updating to match user department..." -Type "WARNING"
                            }
                            
                            $updateResult = Update-DeviceCategory -Token $token -DeviceId $deviceId -CategoryId $categoryId -CategoryName $userDepartment -WhatIf:$WhatIf
                            
                            if ($updateResult) {
                                if ($WhatIf) {
                                    Write-Log "Would have updated device category to '$userDepartment' based on user department for device $deviceName" -Type "WHATIF"
                                }
                                else {
                                    Write-Log "Successfully updated device category to '$userDepartment' based on user department for device $deviceName"
                                }
                                $updatedCount++
                            }
                            else {
                                Write-Log "Failed to update device category for device $deviceName" -Type "ERROR"
                                $errorCount++
                            }
                        }
                        else {
                            Write-Log "Device $deviceName already has category set to '$category' which matches user department. No action needed."
                            $matchCount++
                        }
                    }
                    else {
                        Write-Log "Department '$userDepartment' does not exist as a device category in Intune. Skipping." -Type "WARNING"
                        $skippedCount++
                    }
                }
                else {
                    Write-Log "No department information found for the primary user of device $deviceName. Skipping." -Type "WARNING"
                    $skippedCount++
                }
            }
            else {
                Write-Log "No primary user found for device $deviceName. Keeping existing category." -Type "WARNING"
                $skippedCount++
            }
        }
        catch {
            Write-Log "Error processing device $($device.deviceName): $_" -Type "ERROR"
            $errorCount++
        }
    }
    
    Write-Log "=== Intune Device Category Update Completed ==="
    if ($WhatIf) {
        Write-Log "=== WHATIF SUMMARY - NO CHANGES WERE MADE ===" -Type "WHATIF"
    }
    Write-Log "Devices processed: $($windowsDevices.Count)"
    Write-Log "Already categorized: $matchCount"

    if ($WhatIf) {
        Write-Log "Would be updated: $updatedCount" -Type "WHATIF"
    } else {
        Write-Log "Updated: $updatedCount"
    }

    Write-Log "Skipped (no primary user, no department, or department not a category): $skippedCount"
    Write-Log "Errors: $errorCount"

    $outputObject = [PSCustomObject]@{
        TotalDevices = $windowsDevices.Count
        AlreadyCategorized = $matchCount
        Updated = $updatedCount
        Skipped = $skippedCount
        Errors = $errorCount
        WhatIfMode = $WhatIf
    }
    
    $outputObject
}
catch {
    Write-Log "Script execution failed: $_" -Type "ERROR"
    throw $_
}
finally {
    Write-Log "Script execution completed"
}