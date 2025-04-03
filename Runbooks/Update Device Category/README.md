# Update-IntuneWindowsDeviceCategories.ps1

>**WARNING**: This runbook is no longer supported and has been updated and migrated to another repo.
>You can find the updated version here: [https://github.com/sargeschultz11/Azure-Runbooks/tree/main/DeviceCategorySync](https://github.com/sargeschultz11/Azure-Runbooks/tree/main/DeviceCategorySync)


## Overview
This Azure Automation runbook script automatically updates the device categories of Windows devices in Microsoft Intune based on the primary user's department. It specifically targets devices that either have no category assigned or have a category that doesn't match the user's department.

## Purpose
The primary purpose of this script is to ensure consistent device categorization in Intune by:
- Identifying Windows devices with missing or mismatched categories
- Retrieving the primary user's department information
- Setting the device category to match the user's department when available

This automation helps maintain better organization within the Intune portal and can be used for device targeting, reporting, and policy assignment.

## Prerequisites
- An Azure Automation account
- An Azure AD App Registration with the following:
  - Client ID
  - Client Secret
  - Proper Microsoft Graph API permissions:
    - `DeviceManagementManagedDevices.Read.All`
    - `DeviceManagementManagedDevices.ReadWrite.All`
    - `User.Read.All`
- The following variables defined in the Automation account:
  - `TenantId`: Your Azure AD tenant ID
  - `ClientId`: The App Registration's client ID
  - `ClientSecret`: The App Registration's client secret (stored as an encrypted variable)
- **IMPORTANT**: Device categories must be pre-created in Intune and must match **exactly** the department names in user account properties in Azure AD

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| TenantId | String | No | Your Azure AD tenant ID. If not provided, will be retrieved from Automation variables. |
| ClientId | String | No | The App Registration's client ID. If not provided, will be retrieved from Automation variables. |
| ClientSecret | String | No | The App Registration's client secret. If not provided, will be retrieved from Automation variables. |
| WhatIf | Switch | No | If specified, shows what changes would occur without actually making any updates. |

## Execution Flow
1. **Authentication**: The script authenticates to Microsoft Graph API using the provided client credentials.
2. **Device Category Retrieval**: Retrieves all device categories defined in Intune.
3. **Device Retrieval**: Gets all Windows devices from Intune.
4. **Processing Loop**: For each device:
   - Checks if a device category is already assigned
   - Retrieves the primary user of the device
   - Gets the user's department information
   - If the department exists as a device category and differs from the current device category, updates the device's category

## Output
The script produces a PowerShell custom object with the following properties:

| Property | Description |
|----------|-------------|
| TotalDevices | Total number of Windows devices processed |
| AlreadyCategorized | Number of devices with categories already matching departments |
| Updated | Number of devices that had their categories updated |
| Skipped | Number of devices skipped (no primary user, no department, or department not a category) |
| Errors | Number of devices that encountered errors during processing |
| WhatIfMode | Boolean indicating if WhatIf mode was enabled |

## Logging
The script utilizes verbose logging to provide detailed information about each step:
- All log entries include timestamps and log levels (INFO, WARNING, ERROR, WHATIF)
- Write-Verbose is used for standard logging in Azure Automation
- Specific error cases are captured and logged appropriately

## Usage Examples

### Basic Usage
Run the script without parameters to use the variables defined in the Automation account:
```powershell
Update-IntuneWindowsDeviceCategories
```

### WhatIf Mode
Run the script in WhatIf mode to see what changes would occur without making any updates:
```powershell
Update-IntuneWindowsDeviceCategories -WhatIf
```

### Providing Credentials Directly
Run the script with explicit credentials:
```powershell
Update-IntuneWindowsDeviceCategories -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-client-secret"
```

## Error Handling
The script includes comprehensive error handling:
- Authentication failures are captured and reported
- API request errors are logged with details
- Device processing errors are isolated to prevent the entire script from failing
- Summary statistics include error counts

## Notes
- **CRITICAL REQUIREMENT**: The script depends on exact matching between department names in Azure AD and device category names in Intune. If these don't match exactly, the categorization will not work.
- Before running this script, ensure that all departments used in your organization have corresponding device categories created in Intune with identical naming.
- Devices without primary users or where the user has no department are skipped
- The script counts and reports cases where department names don't exist as device categories
- For devices that already have the correct category assigned, no changes are made
- If department names in Azure AD don't match device categories in Intune exactly (including case, spacing, and special characters), the script will report these as skipped devices

## Author Information
- **Author**: Ryan Schultz
- **Version**: 1.0
- **Creation Date**: 2025-03-26
