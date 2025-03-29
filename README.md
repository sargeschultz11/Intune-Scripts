# Intune-Scripts

A collection of PowerShell platform scripts designed for Microsoft Intune deployment.

## Overview

This repository contains customized PowerShell scripts that can be used as Intune platform scripts for various administrative and deployment tasks in Windows environments. These scripts are designed to be deployed through Intune as platform scripts, or Win32 app installations.

## Scripts

### Add-LAPSuser.ps1
Creates a local administrator account intended for use with Local Administrator Password Solution (LAPS). This script:
- Creates a new local user account with a randomly generated password
- Adds the account to the local Administrators group
- Sets the password to never expire
- Logs all actions for troubleshooting

### EnableFullContextMenu.ps1
Restores the classic context menu in Windows 11 by modifying registry settings. This script:
- Creates necessary registry keys to disable the new Windows 11 context menu
- Checks if the key already exists before attempting to create it
- Prompts for a system restart to apply changes

### ScheduledTaskTemplate.ps1
Provides a template for creating Windows scheduled tasks using an XML definition. This script:
- Creates a scheduled task using an embedded XML definition
- Removes any existing task with the same name before creating a new one
- Logs all actions for troubleshooting

## Usage

These scripts are designed to be deployed through Microsoft Intune as platform scripts. To use them:

1. Upload the script to Intune
2. Configure the script settings (run as system/user, etc.)
3. Assign the script to the appropriate device groups
4. Monitor script execution results through Intune

For detailed information about each script, refer to the script header documentation.

## Requirements

- Windows PowerShell 5.1 or later
- Administrative privileges
- Microsoft Intune environment for deployment

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Ryan Schultz

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue.