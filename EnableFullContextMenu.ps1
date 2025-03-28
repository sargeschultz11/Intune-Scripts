<#
.SYNOPSIS
    Restores the classic context right-click menu in Windows 11.

.DESCRIPTION
    This script creates a registry key in the HKEY_CURRENT_USER hive that
    disables the new Windows 11 context menu and restores the classic (full)
    right-click menu. The script checks if the key already exists before
    attempting to create it.

.NOTES
    File Name      : Restore-ClassicContextMenu.ps1
    Author         : Ryan Schultz
    Prerequisite   : PowerShell 5.0 or later
    Version        : 1.0
    Date           : 2026/03

.EXAMPLE
    .\EnableFullContextMenu.ps1

#>

$registryPath = "HKCU:\SOFTWARE\CLASSES\CLSID\"
$keyName = "{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"

if (-not (Test-Path "$registryPath$keyName")) {
    New-Item -Path "$registryPath$keyName" -Force
    New-Item -Path "$registryPath$keyName\InprocServer32" -Force

    Set-ItemProperty -Path "$registryPath$keyName\InprocServer32" -Name "(Default)" -Value ""
    
    Write-Host "Registry key created successfully. Please restart your computer to apply changes."
} else {
    Write-Host "Registry key already exists."
}

