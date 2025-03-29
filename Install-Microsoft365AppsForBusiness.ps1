<#
.SYNOPSIS
    Installs Microsoft 365 applications using the Office Deployment Tool (ODT).

.DESCRIPTION
    This script automates the download and installation of Microsoft 365 applications using the Office Deployment Tool.
    It can use either an inline XML configuration or a custom XML configuration file from a URL.

.PARAMETER ConfigurationXMLFile
    Optional. URL to a custom XML configuration file for the Office installation.
    If not provided, the script will use a built-in default configuration.
    Example: "https://example.com/office365config.xml"

.PARAMETER OfficeInstallDownloadPath
    Optional. Specifies the directory where the installation files and XML configuration will be stored.
    Default value is "$env:TEMP\Office365Install"

.PARAMETER Restart
    Optional. Switch parameter. If specified, the computer will restart 60 seconds after successful installation.

.EXAMPLE
    # Install using the default configuration
    .\Install-Office365.ps1

.EXAMPLE
    # Install using a custom configuration file from a URL
    .\Install-Office365.ps1 -ConfigurationXMLFile "https://myURL.com/custom-office-config.xml"

.LINK
    https://config.office.com/ - Microsoft's Office Customization Tool for creating XML configurations
#>

[CmdletBinding()]
param(
    [Parameter()]
    [String]$ConfigurationXMLFile,
    
    [Parameter()]
    [String]$OfficeInstallDownloadPath = "$env:TEMP\Office365Install",
    
    [Parameter()]
    [Switch]$Restart
)

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'
$LogFilePath = Join-Path -Path $env:TEMP -ChildPath "OfficeInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if ($env:linkToConfigurationXml -and $env:linkToConfigurationXml -notlike "null") { 
    $ConfigurationXMLFile = $env:linkToConfigurationXml 
}
if ($env:restartComputer -like "true") { 
    $Restart = $True 
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "TLS 1.2 is not supported on this system. This script may fail!"
}

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

function Test-IsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-ODTURL {
    $Uri = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
    
    try {
        $MSWebPage = Invoke-WebRequest -Uri $Uri -UseBasicParsing
        $DownloadURL = $MSWebPage.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | 
                       Select-Object -ExpandProperty href -First 1
        
        if (-not $DownloadURL) {
            Write-Log -Message "Unable to find the download link for the Office Deployment Tool." -Type "ERROR"
            exit 1
        }
        return $DownloadURL
    }
    catch {
        Write-Log -Message "Unable to connect to the Microsoft website: $_" -Type "ERROR"
        exit 1
    }
}

function Invoke-Download {
    param(
        [Parameter(Mandatory=$true)]
        [String]$URL,
        
        [Parameter(Mandatory=$true)]
        [String]$Path
    )
    
    Write-Log -Message "Downloading from '$URL'..."
    
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $URL -OutFile $Path -UseBasicParsing
        $ProgressPreference = 'Continue'
    }
    catch {
        Write-Log -Message "Download failed: $_" -Type "ERROR"
        if (Test-Path -Path $Path) {
            Remove-Item $Path -Force -ErrorAction SilentlyContinue
        }
        exit 1
    }
    
    if (-not (Test-Path $Path)) {
        Write-Log -Message "File failed to download to '$Path'." -Type "ERROR"
        exit 1
    }
    
    return $Path
}

function Set-DefaultXMLFile {
    $OfficeXML = [XML]@"
<Configuration ID="6dbbe3cd-54c5-4e05-9b7f-0123922eb34f">
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="MatchOS" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <RemoveMSI />
</Configuration>
"@
    $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")
}

function Test-OfficeInstalled {
    $UninstallKeys = @(
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($Key in $UninstallKeys) {
        $Result = Get-ItemProperty $Key | Where-Object { $_.DisplayName -like "*Microsoft 365*" }
        if ($Result) {
            return $true
        }
    }
    
    return $false
}

if (-not (Test-IsElevated)) {
    Write-Log -Message "Access Denied. Please run with Administrator privileges." -Type "ERROR"
    exit 1
}

if (-not (Test-Path $OfficeInstallDownloadPath)) {
    New-Item -Path $OfficeInstallDownloadPath -ItemType Directory -Force | Out-Null
}

if (-not $ConfigurationXMLFile) {
    Write-Log -Message "Using default Office configuration XML"
    Set-DefaultXMLFile
}
else {
    Write-Log -Message "Using custom configuration XML from $ConfigurationXMLFile"
    
    if ($ConfigurationXMLFile -notmatch "^https?://") {
        $ConfigurationXMLFile = "https://$ConfigurationXMLFile"
    }
    
    Invoke-Download -URL $ConfigurationXMLFile -Path "$OfficeInstallDownloadPath\OfficeInstall.xml"
    
    try {
        [xml]::new().Load("$OfficeInstallDownloadPath\OfficeInstall.xml")
    }
    catch {
        Write-Log -Message "Invalid XML file: $_" -Type "ERROR"
        exit 1
    }
}

$ConfigurationXMLPath = "$OfficeInstallDownloadPath\OfficeInstall.xml"

Write-Log -Message "Getting Office Deployment Tool download URL..."
$ODTInstallLink = Get-ODTURL

Write-Log -Message "Downloading the Office Deployment Tool..."
Invoke-Download -URL $ODTInstallLink -Path "$OfficeInstallDownloadPath\ODTSetup.exe"

Write-Log -Message "Extracting the Office Deployment Tool..."
try {
    Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait -NoNewWindow
}
catch {
    Write-Log -Message "Error extracting the Office Deployment Tool: $_" -Type "ERROR"
    exit 1
}

try {
    Write-Log -Message "Installing Microsoft 365..."
    $Install = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfigurationXMLPath" -Wait -PassThru -NoNewWindow
    
    if ($Install.ExitCode -ne 0) {
        Write-Log -Message "Installation failed with exit code: $($Install.ExitCode)" -Type "ERROR"
        exit 1
    }
}
catch {
    Write-Log -Message "Error during Office installation: $_" -Type "ERROR"
    exit 1
}

Write-Log -Message "Cleaning up installation files..."
Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse -ErrorAction SilentlyContinue

if (Test-OfficeInstalled) {
    Write-Log -Message "Microsoft 365 installed successfully!"
    
    if ($Restart) {
        Write-Log -Message "Restarting the computer in 60 seconds..."
        Start-Process shutdown.exe -ArgumentList "-r -t 60" -NoNewWindow
    }
    
    exit 0
}
else {
    Write-Log -Message "Microsoft 365 was not detected after the installation!" -Type "ERROR"
    exit 1
}