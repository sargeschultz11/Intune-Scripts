<#
.SYNOPSIS
    Installs Microsoft 365 applications using the Office Deployment Tool (ODT).

.DESCRIPTION
    This script automates the download and installation of Microsoft 365 applications using the Office Deployment Tool.
    It can use either an inline XML configuration or a custom XML configuration file from a URL.
    The script handles downloading the ODT, extracting it, and running the installation with the specified configuration.

.PARAMETER ConfigurationXMLFile
    Optional. URL to a custom XML configuration file for the Office installation.
    If not provided, the script will use a built-in default configuration.
    The URL must include http:// or https:// (the script will add https:// if missing).
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

.EXAMPLE
    # Install with custom download path and restart computer after installation
    .\Install-Office365.ps1 -OfficeInstallDownloadPath "C:\O365Install" -Restart

.NOTES
    
    This script will:
    1. Check for administrative privileges
    2. Create the installation directory if it doesn't exist
    3. Use either the provided configuration XML or the built-in default
    4. Download the latest Office Deployment Tool
    5. Extract the ODT
    6. Execute the installation process
    7. Verify installation success
    8. Clean up installation files (by default)
    9. Restart the computer if specified
    
    Environment Variables:
    - linkToConfigurationXml: Can be used to specify a configuration XML URL
    - restartComputer: Set to "true" to force restart

.LINK
    https://config.office.com/ - Microsoft's Office Customization Tool for creating XML configurations
    https://learn.microsoft.com/en-us/deployoffice/overview-office-deployment-tool - ODT Documentation
#>

[CmdletBinding()]
param(
    # Use a existing config file
    [Parameter()]
    [String]$ConfigurationXMLFile,
    # Path where we will store our install files and our XML file
    [Parameter()]
    [String]$OfficeInstallDownloadPath = "$env:TEMP\Office365Install",
    [Parameter()]
    [Switch]$Restart
)


begin {
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

    if ($env:linkToConfigurationXml -and $env:linkToConfigurationXml -notlike "null") { $ConfigurationXMLFile = $env:linkToConfigurationXml }
    if ($env:restartComputer -like "true") { $Restart = $True }


    $CleanUpInstallFiles = $True


    if ($ConfigurationXMLFile -and $ConfigurationXMLFile -notmatch "^http(s)?://") {
        Write-Log -Message "http(s):// is required to download the file. Adding https:// to your input...." -Type "WARNING"
        $ConfigurationXMLFile = "https://$ConfigurationXMLFile"
        Write-Log -Message "New Url $ConfigurationXMLFile." -Type "WARNING"
    }


    $SupportedTLSversions = [enum]::GetValues('Net.SecurityProtocolType')
    if ( ($SupportedTLSversions -contains 'Tls13') -and ($SupportedTLSversions -contains 'Tls12') ) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol::Tls13 -bor [System.Net.SecurityProtocolType]::Tls12
    }
    elseif ( $SupportedTLSversions -contains 'Tls12' ) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    }
    else {
        Write-Log -Message "TLS 1.2 and or TLS 1.3 are not supported on this system. This script may fail!" -Type "WARNING"
        if ($PSVersionTable.PSVersion.Major -lt 3) {
            Write-Log -Message "PowerShell 2 / .NET 2.0 doesn't support TLS 1.2." -Type "WARNING"
        }
    }


    function Set-XMLFile {
        # XML data that will be used for the download/install
        # Example config below generated from https://config.office.com/
        # To use your own config, just replace <Configuration> to </Configuration> with your xml config file content.
        # Notes:
        #  "@ can not have any character after it
        #  @" can not have any spaces or character before it.
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
    function Get-ODTURL {
        $Uri = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
        $DownloadURL = ""
        for ($i = 1; $i -le 3; $i++) {
            try {
                $MSWebPage = Invoke-WebRequest -Uri $Uri -UseBasicParsing -MaximumRedirection 10
                $DownloadURL = $MSWebPage.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | Select-Object -ExpandProperty href -First 1
                if ($DownloadURL) {
                    break
                }
                Write-Log -Message "Unable to find the download link for the Office Deployment Tool at: $Uri. Attempt $i of 3." -Type "WARNING"
                Start-Sleep -Seconds $($i * 30)
            }
            catch {
                Write-Log -Message "Unable to connect to the Microsoft website. Attempt $i of 3." -Type "WARNING"
            }
        }
        
        if (-not $DownloadURL) {
            Write-Log -Message "Unable to find the download link for the Office Deployment Tool at: $Uri" -Type "ERROR"
            exit 1
        }
        return $DownloadURL
    }
    function Test-IsElevated {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()


        $p = New-Object System.Security.Principal.WindowsPrincipal($id)


        $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }


    # Utility function for downloading files.
    function Invoke-Download {
        param(
            [Parameter()]
            [String]$URL,
            [Parameter()]
            [String]$Path,
            [Parameter()]
            [int]$Attempts = 3,
            [Parameter()]
            [Switch]$SkipSleep
        )
    
        Write-Log -Message "URL '$URL' was given."
        Write-Log -Message "Downloading the file..."


        $SupportedTLSversions = [enum]::GetValues('Net.SecurityProtocolType')
        if ( ($SupportedTLSversions -contains 'Tls13') -and ($SupportedTLSversions -contains 'Tls12') ) {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol::Tls13 -bor [System.Net.SecurityProtocolType]::Tls12
        }
        elseif ( $SupportedTLSversions -contains 'Tls12' ) {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        }
        else {
            Write-Log -Message "TLS 1.2 and/or TLS 1.3 are not supported on this system. This download may fail!" -Type "WARNING"
            if ($PSVersionTable.PSVersion.Major -lt 3) {
                Write-Log -Message "PowerShell 2 / .NET 2.0 doesn't support TLS 1.2." -Type "WARNING"
            }
        }


        $i = 1
        While ($i -le $Attempts) {
            if (!($SkipSleep)) {
                $SleepTime = Get-Random -Minimum 3 -Maximum 15
                Write-Log -Message "Waiting for $SleepTime seconds."
                Start-Sleep -Seconds $SleepTime
            }
            
            if ($i -ne 1) { Write-Log -Message "" }
            Write-Log -Message "Download Attempt $i"


            $PreviousProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try {
                if ($PSVersionTable.PSVersion.Major -lt 4) {
                    $WebClient = New-Object System.Net.WebClient
                    $WebClient.DownloadFile($URL, $Path)
                }
                else {
                    $WebRequestArgs = @{
                        Uri                = $URL
                        OutFile            = $Path
                        MaximumRedirection = 10
                        UseBasicParsing    = $true
                    }


                    Invoke-WebRequest @WebRequestArgs
                }


                $File = Test-Path -Path $Path -ErrorAction SilentlyContinue
            }
            catch {
                Write-Log -Message "An error has occurred while downloading!" -Type "WARNING"
                Write-Log -Message $_.Exception.Message -Type "WARNING"


                if (Test-Path -Path $Path -ErrorAction SilentlyContinue) {
                    Remove-Item $Path -Force -Confirm:$false -ErrorAction SilentlyContinue
                }


                $File = $False
            }


            # Restore the original progress preference setting
            $ProgressPreference = $PreviousProgressPreference
            # If the file was successfully downloaded, exit the loop
            if ($File) {
                $i = $Attempts
            }
            else {
                # Warn the user if the download attempt failed
                Write-Log -Message "File failed to download." -Type "WARNING"
                Write-Log -Message ""
            }


            # Increment the attempt counter
            $i++
        }


        # Final check: if the file still doesn't exist, report an error and exit
        if (!(Test-Path $Path)) {
            Write-Log -Message "Failed to download file." -Type "ERROR"
            Write-Log -Message "Please verify the URL of '$URL'." -Type "ERROR"
            exit 1
        }
        else {
            # If the download succeeded, return the path to the downloaded file
            return $Path
        }
    }


    # Check's the two Uninstall registry keys to see if the app is installed. Needs the name as it would appear in Control Panel.
    function Find-UninstallKey {
        [CmdletBinding()]
        param (
            [Parameter(ValueFromPipeline)]
            [String]$DisplayName,
            [Parameter()]
            [Switch]$UninstallString
        )
        process {
            $UninstallList = New-Object System.Collections.Generic.List[Object]


            $Result = Get-ChildItem HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Get-ItemProperty | 
                Where-Object { $_.DisplayName -like "*$DisplayName*" }


            if ($Result) { $UninstallList.Add($Result) }


            $Result = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Get-ItemProperty | 
                Where-Object { $_.DisplayName -like "*$DisplayName*" }


            if ($Result) { $UninstallList.Add($Result) }


            # Programs don't always have an uninstall string listed here so to account for that I made this optional.
            if ($UninstallString) {
                # 64 Bit
                $UninstallList | Select-Object -ExpandProperty UninstallString -ErrorAction Ignore
            }
            else {
                $UninstallList
            }
        }
    }
}
process {
    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'


    if (-not (Test-IsElevated)) {
        Write-Log -Message "Access Denied. Please run with Administrator privileges." -Type "ERROR"
        exit 1
    }


    if (-not (Test-Path $OfficeInstallDownloadPath )) {
        New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
    }


    if (-not ($ConfigurationXMLFile)) {
        Set-XMLFile
    }
    else {
        Invoke-Download -URL $ConfigurationXMLFile -Path "$OfficeInstallDownloadPath\OfficeInstall.xml"
        try {
            [xml]::new().Load("$OfficeInstallDownloadPath\OfficeInstall.xml")
        }
        catch {
            Write-Log -Message "The XML file is not valid. Please check the file and try again." -Type "ERROR"
            exit 1
        }
    }


    $ConfigurationXMLPath = "$OfficeInstallDownloadPath\OfficeInstall.xml"
    $ODTInstallLink = Get-ODTURL


    #Download the Office Deployment Tool
    Write-Log -Message "Downloading the Office Deployment Tool..."
    Invoke-Download -URL $ODTInstallLink -Path "$OfficeInstallDownloadPath\ODTSetup.exe"


    #Run the Office Deployment Tool setup
    try {
        Write-Log -Message "Running the Office Deployment Tool..."
        Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait -NoNewWindow
    }
    catch {
        Write-Log -Message "Error running the Office Deployment Tool. The error is below:" -Type "WARNING"
        Write-Log -Message "$_" -Type "WARNING"
        exit 1
    }


    #Run the O365 install
    try {
        Write-Log -Message "Downloading and installing Microsoft 365"
        $Install = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfigurationXMLPath" -Wait -PassThru -NoNewWindow


        if ($Install.ExitCode -ne 0) {
            Write-Log -Message "Exit Code does not indicate success!" -Type "ERROR"
            exit 1
        }
    }
    Catch {
        Write-Log -Message "Error running the Office install. The error is below:" -Type "WARNING"
        Write-Log -Message "$_" -Type "WARNING"
    }


    $OfficeInstalled = Find-UninstallKey -DisplayName "Microsoft 365"


    if ($CleanUpInstallFiles) {
        Write-Log -Message "Cleaning up install files..."
        Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
    }


    if ($OfficeInstalled) {
        Write-Log -Message "$($OfficeInstalled.DisplayName) installed successfully!"
        if ($Restart) {
            Write-Log -Message "Restarting the computer in 60 seconds..."
            Start-Process shutdown.exe -ArgumentList "-r -t 60" -Wait -NoNewWindow
        }
        exit 0
    }
    else {
        Write-Log -Message "Microsoft 365 was not detected after the install ran!" -Type "ERROR"
        exit 1
    }
}
end {
    
    
    
}