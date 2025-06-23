# ===================================================================================
# PowerShell Script to Silently Uninstall Microsoft Office 2016 MSI (ProPlus)
#
# Description:
# This script automates the uninstallation of Office 2016 Professional Plus.
# It performs the following actions:
# 1. Adds a 5-minute delay.
# 2. Checks for Administrator privileges.
# 3. Creates the necessary uninstall XML configuration file in a temporary location.
# 4. Automatically detects the correct path to the Office uninstaller.
# 5. Executes the silent uninstallation command.
# 6. Cleans up the temporary XML file after it's done.
# ===================================================================================

# Step 1: Wait for 5 minutes (300 seconds) before starting
Write-Host "Script initiated. Waiting for 5 minutes before proceeding..."
Write-Host "This delay can be used to ensure services have started or to allow time for remote execution setup."
Start-Sleep -Seconds 300

# Step 2: Check if the script is running as an Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "This script must be run as an Administrator."
    Write-Warning "Please re-launch PowerShell using 'Run as administrator' and try again."
    # Pause to allow the user to read the message before the window closes.
    if ($env:PSNativeCommandUseErrorActionPreference -ne $true) {
        Read-Host "Press Enter to exit"
    }
    exit
}

# Step 3: Define the XML content for silent uninstallation
# This is a PowerShell "here-string", which allows for multi-line text.
$xmlContent = @"
<Configuration Product="ProPlus">
    <Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
</Configuration>
"@

# Step 4: Create a temporary file to store the XML configuration
# This is safer than hardcoding a path like the Desktop.
$tempXmlPath = Join-Path $env:TEMP "uninstall_config.xml"
try {
    Set-Content -Path $tempXmlPath -Value $xmlContent
    Write-Host "Successfully created temporary uninstall config at: $tempXmlPath"
}
catch {
    Write-Error "Failed to create the temporary XML file. Error: $_"
    exit
}


# Step 5: Find the correct path for the Office uninstaller (setup.exe)
$officeSetupPath = ""
$path64bit = "C:\Program Files\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe"
$path32bit = "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.exe"

if (Test-Path $path64bit) {
    $officeSetupPath = $path64bit
}
elseif (Test-Path $path32bit) {
    $officeSetupPath = $path32bit
}

# Step 6: Execute the uninstallation if setup.exe was found
if ($officeSetupPath) {
    Write-Host "Found Office uninstaller at: $officeSetupPath"
    Write-Host "Starting silent uninstallation... This may take several minutes."
    
    try {
        # Using Start-Process is the most robust way to run external commands.
        # It handles paths with spaces and allows for running with elevated privileges.
        Start-Process -FilePath $officeSetupPath -ArgumentList "/uninstall PROPLUS /config `"$tempXmlPath`"" -Wait -Verb RunAs
        
        Write-Host "Uninstallation process completed."
        Remove-NinjaTag -Name "Office 2016 Installed"
    }
    catch {
        Write-Error "An error occurred while running the uninstaller. Error: $_"
    }
    finally {
        # Step 7: Clean up by removing the temporary XML file
        if (Test-Path $tempXmlPath) {
            Remove-Item $tempXmlPath -Force
            Write-Host "Temporary config file has been removed."
        }
    }
}
else {
    # This block runs if setup.exe was not found in either standard location.
    Write-Error "Could not find the Office uninstaller (setup.exe) in standard locations."
    Write-Warning "The Office installation may be damaged. Consider using the Microsoft Support and Recovery Assistant."
    
    # Clean up the created XML file even on failure
    if (Test-Path $tempXmlPath) {
        Remove-Item $tempXmlPath -Force
        Write-Host "Temporary config file has been removed."
    }
}

Write-Host "Script finished."
