<#
.SYNOPSIS
    Reliably checks for the installation of Microsoft Office 2016 by querying the Windows Uninstall registry keys.

.DESCRIPTION
    This script determines if any edition of Microsoft Office 2016 is installed on the local machine.
    Instead of checking for potentially leftover application keys, this script queries the
    official registry locations that Windows uses to populate the "Add or Remove Programs" list.
    This provides a much more accurate result.

    It checks both 32-bit and 64-bit installation locations and looks for any product
    with a display name matching "Microsoft Office ... 2016".

.NOTES
    - Version: 2.0
    - Author: Gemini
    - Run this script with PowerShell.
    - This method is more reliable than checking for specific application install paths.
#>

# --- Function to check for installed software ---
function Test-Office2016Installation {
    # Turn off error messages for registry queries where keys might not exist
    $ErrorActionPreference = 'SilentlyContinue'

    # --- Registry Paths for Installed Programs (32-bit and 64-bit locations) ---
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    Write-Host "Searching for Microsoft Office 2016 in the list of installed programs..."

    # Use a wildcard search to find any edition of Office 2016
    $searchPattern = "Microsoft Office*2016*"

    # Get properties from all uninstall keys and filter them
    # We select the DisplayName and check if it's not null before matching
    $installedProgram = Get-ItemProperty -Path $uninstallPaths | 
                        Where-Object { $_.DisplayName -ne $null } | 
                        Where-Object { $_.DisplayName -like $searchPattern } |
                        Select-Object -First 1 # Stop after the first match

    # --- Output the Result ---
    if ($installedProgram) {
        # If a matching program was found
        Write-Host "Result: Microsoft Office 2016 appears to be INSTALLED." -ForegroundColor Green
        Write-Host "Detected Version: $($installedProgram.DisplayName)"
        Set-NinjaTag = -Name "Office 2016 Installed"
    }
    else {
        # If no match was found after checking all programs
        Write-Host "Result: Microsoft Office 2016 is NOT INSTALLED." -ForegroundColor Yellow
    }

    # Reset the error action preference to its default value
    $ErrorActionPreference = 'Continue'
}

# --- Run the function ---
Test-Office2016Installation