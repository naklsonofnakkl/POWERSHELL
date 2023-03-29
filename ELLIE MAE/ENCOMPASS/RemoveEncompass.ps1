#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    Based on BAT script by: Sebastien Dolce (2017/05/18)
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Completely remove Ellie Mae Encompass
.DESCRIPTION
    - Remove several folders linked to Encompass
    - Validate that folders are removed
    - Remove registry values linked to Encompass
    - Validate registry values are removed
    - Clean up and close script
#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\Encompass_Uninstall.log"
$ErrorActionPreference = "Continue"
Start-Transcript -Path $appLogs -Append

$encompassProducts = "Encompass eFolder", "Encompass SmartClient", "Encompass Document Converter", "SmartClient Core"

$folders = @("C:\SmartClientCache", "C:\Program Files (x86)\Ellie Mae", "C:\Encompass", "C:\EncompassData", "$env:LOCALAPPDATA\Encompass Installation", "$env:LOCALAPPDATA\Temp\Encompass", "$env:LOCALAPPDATA\EncompassSC", "$env:LOCALAPPDATA\Low\Apps\Ellie Mae", "$env:APPDATA\EllieMae", "$env:APPDATA\Encompass", "C:\Windows\System32\config\systemprofile\AppData\Local\Encompass")

$global:output = ''

$regKeys = @(
    "HKLM:\SOFTWARE\Ellie Mae",
    "HKLM:\SOFTWARE\Wow6432Node\Ellie Mae",
    "HKLM:\SOFTWARE\Wow6432Node\Black Ice Software LLC\Encompass Document Converter",
    "HKCU:\SOFTWARE\Ellie Mae", "HKCU:\SOFTWARE\Black Ice Software LLC\Encompass Document Converter", "HKCU:\SOFTWARE\Encompass", "HKCC:\SOFTWARE\Encompass"
)

$found = $false

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

# Uninstall Encompass, Document Converter, and eFolder
function Clear-Encompass {
    foreach ($product in $encompassProducts) {
        Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $product } | ForEach-Object { $_.Uninstall() }
    }

    & "C:\Windows\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -d -p "Encompass" | Out-Null
    & "C:\Windows\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -d -p "Encompass Document Converter" | Out-Null
    & "C:\Windows\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -d -p "Encompass eFolder" | Out-Null
}

# Function to clean up the leftover downloaded files
function Clear-Installation {
    Stop-Transcript
    exit
}

# Function to prompt the user with a popup when install cancelled
function Pop-Cancelled {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($global:output, 'Error', 'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Function to prompt the user with a popup when install succeeds
function Pop-Success {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($global:output, 'Install Completed', 'OK', 'Information')
}

# Function to remove multiple Encompass folders
function Remove-EncompassFolders {
    Remove-Item -Path "C:\SmartClientCache" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Program Files (x86)\Ellie Mae" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Encompass" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\EncompassData" -Recurse -Force -ErrorAction SilentlyContinue
    
    Remove-Item -Path "$env:LOCALAPPDATA\Encompass Installation" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:LOCALAPPDATA\Temp\Encompass" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:LOCALAPPDATA\EncompassSC" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:LOCALAPPDATA\Low\Apps\Ellie Mae" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:APPDATA\EllieMae" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:APPDATA\Encompass" -Recurse -Force -ErrorAction SilentlyContinue

    Remove-Item -Path "C:\Windows\System32\config\systemprofile\AppData\Local\Encompass" -Recurse -Force -ErrorAction SilentlyContinue
}

# Function to remove Encompass Registry Files
function Remove-EncompassRegistry {
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Ellie Mae" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Ellie Mae" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Black Ice Software LLC\Encompass Document Converter" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Ellie Mae" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Black Ice Software LLC\Encompass Document Converter" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Encompass" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKCC:\SOFTWARE\Encompass" -Recurse -Force -ErrorAction SilentlyContinue
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Clear-Encompass
Remove-EncompassFolders
$results = foreach ($folder in $folders) {
    Test-Path $folder
}

if ($results -contains $true) {
    $global:output = "At least one folder still exists. Closing script!"
    Pop-Cancelled
    Clear-Installation
}
else {
    $global:output = "All Folders Removed! Press OK to continue."
    Pop-Success
    Remove-EncompassRegistry
    foreach ($key in $regKeys) {
        if (Test-Path $key) {
            $global:output = "Registry key $key still exists. Closing script!"
            $found = $true
            Pop-Cancelled
            Clear-Installation
        }
    }

    if (!$found) {
        $global:output = "None of the registry keys exist!"
        Pop-Success
        $global:output = "Encompass has been removed! Closing script!"
        Pop-Success
        Clear-Installation
    }
}