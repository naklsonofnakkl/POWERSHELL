#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Moves bar in Outlook to bottom instead of sideways
.DESCRIPTION
    - Alters the registry path
#>


<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\Outlook_Registry.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Registry Path
$regPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentEcs\Overrides'
# Registry Value
$property = (Get-ItemProperty -Path $regPath).Microsoft.Office.Outlook.Hub.HubBar

$global:output = ''

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

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

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

# Check the value of the registry
if ($property -eq "False") {
    # Registry Value is already properly set, close script!
    $global:output = "The registry value is already set to False!"
    Pop-Success
    Clear-Installation
}
else {
    # Set the Value of the Registry to False
    $global:output = "Changing the registry value to False!"
    Push-Location
    Set-Location $regPath
    Set-ItemProperty . Microsoft.Office.Outlook.Hub.HubBar "False"
    # Validate if the registry was set properly
    if ($property -eq "False") {
        # Registry was correctly set, close script!
        $global:output = "The registry value is already set to False!"
        Pop-Location
        Pop-Success
        Clear-Installation
    }
    else {
        # Registry Value failed to change, close script!
        $global:output = "Failed to Change Registry Value!"
        Pop-Location
        Pop-Cancelled
        Clear-Installation
    }
}



