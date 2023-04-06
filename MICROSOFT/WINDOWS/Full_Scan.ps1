#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.1
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Checks to see if there are any corrupted files/drives
.DESCRIPTION
    - Using SFC (System File Checker) to scan for corrupted Windows files
    - Using CHKDSK to scan for HDD errors
    - Using DISM (Deployment Image Servicing and Management) to repair a Windows image
    - 

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\FullScan_PS.log"
$ErrorActionPreference = "continue"
Start-Transcript -Path $appLogs -Append

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

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

sfc /scannow
chkdsk C: /f /r /x
#Dism /Online /Cleanup-Image /RestoreHealth

# Ask if user needs to Restart
Add-Type -AssemblyName System.Windows.Forms
$option = [System.Windows.Forms.MessageBox]::Show('Do you want to restart your computer?', 'Restart Computer', 'YesNo', 'Question')

if ($option -eq 'Yes') {
    Write-Host 'Sweet dreams my friend...'
    Stop-Transcript
    Restart-Computer
} else {
    Write-Host 'Computer will not be restarted!'
    Clear-Installation
}
