#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Restart the PrinterLogic Service
.DESCRIPTION
    - Restart the PrinterLogic Service
#>


<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\PrinterLogic_Close.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Set the Service name
$Service = "PrinterInstallerLauncher"

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

# Restart the PrinterLogic service
Restart-Service -force -Name $Service
Clear-Installation

