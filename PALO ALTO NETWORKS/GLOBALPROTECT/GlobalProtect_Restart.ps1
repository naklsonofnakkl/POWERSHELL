#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Restart the Global Protect Application and Service
.DESCRIPTION
    - Closes the Global Protect application
    - Restarts the Global Protect Service
#>


<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\PanGPA_Cache.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Set the Process and Service name
$Process = "PanGPA"
$Service = "PanGPS"

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

# Close the application
Stop-Process -Name $Process
# Restart the PanGPS service
Restart-Service -Name $Service
Clear-Installation