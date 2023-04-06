#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.1
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Gathers all Display Drivers and disables them before re-enabling them
.DESCRIPTION
    - Grabs all the Display devices
    - Disables them and validates they are disabled
    - wait for 30 seconds
    - re-enable all Display devices
    - validate they have all been enabled
#>


<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\Disable_Display.log"
$ErrorActionPreference = "Stop"
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
# Get all enabled display devices
$enabledDevices = Get-PnpDevice -Class "Display" | Where-Object {$_.Status -eq "OK"}

# Disable all enabled display devices
$enabledDevices | ForEach-Object {
    Disable-PnpDevice -InstanceId $_.InstanceId -Confirm:$false
}

# Wait for 30 seconds
Start-Sleep -Seconds 30

# Get all disabled display devices
$disabledDevices = Get-PnpDevice -Class "Display" | Where-Object {$_.Status -ne "OK"}

# Enable all disabled display devices
$disabledDevices | ForEach-Object {
    Enable-PnpDevice -InstanceId $_.InstanceId -Confirm:$false
}
Clear-Installation