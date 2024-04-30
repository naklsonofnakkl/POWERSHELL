#Requires -RunAsAdministrator
<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.1.3
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Automatically disable and re-enable Intel(R) USB 3 ports on Lenovo 

.DESCRIPTION
    - finds any devices on the machine matching the wildcard
    - Disables the found devices
    - Re-enables the found devices
#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>
#Directories
$tempDir = New-Item -ItemType Directory -Path "$env:TEMP" -Name Custom_Scripts -Force

## LOGS
## C:\Users\[USERNAME]\AppData\Local\Temp\Custom_Scripts
$appLogs = "$tempDir\Disable_USB.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>
function Clear-Installation {
    Stop-Transcript
    exit
}

function Group-USBDevices {
## wildcard to grab the generally affected USB drivers
$deviceNamePattern = "Intel(R) USB 3.10*"
## Get the USB Controller devices matching the wildcard
$usbDevices = Get-PnpDevice | Where-Object { $_.Class -eq "USB" -and $_.FriendlyName -like $deviceNamePattern }

## If no devices are found, exit gracefully
if ($usbDevices.Count -eq 0) {
    Write-Host "No matching USB devices found."
    exit
}
}

function Reset-USB {
## Attempt to disable each device and handle errors
foreach ($device in $usbDevices) {
    $deviceId = $device.InstanceId
    $friendlyName = $device.FriendlyName
    
    try {
        Disable-PnpDevice -InstanceId $deviceId -Confirm:$false -ErrorAction Stop
        Write-Host "Device with Instance ID $deviceId disabled successfully."
    } catch {
        Write-Host "Failed to disable device $friendlyName with Instance ID $deviceId."
        Write-Host "Error: $_"
    }
}

## Wait for 5 seconds
Start-Sleep -Seconds 5

## Re-enable the matching devices
foreach ($device in $usbDevices) {
    Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false
}
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>
Group-USBDevices
$device = $usbDevices.InstanceId
Reset-USB
Clear-Installation