#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.1
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Checks for open LogMeIn Rescue windows and closes them / removes local cache
.DESCRIPTION
    - Search available services for LMIRescue
    - Copy the service name and description to a variable
    - Delete the service associated with LMI
    - Remove the LogMeIn Rescue Applet cache

#>


$services = Get-Service | Where-Object { $_.DisplayName -like "LMIRescue_*" } | Select-Object -Property Name, DisplayName

$serviceHashTable = @{}

foreach ($service in $services) {
    $serviceHashTable.Add($service.DisplayName, $service.Name)
}

foreach ($serviceName in $serviceHashTable.Values) {
    Write-Output "Deleting service: $serviceName"
    sc.exe delete $serviceName
}

Remove-Item -Path "$env:LOCALAPPDATA\LogMeIn Rescue Applet" -Recurse -Force

