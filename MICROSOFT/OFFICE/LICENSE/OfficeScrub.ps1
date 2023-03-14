<#
.SYNOPSIS
    scrub the licenses for Office 2016+ applications
.DESCRIPTION
    This script will scrub the licenses for Office 2016+ applications using a Microsoft-Made VBS script
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL
#>

# Kill explorer.exe forcefully
Stop-Process -ProcessName explorer -Force

# Stop CryptSvc service
Stop-Service CryptSvc

# validate the existance of the temp folder and toss the BrokerPlugin into it
$source = "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy"
$destination = "C:\temp"

if (Test-Path $destination) {
    Move-Item $source $destination -Force
} else {
    New-Item -ItemType Directory -Force -Path $destination
    Move-Item $source $destination -Force
}

# Start CryptSvc service
Start-Service CryptSvc

# Start explorer.exe
Start-Process explorer

# Download and Execute the OfficeLicenseScrub.vbs Script
$url = "https://raw.githubusercontent.com/username/repo/main/OfficeLicenseScrub.vbs"
$outputPath = "C:\temp\OfficeLicenseScrub.vbs"

Invoke-WebRequest -Uri $url -OutFile $outputPath
Start-Process -FilePath "C:\Windows\System32\cscript.exe" -ArgumentList $outputPath -Wait

# Log the user off
logoff
