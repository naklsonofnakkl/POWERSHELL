<#
.NOTES
    Author: Andrew Wilson
    Version: 1.2.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Clear the cache for Microsoft Outlook
.DESCRIPTION
    - Checks if Outlook is running and closes if necessary
    - Checks if an OLD folder exists in the Local AppData location
    - If OLD folder exists, clear out contents and create fresh folder
    - Move files into OLD folder and rename folders to end with .old

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Directories
$tempDir = $env:TEMP
$outlook = "$env:LOCALAPPDATA\Microsoft\Outlook"
$currentDate = Get-Date -Format "MM-dd-yyyy"
$oldCache = "$env:LOCALAPPDATA\Microsoft\Outlook\OLD-$currentDate"
$folders = Get-ChildItem -Path $outlook -Directory
$officeApps = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTEM", "ONENOTE", "ms-teams", "CiscoCollabHost", "wbxcOIEx64")

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$appLogs = "$tempDir\MSOutlook_Cache.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

# Function to clean up the leftover downloaded files
function Clear-Cleanup {
  Stop-Transcript
}

function Clear-OldCache { 
# This will delete any previous OLD folders and contents then create a new one and start the move process
try {
    Get-ChildItem -Path $outlook -Directory |
    Where-Object {$_.Name -like 'OLD*'} |
    Remove-Item -Recurse -Force
    New-Item -path $outlook -name "OLD-$currentDate" -ItemType Directory
    Move-Item -Path $outlook\*.srs $oldCache
    Move-Item -Path $outlook\*.xml $oldCache
    Move-Item -Path $outlook\*.nst $oldCache
    Move-Item -Path $outlook\*.ost $oldCache
    Move-Item -Path $outlook\*.old $oldCache
  }
catch {
    Write-Host "Failed to clear Outlook Cache..."
}
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

# Close each app listed in the $officeapps array to prevent issues with renaming folders and files
foreach ($app in $officeApps) {
    try {
        Stop-Process -Name $app -Force -ErrorAction SilentlyContinue
        Write-Host "Application Closed: $app"
    } catch {
        Write-Host "Failed to close application: $app"
    }
}

# Add .old suffix to all folders in Outlook folder, excluding any folders starting with the name "OLD"
foreach ($folder in $folders) {
    if ($folder.Name -notlike 'OLD*') {
        $newName = "$($folder.Name).old"
        Rename-Item -Path $folder.FullName -NewName $newName
    }
}

Clear-OldCache
Clear-Cleanup
