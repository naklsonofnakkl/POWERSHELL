#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.1.0.5
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Clear the cache for Microsoft Outlook
.DESCRIPTION
    - Checks if Outlook is running and closes if necessary
    - Checks if an OLD folder exists in the Local and Roaming locations
    - If OLD folder exists, clear out contents and create fresh folder
    - Move files into OLD folder and rename folders to end with .old
    - Ask if user wants to open Outlook back up

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Directories
$tempDir = $env:TEMP
$outlookDir = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$outroam = "$env:APPDATA\Microsoft\Outlook"
$outlocal = "$env:LOCALAPPDATA\Microsoft\Outlook"
$oldroam = "$env:APPDATA\Microsoft\Outlook\OLD"
$oldoffline = "$outlocal\Offline Address Books"

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
function Clear-Installation {
  Stop-Transcript
}

#Function to automatically close Microsoft Outlook
function Close-MicrosoftOutlook {
  
  if (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue) {
    # Close out of Outlook
    Stop-Process -name OUTLOOK -force

    # Set the duration of the timer in seconds
    $duration = 10

    # Initialize the progress bar
    Write-Progress -Activity "Waiting for $duration seconds while Outlook closes..." -PercentComplete 0

    # Loop through the timer and update the progress bar
    for ($i = 1; $i -le $duration; $i++) {
      # Update the progress bar with the current progress
      $percent = ($i / $duration) * 100
      Write-Progress -Activity "Waiting for $duration seconds while Outlook closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
    
      # Pause for 1 second
      Start-Sleep -Seconds 1
    }

    # Clear the progress bar once the timer is complete
    Write-Progress -Completed -Activity "Microsoft Outlook is Closed!"
  }
  else {

  }
}

#Function to automatically open Microsoft Outlook
function Open-MicrosoftOutlook {
  Add-Type -AssemblyName System.Windows.Forms
  $caption = "Outlook Restart Prompt"
  $message = "Do you wish to re-open Microsoft Outlook?"
  $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
  $result = [System.Windows.Forms.MessageBox]::Show($message, $caption, $buttons)

  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    # User clicked "Yes"
    # Launch Microsoft Outlook with the target command
    Start-Process -FilePath $outlookDir
    Clear-Installation
    exit
  }
  else {
    # User clicked "No"
    Clear-Installation
    exit
  }
}

#Function to automatically clear the cache of Microsoft Outlook
function Reset-MicrosoftOutlook {
  #If there is no OLD folder create one and copy files into it
  #ROAMING
  if ( -not ( Test-Path -Path $oldroam ) ) {
    New-Item -path $outroam -name OLD -ItemType Directory
    Move-Item -Path $outroam\*.srs $outroam\OLD
    Move-Item -Path $outroam\*.xml $outroam\OLD
  }
  #If there is an OLD folder erase the OLD folder and create fresh Copy
  #ROAMING
  else {
    Remove-Item -Path "$outroam\OLD" -Recurse -Force
    New-Item -path $outroam -name OLD -ItemType Directory
    Move-Item -Path $outroam\*.srs $outroam\OLD
    Move-Item -Path $outroam\*.xml $outroam\OLD
  }
  #If there are no .old folders then rename all folders to end in .old
  #LOCAL
  if ( -not ( Test-Path -Path "$outlocal\RoamCache.old" ) ) {
    Rename-Item -Path "$outlocal\RoamCache" "$outlocal\RoamCache.old"
    # Check if the Offline Address Books directory is writable
    $accessControl = (Get-Acl $oldoffline).Access
    $writeAccess = $accessControl | Where-Object { $_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Write }
    if (!$writeAccess) {
      # The folder is not writable, so close the Excel process
      Get-Process -Name "Excel" | ForEach-Object { $_.CloseMainWindow() }
    }
    Rename-Item -Path "$outlocal\Offline Address Books" "$outlocal\Offline Address Books.old"
  }
  #If there are .old folders, delete them and convert current foldres into .old
  #LOCAL
  else {
    Remove-Item -Path $outlocal\*.old -Recurse
    Rename-Item -Path "$outlocal\RoamCache" "$outlocal\RoamCache.old"
    # Check if the Offline Address Books directory is writable
    $accessControl = (Get-Acl $oldoffline).Access
    $writeAccess = $accessControl | Where-Object { $_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Write }
    if (!$writeAccess) {
      # The folder is not writable, so close the Excel process
      Get-Process -Name "Excel" | ForEach-Object { $_.CloseMainWindow() }
    }
    Rename-Item -Path "$outlocal\Offline Address Books" "$outlocal\Offline Address Books.old"
  }
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Close-MicrosoftOutlook
Reset-MicrosoftOutlook
Open-MicrosoftOutlook