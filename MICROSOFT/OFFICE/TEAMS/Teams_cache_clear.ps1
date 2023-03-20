#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.1.0.1
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Clear the cache for Microsoft Teams
.DESCRIPTION
    - Checks if Teams is running and closes if necessary
    - Checks if an OLD folder exists
    - If OLD folder exists, clear out contents and create fresh folder
    - If no OLD folder exists, create one
    - Move all files and folders except for OLD and Meeting-Addin to OLD folder
    - Ask if user wants to open Teams back up

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Directories
$tempDir = $env:TEMP
$teamroam = "$env:APPDATA\Microsoft\Teams"
$oldroam = "$env:APPDATA\Microsoft\Teams\OLD"

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$appLogs = "$tempDir\MSTEAMS_Cache.log"
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

#Function to automatically close Microsoft Teams
function Close-MicrosoftTeams {
  
  if (Get-Process -Name "Teams" -ErrorAction SilentlyContinue) {
    # Close out of TEAMS
    Stop-Process -name Teams -force

    # Set the duration of the timer in seconds
    $duration = 10

    # Initialize the progress bar
    Write-Progress -Activity "Waiting for $duration seconds while Teams closes..." -PercentComplete 0

    # Loop through the timer and update the progress bar
    for ($i = 1; $i -le $duration; $i++) {
      # Update the progress bar with the current progress
      $percent = ($i / $duration) * 100
      Write-Progress -Activity "Waiting for $duration seconds while Teams closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
    
      # Pause for 1 second
      Start-Sleep -Seconds 1
    }

    # Clear the progress bar once the timer is complete
    Write-Progress -Completed -Activity "Microsoft Teams is Closed!"
  }
  else {

  }
}

#Function to automatically open Microsoft Teams
function Open-MicrosoftTeams {
  Add-Type -AssemblyName System.Windows.Forms
  $caption = "Teams Restart Prompt"
  $message = "Do you wish to re-open Microsoft Teams?"
  $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
  $result = [System.Windows.Forms.MessageBox]::Show($message, $caption, $buttons)

  if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    # User clicked "Yes"
    # Launch Microsoft Teams with the target command
    Start-Process "C:\Users\$env:USERNAME\AppData\Local\Microsoft\Teams\Update.exe" `
      -ArgumentList "--processStart", "Teams.exe", "--process-start-args", "--profile=AAD"
    Clear-Installation
    exit
  }
  else {
    # User clicked "No"
    Clear-Installation
    exit
  }
}

#Function to automatically clear the Microsoft Teams cache
function Reset-MicrosoftTeams {
  #If there is no OLD folder create one and copy files into it
  #ROAMING
  if ( -not ( Test-Path -Path $oldroam ) ) {
    New-Item -path $teamroam -name OLD -ItemType Directory
    set-location $teamroam
    $filedest = $oldroam
    $exclude = ".\meeting-addin", ".\OLD"
    $Files = Get-ChildItem -path $teamroam | Where-object { $_.name -ne $exclude }
    foreach ($file in $files) { move-item -path $file -destination $filedest -ErrorAction SilentlyContinue }
  }
  #If there is an OLD folder erase the OLD folder and create fresh Copy
  #ROAMING
  else {
    Remove-Item -Path "$teamroam\OLD" -Recurse -Force
    New-Item -path $teamroam -name OLD -ItemType Directory
    set-location $teamroam
    $filedest = $oldroam
    $exclude = ".\meeting-addin", ".\OLD"
    $Files = Get-ChildItem -path $teamroam | Where-object { $_.name -ne $exclude }
    foreach ($file in $files) { move-item -path $file -destination $filedest -ErrorAction SilentlyContinue }
  }
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Close-MicrosoftTeams
Reset-MicrosoftTeams
Open-MicrosoftTeams