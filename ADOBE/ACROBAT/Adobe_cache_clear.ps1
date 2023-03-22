#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.1.0.7
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Clear the cache for Microsoft Adobe
.DESCRIPTION
    - Checks if Adobe is running and closes if necessary
    - Checks if an OLD folder exists in the Local and Roaming locations
    - If OLD folder exists, clear out contents and create fresh folder
    - Move files into OLD folder and rename folders to end with .old
    - Ask if user wants to open Adobe back up

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Directories
$tempDir = $env:TEMP
$adobelocal = "$env:LOCALAPPDATA\Adobe\Acrobat"
$adobeDC = "$env:LOCALAPPDATA\Adobe\Acrobat\DC"
$adobeXI = "$env:LOCALAPPDATA\Adobe\Acrobat\XI"
$oldDirs = Get-ChildItem -Path $adobelocal -Directory -Filter *.old | Where-Object { Test-Path $_.FullName -PathType Container }

# Progress Bar value
$Variable = 0

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$appLogs = "$tempDir\Adobe_Cache.log"
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

# Function to prompt the user with a popup when install cancelled
function Pop-Cancelled {
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show($global:output, 'Cancelled', 'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Function to prompt the user with a popup when install succeeds
function Pop-Success {
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show($global:output, 'Completed', 'OK', 'Information')
}

#Function to automatically close Adobe Acrobat
function Close-AdobeAcrobat {
  
  if (Get-Process -Name "Acrobat" -ErrorAction SilentlyContinue) {
    # Close out of Adobe
    Stop-Process -name Acrobat -force

    # Set the duration of the timer in seconds
    $duration = 10

    # Initialize the progress bar
    Write-Progress -Activity "Waiting for $duration seconds while Adobe closes..." -PercentComplete 0

    # Loop through the timer and update the progress bar
    for ($i = 1; $i -le $duration; $i++) {
      # Update the progress bar with the current progress
      $percent = ($i / $duration) * 100
      Write-Progress -Activity "Waiting for $duration seconds while Adobe closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
    
      # Pause for 1 second
      Start-Sleep -Seconds 1
    }

    # Clear the progress bar once the timer is complete
    Write-Progress -Completed -Activity "Microsoft Adobe is Closed!"
  }
  else {

  }
}

#Function to automatically clear the cache of Adobe Acrobat
function Reset-AdobeAcrobat {

  $global:output = ''
# Create a Windows form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Adobe Cache Clear"
$Form.Size = New-Object System.Drawing.Size(300,200)
$Form.StartPosition = "CenterScreen"
$form.TopMost = $true

# Create the "start" button
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Point(100, 90)
$StartButton.Size = New-Object System.Drawing.Size(75, 23)
$StartButton.Text = "Start"
$StartButton.Enabled = $true
$StartButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $StartButton
$Form.Controls.Add($StartButton)

# Create a progress bar and set its properties
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(50,50)
$ProgressBar.Size = New-Object System.Drawing.Size(200,20)
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = 100

# Add the progress bar to the form
$Form.Controls.Add($ProgressBar)

# Set the value of the progress bar based on a variable
$ProgressBar.Value = $Variable

# Show the form
$Result = $Form.ShowDialog()


if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
if ((Test-Path $adobeDC) -and (Test-Path $adobeXI)) {
  $Variable = 25
    # Both paths exist
    $adobeTest = "Both Adobe DC and Adobe XI are installed"
}
elseif (Test-Path $adobeDC) {
  $Variable = 25
    # Only Adobe DC path exists
    $adobeTest = "Adobe DC is installed"
}
elseif (Test-Path $adobeXI) {
  $Variable = 25
    # Only Adobe XI path exists
    $adobeTest = "Adobe XI is installed"
}
else {
  $Variable = 25
    # Neither path exists
    $adobeTest = "Adobe is not installed"
}

# Perform action based on $adobeTest
if ($adobeTest -eq "Both Adobe DC and Adobe XI are installed") {
  $Variable = 50
  
if ($oldDirs) {
  $Variable = 75
    foreach ($dir in $oldDirs) {
      Remove-Item -Path "$adobelocal\*.old" -Recurse -ErrorAction SilentlyContinue
    }
    $Variable = 80
    Rename-Item -Path "$adobelocal\XI" "$adobelocal\XI.old" -ErrorAction SilentlyContinue
    $Variable = 90
    Rename-Item -Path "$adobelocal\DC" "$adobelocal\DC.old" -ErrorAction SilentlyContinue
    $global:output = "Cache is cleared!"
    $Variable = 100
} else {
    $Variable = 75
    Rename-Item -Path "$adobelocal\XI" "$adobelocal\XI.old" -ErrorAction SilentlyContinue
    $Variable = 90
    Rename-Item -Path "$adobelocal\DC" "$adobelocal\DC.old" -ErrorAction SilentlyContinue
    $global:output = "Cache is cleared!"
    $Variable = 100
}
}
elseif ($adobeTest -eq "Adobe DC is installed") {
  $Variable = 50
  
  if ($oldDirs) {
    $Variable = 75
      foreach ($dir in $oldDirs) {
        Remove-Item -Path "$adobelocal\*.old" -Recurse -ErrorAction SilentlyContinue
      }
      $Variable = 80  
  Rename-Item -Path "$adobelocal\DC" "$adobelocal\DC.old" -ErrorAction SilentlyContinue
  $global:output = "Cache is cleared!"
  $Variable = 100
} else {
  $Variable = 75 
  Rename-Item -Path "$adobelocal\DC" "$adobelocal\DC.old" -ErrorAction SilentlyContinue
  $global:output = "Cache is cleared!"
  $Variable = 100
}
}
elseif ($adobeTest -eq "Adobe XI is installed") {
  $Variable = 50
  
  if ($oldDirs) {
    $Variable = 75
      foreach ($dir in $oldDirs) {
        Remove-Item -Path "$adobelocal\*.old" -Recurse -ErrorAction SilentlyContinue
      }
      $Variable = 80  
  Rename-Item -Path "$adobelocal\XI" "$adobelocal\XI.old" -ErrorAction SilentlyContinue
  $global:output = "Cache is cleared!"
  $Variable = 100
} else {
  $Variable = 75 
  Rename-Item -Path "$adobelocal\XI" "$adobelocal\XI.old" -ErrorAction SilentlyContinue
  $global:output = "Cache is cleared!"
  $Variable = 100
}
}
else {
  $Variable = 100
    # Nothing to do but quit
    $global:output = $adobeTest
}
}
if ($Variable -eq 100) {
  $form.Close()
  Pop-Success
  Clear-Installation
  exit
}
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Close-AdobeAcrobat
Reset-AdobeAcrobat




