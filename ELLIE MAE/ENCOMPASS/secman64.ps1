#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.1
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Script to install the Secman DLL files necessary for Encompass
.DESCRIPTION
    - Closes Ellie Mae Encompass
    - Installs the Secman DLL's
    - Validates they have been installed
    - Cleans itself up

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Location the DLL will be installed to
$secmanPath = "C:\Program Files (x86)\Common Files\Outlook Security Manager\"
# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\MSEncompass_Cache.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

#create a global output for popups
$global:output = ''


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

# Function to prompt the user with a popup when install cancelled
function Pop-Cancelled {
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show($global:output, 'Error', 'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
}

# Function to prompt the user with a popup when install succeeds
function Pop-Success {
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show($global:output, 'Install Completed', 'OK', 'Information')
}

#Function to automatically close EllieMae Encompass
function Close-EllieMaeEncompass {
  
  if (Get-Process -Name "encompass" -ErrorAction SilentlyContinue) {
    # Close out of Encompass
    Stop-Process -name encompass -force

    # Set the duration of the timer in seconds
    $duration = 10

    # Initialize the progress bar
    Write-Progress -Activity "Waiting for $duration seconds while Encompass closes..." -PercentComplete 0

    # Loop through the timer and update the progress bar
    for ($i = 1; $i -le $duration; $i++) {
      # Update the progress bar with the current progress
      $percent = ($i / $duration) * 100
      Write-Progress -Activity "Waiting for $duration seconds while Encompass closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
    
      # Pause for 1 second
      Start-Sleep -Seconds 1
    }

    # Clear the progress bar once the timer is complete
    Write-Progress -Completed -Activity "EllieMae Encompass is Closed!"
  }
  else {

  }
}

function Get-SecmanDLL {
  # Register secman.dll and secman64.dll
  $regsvr32 = "regsvr32.exe"
  $regsvr32Args = "/s ""C:\Program Files (x86)\Common Files\Outlook Security Manager\secman.dll"""
  Start-Process -FilePath $regsvr32 -ArgumentList $regsvr32Args -Wait

  $regsvr32Args = "/s ""C:\Program Files (x86)\Common Files\Outlook Security Manager\secman64.dll"""
  Start-Process -FilePath $regsvr32 -ArgumentList $regsvr32Args -Wait

  # Check if secman.dll exists
  if (Test-Path ($secmanPath + "secman.dll")) {
    $global:output = "secman.dll is installed!"
    Pop-Success
  }
  else {
    $global:output = "secman.dll is missing!"
    Pop-Cancelled
  }

  # Check if secman64.dll exists
  if (Test-Path ($secmanPath + "secman64.dll")) {
    $global:output = "secman64.dll is installed!"
    Pop-Success
  }
  else {
    $global:output = "secman64.dll is missing!"
    Pop-Cancelled
  }
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Close-EllieMaeEncompass
Get-SecmanDLL
Clear-Installation