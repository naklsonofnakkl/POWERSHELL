#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.5
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    The Big Beautiful Cache Clearing Module!
.DESCRIPTION
    - Clears the Cache of multiple applications
    - All packed into one easy to call module
    - And at the total mercy of your imagination!
#>

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

function Clear-Outlook {
  <#
--------------------
 VARIBALE PARADISE!
--------------------
#>

  #Directories
  $tempDir = $env:TEMP
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
    $caption = "Outlook Cache Cleared"
    $message = "The Outlook Cache has been cleared!"
    $buttons = [System.Windows.Forms.MessageBoxButtons]::Ok
    $result = [System.Windows.Forms.MessageBox]::Show($message, $caption, $buttons)
    Clear-Installation
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
}

function Clear-Teams {
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
}

function Clear-AdobeAcrobat {
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

  $global:output = ''

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

    # Create a Windows form
    Add-Type -AssemblyName System.Windows.Forms
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Adobe Cache Clear"
    $Form.Size = New-Object System.Drawing.Size(300, 200)
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
    $ProgressBar.Location = New-Object System.Drawing.Point(50, 50)
    $ProgressBar.Size = New-Object System.Drawing.Size(200, 20)
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
        }
        else {
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
        }
        else {
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
        }
        else {
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
}

function Clear-Internet {
  $chromeCachePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
  $edgeCachePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"

if (Test-Path $chromeCachePath) {
    Remove-Item -Path $chromeCachePath\* -Force -Recurse
    Write-Host "Cache for Google Chrome has been cleared."
} else {
    exit
}
if (Test-Path $edgeCachePath) {
  Remove-Item -Path $edgeCachePath\* -Force -Recurse
  Write-Host "Cache for Microsoft Edge has been cleared."
} else {
  exit
}

}