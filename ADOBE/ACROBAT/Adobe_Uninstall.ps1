#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.1.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Uninstall any Adobe product without needing to touch pesky UI
.DESCRIPTION
    - Closes Microsoft Applications
    - Uninstalls Adobe Acrobat DC
    [You can alter this to work for other versions as well by altering the
    function 'Remove-AdobeAcrobatDC' and replacing -Filter "Name = 'Adobe Acrobat (64-bit)'"
    with the name of your version of Adobe found using the command 'Get-WmiObject']
    - Clears the registry of Adobe products
    - Restarts computer

#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Directories
$tempDir = $env:TEMP
$verifyFolder = "$env:TEMP\AdobeUninstall"
 
# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$appLogs = "$tempDir\Adobe_Uninstall.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

# Functions to display the popup messages
function Show-RestartPopup {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object Drawing.Size(300, 150)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    $label = New-Object Windows.Forms.Label
    $label.Text = "Press OK to restart the computer."
    $label.Location = New-Object Drawing.Point(20, 20)
    $label.AutoSize = $true

    $okButton = New-Object Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object Drawing.Point(110, 80)
    $okButton.add_Click({
            $form.Close()
            Restart-Computer -Force
        })

    $form.Controls.Add($label)
    $form.Controls.Add($okButton)

    $form.ShowDialog()
}
function Show-MissingPopup {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object Drawing.Size(300, 150)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    
    $label = New-Object Windows.Forms.Label
    $label.Text = "Adobe Acrobat DC was not found07081989.`nContinue running script?"
    $label.Location = New-Object Drawing.Point(70, 20)
    $label.AutoSize = $true
    
    $yesButton = New-Object Windows.Forms.Button
    $yesButton.Text = "Yes"
    $yesButton.Location = New-Object Drawing.Point(50, 80)
    $yesButton.add_Click({
            $form.Close()
    
        })
    $noButton = New-Object Windows.Forms.Button
    $noButton.Text = "No"
    $noButton.Location = New-Object Drawing.Point(150, 80)
    $noButton.add_Click({
            $form.Close()
            Clear-Installation
            exit
        })
    
    $form.Controls.Add($label)
    $form.Controls.Add($yesButton)
    $form.Controls.Add($noButton)
    
    $form.ShowDialog()
}


# Download the Adobe cleaner files
function Get-Download {
    # Specify the URL of the zip file to download
    $adobeTemp = "$env:TEMP\AdobeUninstall"
    # Check if the folder exists
    if (Test-Path -Path $adobeTemp -PathType Container) {
        # Delete the folder if it exists
        Remove-Item -Path $adobeTemp -Recurse -Force
    }
    $folderName = New-Item -Path "$env:TEMP\AdobeUninstall" -ItemType Directory -Force
    $adobe32down = (Invoke-WebRequest -Uri "https://www.adobe.com/devnet-docs/acrobatetk/tools/Labs/AcroCleaner_DC2015.zip" -OutFile $folderName\AcroCleaner_DC2015.zip -UseDefaultCredentials -UseBasicParsing ).Content
    $adobe64down = (Invoke-WebRequest -Uri "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2100120135/x64/AdobeAcroCleaner_DC2021.exe" -OutFile $folderName\AdobeAcroCleaner_DC2021.exe -UseDefaultCredentials -UseBasicParsing ).Content
    Expand-Archive -Path $folderName\AcroCleaner_DC2015.zip -DestinationPath $folderName\AcroCleaner -Force
}

# Function to clean up the leftover downloaded files
function Clear-Installation {
    Remove-Item -Path "$env:TEMP\AdobeUninstall" -Force
    Stop-Transcript    
}

# Function to automatically close Adobe Acrobat
function Close-AdobeAcrobat {
  
    if (Get-Process -Name "Acrobat" -ErrorAction SilentlyContinue) {
        # Close out of Adobe
        Stop-Process -ProcessName Acrobat -force

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
        Write-host "Adobe Acrobat already closed, skipping..."
    }
}

#Function to automatically close Microsoft Products
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
        Write-host "Outlook already closed, skipping..."
    }
}
function Close-MicrosoftWord {
  
    if (Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue) {
        # Close out of Outlook
        Stop-Process -name WINWORD -force
  
        # Set the duration of the timer in seconds
        $duration = 5
  
        # Initialize the progress bar
        Write-Progress -Activity "Waiting for $duration seconds while Word closes..." -PercentComplete 0
  
        # Loop through the timer and update the progress bar
        for ($i = 1; $i -le $duration; $i++) {
            # Update the progress bar with the current progress
            $percent = ($i / $duration) * 100
            Write-Progress -Activity "Waiting for $duration seconds while Word closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
            # Pause for 1 second
            Start-Sleep -Seconds 1
        }
  
        # Clear the progress bar once the timer is complete
        Write-Progress -Completed -Activity "Microsoft Word is Closed!"
    }
    else {
        Write-host "Word already closed, skipping..."
    }
}
function Close-MicrosoftExcel {
  
    if (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue) {
        # Close out of Outlook
        Stop-Process -name EXCEL -force
  
        # Set the duration of the timer in seconds
        $duration = 5
  
        # Initialize the progress bar
        Write-Progress -Activity "Waiting for $duration seconds while Excel closes..." -PercentComplete 0
  
        # Loop through the timer and update the progress bar
        for ($i = 1; $i -le $duration; $i++) {
            # Update the progress bar with the current progress
            $percent = ($i / $duration) * 100
            Write-Progress -Activity "Waiting for $duration seconds while Excel closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
            # Pause for 1 second
            Start-Sleep -Seconds 1
        }
  
        # Clear the progress bar once the timer is complete
        Write-Progress -Completed -Activity "Microsoft Excel is Closed!"
    }
    else {
        Write-host "Excel already closed, skipping..."
    }
}
function Close-MicrosoftPowerPoint {
  
    if (Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue) {
        # Close out of Outlook
        Stop-Process -name POWERPNT -force
  
        # Set the duration of the timer in seconds
        $duration = 5
  
        # Initialize the progress bar
        Write-Progress -Activity "Waiting for $duration seconds while Powerpoint closes..." -PercentComplete 0
  
        # Loop through the timer and update the progress bar
        for ($i = 1; $i -le $duration; $i++) {
            # Update the progress bar with the current progress
            $percent = ($i / $duration) * 100
            Write-Progress -Activity "Waiting for $duration seconds while Powerpoint closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
            # Pause for 1 second
            Start-Sleep -Seconds 1
        }
  
        # Clear the progress bar once the timer is complete
        Write-Progress -Completed -Activity "Microsoft Powerpoint is Closed!"
    }
    else {
        Write-host "PowerPoint already closed, skipping..."
    }
}
function Close-MicrosoftOneNote {
  
    if (Get-Process -Name "ONENOTE" -ErrorAction SilentlyContinue) {
        # Close out of Outlook
        Stop-Process -name ONENOTE -force
  
        # Set the duration of the timer in seconds
        $duration = 5
  
        # Initialize the progress bar
        Write-Progress -Activity "Waiting for $duration seconds while OneNote closes..." -PercentComplete 0
  
        # Loop through the timer and update the progress bar
        for ($i = 1; $i -le $duration; $i++) {
            # Update the progress bar with the current progress
            $percent = ($i / $duration) * 100
            Write-Progress -Activity "Waiting for $duration seconds while OneNote closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
            # Pause for 1 second
            Start-Sleep -Seconds 1
        }
  
        # Clear the progress bar once the timer is complete
        Write-Progress -Completed -Activity "Microsoft OneNote is Closed!"
    }
    else {
        Write-host "OneNote already closed, skipping..."
    }
}
function Close-MicrosoftEdge {
  
    if (Get-Process -Name "msedge" -ErrorAction SilentlyContinue) {
        # Close out of Outlook
        Stop-Process -name msedge -force
  
        # Set the duration of the timer in seconds
        $duration = 5
  
        # Initialize the progress bar
        Write-Progress -Activity "Waiting for $duration seconds while Edge closes..." -PercentComplete 0
  
        # Loop through the timer and update the progress bar
        for ($i = 1; $i -le $duration; $i++) {
            # Update the progress bar with the current progress
            $percent = ($i / $duration) * 100
            Write-Progress -Activity "Waiting for $duration seconds while Edge closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
            # Pause for 1 second
            Start-Sleep -Seconds 1
        }
  
        # Clear the progress bar once the timer is complete
        Write-Progress -Completed -Activity "Microsoft Edge is Closed!"
    }
    else {
        Write-host "Edge already closed, skipping..."
    }
}

# Funvtion to automatically close Google Chrome
function Close-GoogleChrome {
  
    if (Get-Process -Name "chrome" -ErrorAction SilentlyContinue) {
        # Close out of Outlook
        Stop-Process -name chrome -force
  
        # Set the duration of the timer in seconds
        $duration = 5
  
        # Initialize the progress bar
        Write-Progress -Activity "Waiting for $duration seconds while Chrome closes..." -PercentComplete 0
  
        # Loop through the timer and update the progress bar
        for ($i = 1; $i -le $duration; $i++) {
            # Update the progress bar with the current progress
            $percent = ($i / $duration) * 100
            Write-Progress -Activity "Waiting for $duration seconds while Chrome closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
            # Pause for 1 second
            Start-Sleep -Seconds 1
        }
  
        # Clear the progress bar once the timer is complete
        Write-Progress -Completed -Activity "Google Chrome is Closed!"
    }
    else {
        Write-host "Chrome already closed, skipping..."
    }
}

# Remove the Adobe Software
function Remove-AdobeAcrobatRegistry {
    # Specify the path to the executable file
    $adobea = "$folderName\AdobeAcroCleaner_DC2021.exe"
    $adobeb = "$folderName\AcroCleaner\AdobeAcroCleaner_DC2015.exe"
    $parametersa = "/Silent /ScanForOthers=1"
    $parametersb = "/Silent /ScanForOthers=1 /Product=0"
    $parametersc = "/Silent /ScanForOthers=1 /Product=1"

    # Execute the executable file
    Write-host "Cleaning Adobe Cache..."
    Start-Process -FilePath $adobea -ArgumentList $parametersa -Wait
    Start-Process -FilePath $adobeb -ArgumentList $parametersb -Wait
    Start-Process -FilePath $adobeb -ArgumentList $parametersc -Wait
}

#Function to remove Adobe Acrobat DC 64-bit
function Remove-AdobeAcrobatDC {

    $app = Get-WmiObject -Class Win32_Product -Filter "Name = 'Adobe Acrobat (64-bit)'"
    if ($null -ne $app) {
        Write-host "Uninstalling Adobe Acrobat..."
        $app.Uninstall()
    }
    else {
        Write-host "Adobe Acrobat DC isn't installed!"
        Show-MissingPopup
    }
}

#Functions to setup script to resume on restart
function Restart-AfterUninstall {
    Write-host "Preparing to Restart..."
    Show-RestartPopup
    Start-Sleep -Seconds 60
    Restart-Computer -Force
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>
    Write-Host "Beginning Adobe Uninstall..."
    Start-Sleep 3
    $folderName = New-Item -Path "$env:TEMP\AdobeUninstall" -ItemType Directory -Force
    Get-Download
    Close-AdobeAcrobat
    Close-MicrosoftOutlook
    Close-MicrosoftExcel
    Close-MicrosoftWord
    Close-MicrosoftPowerPoint
    Close-MicrosoftOneNote
    Close-MicrosoftEdge
    Close-GoogleChrome
    Remove-AdobeAcrobatDC
    Remove-AdobeAcrobatRegistry
    Restart-AfterUninstall