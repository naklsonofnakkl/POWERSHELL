#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.1
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Remove the Intel UHD Graphics driver
.DESCRIPTION
    - 
#>


<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\Driver_Uninstall.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Find the device ID of the user inputed driver
$device = Get-PnpDevice | Where-Object { $_.FriendlyName -eq $driverName }
$deviceId = $device.InstanceId
$driverName = ''
$infName = ''
$driver = Get-PnpDevice -PresentOnly | Where-Object { $_.DeviceId -eq $deviceId }

# Popup Message Box
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

# Function to prompt the user with a popup informing them to restart machine
function Pop-Restart {
    $restartCaption = "Restart Required"
    $restartTimeout = 120
    $restartIcon = [System.Windows.Forms.MessageBoxIcon]::Information

    $result = [System.Windows.Forms.MessageBox]::Show($global:output, $restartCaption, $restartIcon, $restartTimeout)

    if ($result -eq "Yes") {
        Clear-Installation
        Restart-Computer -Force
    }
    else {
        $global:output = "You have declined to restart!"
        Pop-Cancelled
        Clear-Installation
    }

}

# Function to prompt user for the Driver to uninstall
function Get-DriverName {

}

function Remove-DeviceDriver {
    # Uninstall the device driver using the device ID
    Try {
        # Disable-PnpDevice -InstanceId $deviceId -Confirm:$false

        # Uninstall the driver
        # pnputil /delete-driver $infName /uninstall
        write-host = $infName
        # Validate if the driver has been uninstalled
        if ($driver) {
            $global:output = "$driverName is installed."
            Pop-Cancelled
            Clear-Installation
        }
        else {
            $global:output = "$driverName is not installed."
            Pop-Success
            $global:output = "Restart is required to complete installation. Do you want to restart now?"
            # Pop-Restart
        }
    }
    Catch {
        $global:output = $_.Exception.Message
        Pop-Cancelled
    }
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

#Get-DriverName
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Uninstall Device Driver"
$form.Width = 400
$form.Height = 150
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$Form.AutoSize = $true
$form.TopMost = $true
$Form.AutoSizeMode = "GrowAndShrink"

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(380, 20)
$label.Text = "Type the name of the driver you wish to uninstall:"
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 50)
$textBox.Size = New-Object System.Drawing.Size(380, 20)
$form.Controls.Add($textBox)

$continueButton = New-Object System.Windows.Forms.Button
$continueButton.Location = New-Object System.Drawing.Point(220, 85)
$continueButton.Size = New-Object System.Drawing.Size(80, 30)
$continueButton.Text = "Continue"
$continueButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $continueButton
$form.Controls.Add($continueButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(310, 85)
$cancelButton.Size = New-Object System.Drawing.Size(80, 30)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$resultDriver = $form.ShowDialog()
if ($resultDriver -eq [System.Windows.Forms.DialogResult]::OK) {
    $driverName = $textBox.Text
    $global:output = "Attemping to uninstall the driver: $driverName"
    Pop-Success
    # Find the INF file name for the driver
    $infName = pnputil /enum-drivers | Select-String -Pattern $driverName | ForEach-Object { $_.ToString().Split(',')[1].Trim() }
    Remove-DeviceDriver
}
else {
    $global:output = "You have cancelled the script!"
    Pop-Cancelled
    Clear-Installation
}
