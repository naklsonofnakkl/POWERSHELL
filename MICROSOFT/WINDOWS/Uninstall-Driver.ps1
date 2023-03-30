#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.2
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Prompts the user for which driver they would like to uninstall
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

# Get all drivers and store them in a hashtable
$driverHashTable = @{}


$pnputilOutput = pnputil.exe /enum-drivers

$pnputilOutput | Where-Object { $_ -like "Published name:*" } | ForEach-Object {
    $driverLine = $_
    $publishedName = $driverLine.Substring(16).Trim()

    $nextLine = $pnputilOutput[$pnputilOutput.IndexOf($driverLine) + 1]
    $classGuid = $nextLine.Substring(14).Trim()

    $nextLine = $pnputilOutput[$pnputilOutput.IndexOf($driverLine) + 3]
    $driverDescription = $nextLine.Substring(25).Trim()

    $driverHashTable[$classGuid] = @{
        "PublishedName"     = $publishedName
        "DriverDescription" = $driverDescription
    }
}


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


<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

# Create a form to prompt the user to select a driver to uninstall
$form = New-Object System.Windows.Forms.Form
$form.Text = "Uninstall Driver"
$form.Width = 400
$form.Height = 200
$label = New-Object System.Windows.Forms.Label
$label.Text = "Select a driver to uninstall:"
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.AutoSize = $true
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10, 50)
$dropdown.Width = 300
$dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
foreach ($driver in $driverHashTable.Values) {
    if (![string]::IsNullOrEmpty($driver.DriverDescription)) {
        $dropdown.Items.Add($driver.DriverDescription)
    }
}
if ($dropdown.Items.Count -gt 0) {
    $dropdown.SelectedIndex = 0
}
$dropdown.Location = New-Object System.Drawing.Point(10, 50)
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10, 100)
$button.Text = "Uninstall"
$button.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $button
$form.Controls.Add($label)
$form.Controls.Add($dropdown)
$form.Controls.Add($button)

# Show the form and uninstall the selected driver
$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedDriver = $driverHashTable.Values | Where-Object { $_.DriverDescription -eq $dropdown.SelectedItem.ToString() }
    $global:output = "Press Ok to uninstall the $selectedDriver driver!"
    Pop-Success
    Try {
        pnputil.exe /delete-driver $selectedDriver.PublishedName /force /uninstall
        $global:output = "Driver '$($selectedDriver.DriverDescription)' has been uninstalled."
        Pop-Success
        
    }
    catch {
        $global:output = "Something went wrong!"
        Pop-Cancelled
        Clear-Installation
    }
    Pop-Restart
    
}
else {
    $global:output = "You have cancelled the script!"
    Pop-Cancelled
    Clear-Installation
}
