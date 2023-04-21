#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Script to change the timezone of a computer locked via GPO
.DESCRIPTION
    - Prompts the user to select a timezone
    - changes timezone to selected option
#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\change_timezone.log"
$ErrorActionPreference = "Continue"
Start-Transcript -Path $appLogs -Append

$global:output = ''

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

# Function to clean up the leftover files
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

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the popup window
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select Timezone'
$form.Width = 300
$form.Height = 150
$form.AutoSize = $true
$form.StartPosition = 'CenterScreen'
$form.AutoSizeMode = "GrowAndShrink"

# Create the message label
$messageLabel = New-Object System.Windows.Forms.Label
$messageLabel.Location = New-Object System.Drawing.Point(10, 10)
$messageLabel.Size = New-Object System.Drawing.Size(280, 20)
$messageLabel.Text = "Please select a timezone:"
$form.Controls.Add($messageLabel)

# Define the dropdown menu
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.DropDownStyle = 'DropDownList'
$dropdown.Location = New-Object System.Drawing.Point(10, 30)
$dropdown.Size = New-Object System.Drawing.Size(280, 20)

# Add timezone options to dropdown menu
$timezoneOptions = [System.TimeZoneInfo]::GetSystemTimeZones()
foreach ($timezone in $timezoneOptions) {
    $dropdown.Items.Add($timezone.DisplayName)
}

# Add the dropdown menu to the window
$form.Controls.Add($dropdown)

# Add a 'set' button to the window
$button = New-Object System.Windows.Forms.Button
$button.Text = 'Set'
$button.Location = New-Object System.Drawing.Point(100, 80)
$button.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $button
$form.Controls.Add($button)

# Show the popup window and get user selection
$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedTimezone = $timezoneOptions[$dropdown.SelectedIndex]
    $utcOffset = [int]($selectedTimezone.BaseUtcOffset.TotalMinutes / 60)
    Set-TimeZone -Id $selectedTimezone.Id
    Write-Output = $selectedTimezone
    $global:output = "The timezone has been set to $($selectedTimezone.StandardName).`nUTC offset: $utcOffset hours."
    Pop-Success
    Clear-Installation
}
else {
    $global:output = "script has been cancelled!"
    Pop-Cancelled
    Clear-Installation
}
