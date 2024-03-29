#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.1.2.1
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Script to download entire Steam Library
.DESCRIPTION
    - Promps the user for their API key and SteamID64
#>

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

# VARIABLE PARADISE

# create a temp folder & Start log file
# THIS IS WHERE LOGS WILL ALSO BE LOCATED!
# C:\Users\[USERNAME]\AppData\Local\Temp\SteamLib_temp
$tempDirCreate = New-Item -ItemType Directory -Path $env:TEMP -Name SteamLib_temp -Force
$tempDir = $tempDirCreate
# LOGS!!
$ErrorActionPreference = "Stop"
$appLogs = "$tempDir\SteamLib.log"
# Check if a transcript is running
if ($global:PSIsTranscripting) {
    # Stop the transcript
    Stop-Transcript
    Start-Transcript -Path $appLogs -Append
}
else {
    Start-Transcript -Path $appLogs -Append
}

# Grab the user's name for json file nameing
$username = Split-Path $env:USERPROFILE -Leaf


# FUNCTION JUNCTION!!

# Check to make sure AWSPowershell module is instlled
<#
Function Find-AWSPowerShell {
    # Check if the AWSPowershell Module is installed
    if (-not(Get-Module -Name AWSPowershell -ListAvailable)) {
        # Install the Module
        Install-Module -Name AWSPowershell -Scope CurrentUser -Force
    } 
}
#>

# Function that runs whenever script ends
function Clear-Installation {
    # Dispose of any forms
    $form.Dispose()
    $form.Close()
    Stop-Transcript
    # Get all the files in the directory
    $files = Get-ChildItem $tempDir

    # Loop through each file
    foreach ($file in $files) {
        # Check if the file is not the one you want to keep
        if ($file.Name -ne "$username.json" -and $file.Name -ne "SteamLib.log") {
            # Delete the file
            Remove-Item $file.FullName
        }
    }
}

# install modules, import them and then export JSON file
function Join-JsonTable {
    # Import the JSON data
    $Json_Url = $apiUrl
    $owned_games_Json = "$tempDir\$username.json"
    Invoke-WebRequest -Uri $Json_Url -OutFile $owned_games_Json
}


#This is a form to ask for the users API and SteamID
Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Enter API Key"
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form.AutoSize = $true
$form.TopMost = $true
$Form.AutoSizeMode = "GrowAndShrink"

# Create the message label
$MessageLabel = New-Object System.Windows.Forms.Label
$MessageLabel.Location = New-Object System.Drawing.Point(10, 10)
$MessageLabel.Size = New-Object System.Drawing.Size(280, 20)
$MessageLabel.Text = "Please enter your API Key (must be 32 characters):"
$Form.Controls.Add($MessageLabel)

# Create the API Key textbox
$ApiKeyTextBox = New-Object System.Windows.Forms.TextBox
$ApiKeyTextBox.Location = New-Object System.Drawing.Point(10, 30)
$ApiKeyTextBox.Size = New-Object System.Drawing.Size(280, 20)
$Form.Controls.Add($ApiKeyTextBox)

# Create the "Next" button
$NextButton = New-Object System.Windows.Forms.Button
$NextButton.Location = New-Object System.Drawing.Point(115, 70)
$NextButton.Size = New-Object System.Drawing.Size(75, 23)
$NextButton.Text = "Next"
$NextButton.Enabled = $false
$NextButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$Form.AcceptButton = $NextButton
$Form.Controls.Add($NextButton)

# Create the event handler for the API Key textbox
$ApiKeyTextBox.add_TextChanged({
        if ($ApiKeyTextBox.Text.Length -eq 32) {
            $NextButton.Enabled = $true
        }
        else {
            $NextButton.Enabled = $false
        }
    })

# Show the form and wait for a result
$Result = $Form.ShowDialog()

# If the "Next" button was clicked and the API Key is 32 characters long, set the $ApiKey variable
if ($Result -eq [System.Windows.Forms.DialogResult]::OK -and $ApiKeyTextBox.Text.Length -eq 32) {
    $Apikey = $ApiKeyTextBox.Text
    Add-Type -AssemblyName System.Windows.Forms

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Enter SteamID64"
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Form.AutoSize = $true
    $form.TopMost = $true
    $Form.AutoSizeMode = "GrowAndShrink"

    # Create the message label
    $MessageLabel = New-Object System.Windows.Forms.Label
    $MessageLabel.Location = New-Object System.Drawing.Point(10, 10)
    $MessageLabel.Size = New-Object System.Drawing.Size(280, 20)
    $MessageLabel.Text = "Please enter your SteamID64 (must be 17 characters):"
    $Form.Controls.Add($MessageLabel)

    # Create the API Key textbox
    $SteamID64TextBox = New-Object System.Windows.Forms.TextBox
    $SteamID64TextBox.Location = New-Object System.Drawing.Point(10, 30)
    $SteamID64TextBox.Size = New-Object System.Drawing.Size(280, 20)
    $Form.Controls.Add($SteamID64TextBox)

    # Create the "Next" button
    $NextButton = New-Object System.Windows.Forms.Button
    $NextButton.Location = New-Object System.Drawing.Point(115, 70)
    $NextButton.Size = New-Object System.Drawing.Size(75, 23)
    $NextButton.Text = "Next"
    $NextButton.Enabled = $false
    $NextButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $NextButton
    $Form.Controls.Add($NextButton)

    # Create the event handler for the API Key textbox
    $SteamID64TextBox.add_TextChanged({
            if ($SteamID64TextBox.Text.Length -eq 17) {
                $NextButton.Enabled = $true
            }
            else {
                $NextButton.Enabled = $false
            }
        })

    # Show the form and wait for a result
    $Result = $Form.ShowDialog()

    # If the "Next" button was clicked and the API Key is 32 characters long, set the $ApiKey variable
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK -and $SteamID64TextBox.Text.Length -eq 17) {
        $SteamID64 = $SteamID64TextBox.Text
        #URL for API
        $apiUrl = "https://api.steampowered.com/IPlayerService/GetOwnedGames/v1/?key=$Apikey&steamid=$SteamID64&include_appinfo=1&format=json"
        
        # Grabs the list of Steam Games
        # Export the JSON file
        Join-JsonTable
<#
        # Check that AWSPowershell module is installed
        Find-AWSPowerShell

        # Import the AWSPowershell Module
        Import-AWSPowershell

        #Set the name of S3 bucket
        $bucketname = ""
        $localFilePath = $tempDir
        $keyname = ""

        # Upload to S3


#>        
        # Open up file explorer for user
        Invoke-Item -Path $tempDir

        # Clean up after itself
        Clear-Installation        
    }
    else {
        Clear-Installation
        exit
    }
}
else {
    Clear-Installation
    exit
}