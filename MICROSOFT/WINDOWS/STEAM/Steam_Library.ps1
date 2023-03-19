<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.7
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Script to download entire Steam Library that are Multiplayer
.DESCRIPTION
    - Promps the user for their API key and SteamID64
#>

# FUNCTION JUNCTION!!
function Clear-Installation {
    # Dispose of any forms
    $form.Dispose()
    $form.Close()
    Stop-Transcript
    # Get all the files in the directory
    $files = Get-ChildItem $appDownloadpath

    # Loop through each file
    foreach ($file in $files) {
        # Check if the file is not the one you want to keep
        if ($file.Name -ne "Steam_Multiplayer.xlsx" -and $file.Name -ne "SteamLib.log") {
            # Delete the file
            Remove-Item $file.FullName
        }
    }
}

function Get-NuGet {
    $packageName = "NuGet"
    $version = "2.8.5.208"
    if (Get-Package -Name $packageName -ErrorAction SilentlyContinue | Where-Object { $_.Version -eq $version }) {
        Write-Host "NuGet is Installed!"
    }
    else {
        Write-Host "Installing NuGet!"
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    }
}

function Join-JsonTable {

    
    # Install the ImportExcel and PSWriteExcel modules if they're not already installed
    Install-Module -Name ImportExcel, PSWriteExcel

    # Load the modules into the current session
    Import-Module -Name ImportExcel, PSWriteExcel

    # Define the URL of the JSON data
    $jsonUrl = $owned_games

    # Import the JSON data
    $jsonData = Invoke-RestMethod -Uri $jsonUrl

    # Convert the JSON data to a PowerShell object
    $data = ConvertFrom-Json -InputObject $jsonData

    # Export the data to an Excel file
    $excelFilePath = "$appDownloadPath\steam_library_stats.xlsx"
    $data | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter

}

# Function to grab my Big Beautiful Excel module
function Get-BBE {
    # Modules to Import!
    $url = "http://github.naklwilson.net/MICROSOFT/OFFICE/EXCEL/BBE.psm1"
    $outputFile = "$appDownload\BBE.psm1"
    Invoke-WebRequest -Uri $url -OutFile $outputFile
    Import-Module $outputFile
}
#Use the Steam API to create a varible with all the users games listed
function Get-SteamLibrary {
    # Fetch the list of owned games
    $owned_games_url = "https://api.steampowered.com/IPlayerService/GetOwnedGames/v1/?key=$secureAPI&steamid=$SteamID64&include_appinfo=1&format=json"
    $owned_games_response = Invoke-RestMethod -Uri $owned_games_url
    $owned_games = $owned_games_response.response.games
    Write-Host "I HAS THE GAMES!"
}

# create a temp folder & Start log file
Function New-ApplicationStart {
    # THIS IS WHERE LOGS WILL ALSO BE LOCATED!
    # C:\Users\[USERNAME]\AppData\Local\Temp\SteamLib_temp
    $tempDir = New-Item -ItemType Directory -Path $env:TEMP -Name SteamLib_temp -Force
    # LOGS!!
    $ErrorActionPreference = "Stop"
    $appLogs = "$tempDir\SteamLib.log"
    Start-Transcript -Path $appLogs -Append
}

function Add-ModuleExcel {

    $moduleName = "ImportExcel"
    
    # Check if the module is installed
    $installedModule = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName }
        
    if ($installedModule) {
        Write-Host "The $moduleName module is installed."
    
        # execute the rest of the script 
        Join-JsonTable
        Format-SteamXlsx
        Clear-Installation
        Invoke-Item -Path $appDownloadPath
    }
    else {
        Write-Host "The $moduleName module is not installed."
        # Install the missing module and run the rest of the script
        Get-NuGet
        Join-JsonTable
        Format-SteamXlsx
        Clear-Installation
        Invoke-Item -Path $appDownloadPath
    }
    
} 

New-ApplicationStart
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
    $secureAPI = $Apikey | ConvertTo-SecureString -AsPlainText -Force
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

        Get-SteamLibrary
        Get-BBE
        Write-Host $outputFile
        Add-ModuleExcel
        
        
    }
    Clear-Installation
}



