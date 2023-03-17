#Requires -Modules Appx
#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.3
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Script to install Microsoft Store Apps without using the MS Store App!
.DESCRIPTION
    - Creates a temporary folder
    - install/import the Appx module (Required!)
    - Prompt which app you wish to install from the $options list
    - Validates the app isn't already installed
    - Downoads necessary appbundle for installation
    - installs the requested Application
    - Cleans itself up
#>

# create a temp folder to download the app
# THIS IS WHERE LOGS WILL ALSO BE LOCATED!
$tempDir = New-Item -ItemType Directory -Path $env:TEMP -Name MSSTORE_temp -Force
$appDownloadPath = $tempDir
$appDownload = $appDownloadPath
$appLogs = "$appDownload\MSSTORE.log"

# LOGS!!
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Import Necessary Modules
# Specify the name of the module to check
$moduleName = "Appx"

# List of apps to install
$appxBundles = Get-ChildItem -Path $appDownload -Filter *.appxbundle

# Set the download for the app.
# You can find the URL here: https://apps.microsoft.com/store/apps
# Define the hashtable with application names and URLs
$options = @{
    'Microsoft Snip'             = 'https://apps.microsoft.com/store/detail/snipping-tool/9MZ95KL8MR0L, Microsoft.ScreenSketch';
    'Microsoft Sticky Notes'     = 'https://apps.microsoft.com/store/detail/microsoft-sticky-notes/9NBLGGH4QGHW?ocid=Apps_O_WOL_FavTile_App_ForecaWeather_Pos5, Microsoft.MicrosoftStickyNotes' ;
    'Microsoft Quick Assist' = 'https://apps.microsoft.com/store/detail/quick-assist/9P7BP5VNWKX5?ocid=Apps_O_WOL_FavTile_App_ForecaWeather_Pos5, MicrosoftCorporationII.QuickAssist' ;
    'Microsoft Clock'            = 'https://apps.microsoft.com/store/detail/windows-clock/9WZDNCRFJ3PR?ocid=Apps_O_WOL_FavTile_App_ForecaWeather_Pos5, Microsoft.WindowsAlarms'
}

#Function to clean up the mess this makes
function Cleanup-Installation {
    # Dispose of any forms
    $form.Dispose()
    $form.Close()
    Stop-Transcript
    Move-Item -Path $appLogs -Destination "C:\" -Force -ErrorAction SilentlyContinue
    Get-ChildItem -Path $appDownloadPath -Include *.appx, *.appxbundle -Recurse | Remove-Item -Force
}

# Make sure you can even install apps
# Add-WindowsCapability -Online -Name AppxVC

# Check if the Appx module is already installed
if (Get-Module -Name $moduleName -ListAvailable) {
    Write-Host "$moduleName module is already installed!"
    Write-Host "Moving right along now..."
}
else {
    # popup variables based on pass or fail state
    $appxFailTxt = "$moduleName must be updated or installed before the Application can continue!"
    $appxPassTxt = "$moduleName module has been installed."
    $appxFailBtn = New-Object System.Management.Automation.Host.ChoiceDescription "&OK", "OK"
    $appxPassBtn = New-Object System.Management.Automation.Host.ChoiceDescription "&OK", "OK"
    $appxFailRes = $Host.UI.PromptForChoice("Message", $appxFailTxt, @($appxFailBtn), 0)
    $appxPassRes = $Host.UI.PromptForChoice("Message", $appxPassTxt, @($appxPassBtn), 0)

    Add-Type -AssemblyName System.Windows.Forms
    $messageBoxInput = New-Object System.Windows.Forms.MessageBoxButtons
    $messageBoxInput.YesNo
    $messageBoxResult = [System.Windows.Forms.MessageBox]::Show('$moduleName module is not installed. Do you want to install it? *Required!', 'Confirmation', $messageBoxInput)

    if ($messageBoxResult -eq 'Yes') {
        # Install the module
        Install-Module $moduleName -Scope CurrentUser -Force
        if (Get-Module -Name $moduleName -ListAvailable) {
            Write-Host "$moduleName module has been installed!"
            $appxPassRes
        }
        else {
            $appxFailRes
            Write-Host "$moduleName must be updated or installed before the Application can continue!" 
            Cleanup-Installation
        } 
        else {
            $appxFailRes
            Write-Host "$moduleName must be updated or installed before the Application can continue!"
            Cleanup-Installation 
        }
    }
}
Import-Module Appx

# Function to download the correct AppPackage
# Creator: https://serverfault.com/users/616108/yorai-levi
# Source: https://serverfault.com/questions/1018220/how-do-i-install-an-app-from-windows-store-using-powershell

function Download-AppxPackage {
    [CmdletBinding()]
    param (
        [string]$Uri,
        [string]$Path = "."
    )
       
    process {
        echo ""
        $StopWatch = [system.diagnostics.stopwatch]::startnew()
        $Path = (Resolve-Path $Path).Path
        #Get Urls to download
        Write-Host -ForegroundColor Yellow "Processing $Uri"
        $WebResponse = Invoke-WebRequest -UseBasicParsing -Method 'POST' -Uri 'https://store.rg-adguard.net/api/GetFiles' -Body "type=url&url=$Uri&ring=Retail" -ContentType 'application/x-www-form-urlencoded'
        $LinksMatch = ($WebResponse.Links | where { $_ -like '*.appx*' } | where { $_ -like '*_neutral_*' -or $_ -like "*_" + $env:PROCESSOR_ARCHITECTURE.Replace("AMD", "X").Replace("IA", "X") + "_*" } | Select-String -Pattern '(?<=a href=").+(?=" r)').matches.value
        $Files = ($WebResponse.Links | where { $_ -like '*.appx*' } | where { $_ -like '*_neutral_*' -or $_ -like "*_" + $env:PROCESSOR_ARCHITECTURE.Replace("AMD", "X").Replace("IA", "X") + "_*" } | where { $_ } | Select-String -Pattern '(?<=noreferrer">).+(?=</a>)').matches.value
        #Create array of links and filenames
        $DownloadLinks = @()
        for ($i = 0; $i -lt $LinksMatch.Count; $i++) {
            $Array += , @($LinksMatch[$i], $Files[$i])
        }
        #Sort by filename descending
        $Array = $Array | sort-object @{Expression = { $_[1] }; Descending = $True }
        $LastFile = "temp123"
        for ($i = 0; $i -lt $LinksMatch.Count; $i++) {
            $CurrentFile = $Array[$i][1]
            $CurrentUrl = $Array[$i][0]
            #Find first number index of current and last processed filename
            if ($CurrentFile -match "(?<number>\d)") {}
            $FileIndex = $CurrentFile.indexof($Matches.number)
            if ($LastFile -match "(?<number>\d)") {}
            $LastFileIndex = $LastFile.indexof($Matches.number)
    
            #If current filename product not equal to last filename product
            if (($CurrentFile.SubString(0, $FileIndex - 1)) -ne ($LastFile.SubString(0, $LastFileIndex - 1))) {
                #If file not already downloaded, add to the download queue
                if (-Not (Test-Path "$Path\$CurrentFile")) {
                    "Downloading $Path\$CurrentFile"
                    $FilePath = "$Path\$CurrentFile"
                    $FileRequest = Invoke-WebRequest -Uri $CurrentUrl -UseBasicParsing #-Method Head
                    [System.IO.File]::WriteAllBytes($FilePath, $FileRequest.content)
                }
                #Delete file outdated and already exist
            }
            elseif ((Test-Path "$Path\$CurrentFile")) {
                Remove-Item "$Path\$CurrentFile"
                "Removing $Path\$CurrentFile"
            }
            $LastFile = $CurrentFile
        }
        "Time to process: " + $StopWatch.ElapsedMilliseconds
    }
}

# Create the form
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Avoid Microsoft Store"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.AutoSize = $true
$form.AutoSizeMode = "GrowAndShrink"

# Create the message label
$messageLabel = New-Object System.Windows.Forms.Label
$messageLabel.Location = New-Object System.Drawing.Point(10, 10)
$messageLabel.Size = New-Object System.Drawing.Size(280, 20)
$messageLabel.Text = "Choose an Application to Install:"
$form.Controls.Add($messageLabel)

# Create the dropdown
$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10, 30)
$dropdown.Size = New-Object System.Drawing.Size(280, 20)

foreach ($option in $options.GetEnumerator()) {
    [void] $dropdown.Items.Add($option.Key)
}

$form.Controls.Add($dropdown)

# Create the install button
$installButton = New-Object System.Windows.Forms.Button
$installButton.Location = New-Object System.Drawing.Point(115, 70)
$installButton.Size = New-Object System.Drawing.Size(75, 23)
$installButton.Text = "Install"
$installButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $installButton
$form.Controls.Add($installButton)

# Create the cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(195, 70)
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# Show the form and wait for a result
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Get the selected option
        $selectedOption = $dropdown.SelectedItem.ToString()

        # Split the option's value into two variables
        $appUrl, $appName = $options[$selectedOption] -split ', '
    
        # Show the installation message
        $installMessage = "$selectedOption is currently installing!"
        Write-Host $installMessage
    # Set the appUrl and appName variables
    #$appName = $dropdown.SelectedItem.ToString()
    #$appUrl = $applications[$appName]
    # Check if the app is installed
    $package = Get-AppxPackage | Where-Object { $_.Name -eq $appName }
    if ($package) {
        Write-Host "$appName is already installed."
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show('Application is already installed!', 'Install Completed', 'OK', 'Information')
        Cleanup-Installation
    }
    else {
        Write-Host "$appName is not installed."
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $appcaption = "Confirmation"
        $appmessage = "Do you want to install $appName ?"
        $appicon = [System.Windows.Forms.MessageBoxIcon]::Question
        $appbuttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
        $appresult = [System.Windows.Forms.MessageBox]::Show($appmessage, $appcaption, $appbuttons, $appicon)

        # If the user selects Yes and continues with the installation    
        if ($appresult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "User clicked Yes"

            # If the user didn't select cancel, retrieve the URL for the selected application and proceed with installation
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "You selected $appName. The download URL is $appUrl."

                # Download the app files
                Download-AppxPackage $appUrl $appDownloadPath

                # Set downloaded files into a variable
                $appxbundles = Get-ChildItem $appDownloadPath -Filter *.appxbundle
                
                # Install the selected app(s)
                foreach ($appxbundle in $appxbundles) {
                    Add-AppxPackage -Path $appxbundle.FullName
                }
                
            }
                # Validate if the application installed successfully
                if (Get-AppxPackage -Name $appName) {
                    Write-Host "$appName is installed."
                    Add-Type -AssemblyName System.Windows.Forms
                    [System.Windows.Forms.MessageBox]::Show('Application was successfully installed!', 'Install Completed', 'OK', 'Information')
                    Cleanup-Installation
                }
                else {
                    Write-Host "$appName is not installed."
                    Add-Type -AssemblyName System.Windows.Forms
                    $appErrorMessage = "$appName did not install correctly! Please view the log file for more details: C:\app.log"
                    [System.Windows.Forms.MessageBox]::Show($appErrorMessage, 'Error', 'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
                    Cleanup-Installation
                }
        }
            else {
                Write-Host "You canceled the installation."
                Add-Type -AssemblyName System.Windows.Forms
                $appCancelMessage = "$appName installation has been cancelled!"
                [System.Windows.Forms.MessageBox]::Show($appCancelMessage, 'Error', 'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
                Cleanup-Installation
            }
        }
    }
else {
    Write-Host "User clicked No"
    Add-Type -AssemblyName System.Windows.Forms
    $appErrorMessage = "$appName install was declined. Appication will now close."
    [System.Windows.Forms.MessageBox]::Show($appErrorMessage, 'Error', 'OK', [System.Windows.Forms.MessageBoxIcon]::Error)
    Cleanup-Installation
}
exit