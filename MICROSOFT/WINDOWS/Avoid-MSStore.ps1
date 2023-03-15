<#
.SYNOPSIS
    Script to install Microsoft Store Apps progromatically!
.DESCRIPTION
    This script will install/import the Appx module
    Import a function to download Appx Packages
    check if the app you wish to install is either:
    currently installed or needs to be installed
    and then installs the application
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL
#>

# create a temp folder to download the app
# THIS IS WHERE LOGS WILL ALSO BE LOCATED!
$appDownload = New-Item -ItemType Directory -Path "$ENV:USERPROFILE\App_Temp"
$appLogs = "$appDownload\Console.log"
# LOGS!!
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Import Necessary Modules
# Specify the name of the module to check
$moduleName = "Appx"

# Set the name of the app you want to check for
$appName = "Microsoft.ScreenSketch"

# Set the download for the app.
# You can find the URL here: https://apps.microsoft.com/store/apps
$appUrl = "https://apps.microsoft.com/store/detail/snipping-tool/9MZ95KL8MR0L"



# Check if the Appx module is already installed
if (Get-Module -Name $moduleName -ListAvailable) {
    Write-Host "$moduleName module is already installed!" | Out-File -FilePath $appLogs -Append
    Write-Host "Moving right along now..." | Out-File -FilePath $appLogs -Append
} else {
$appxFailTxt = "$moduleName must be updated or installed before the Application can continue!"
$appxPassTxt = "$moduleName module has been installed."
$appxFailBtn = New-Object System.Management.Automation.Host.ChoiceDescription "&OK","OK"
$appxPassBtn = New-Object System.Management.Automation.Host.ChoiceDescription "&OK","OK"
$appxFailRes = $Host.UI.PromptForChoice("Message", $appxFailTxt, @($appxFailBtn), 0)
$appxPassRes = $Host.UI.PromptForChoice("Message", $appxPassTxt, @($appxPassBtn), 0)

Add-Type -AssemblyName System.Windows.Forms

$messageBoxInput = New-Object System.Windows.Forms.MessageBoxButtons
$messageBoxInput.YesNo

$messageBoxResult = [System.Windows.Forms.MessageBox]::Show('$moduleName module is not installed. Do you want to install it? *Required!', 'Confirmation', $messageBoxInput)

    if ($messageBoxResult -eq 'Yes') {
        # Install the module
        Install-Module $moduleName -Scope CurrentUser -Force
        $appxPassRes
        Write-Host "$moduleName module has been installed!" | Out-File -FilePath $appLogs -Append
    } 
    else {
    $appxFailRes
    Write-Host "$moduleName must be updated or installed before the Application can continue!" | Out-File -FilePath $appLogs -Append
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
       
      process {-
        echo ""
        $StopWatch = [system.diagnostics.stopwatch]::startnew()
        $Path = (Resolve-Path $Path).Path
        #Get Urls to download
        Write-Host -ForegroundColor Yellow "Processing $Uri"
        $WebResponse = Invoke-WebRequest -UseBasicParsing -Method 'POST' -Uri 'https://store.rg-adguard.net/api/GetFiles' -Body "type=url&url=$Uri&ring=Retail" -ContentType 'application/x-www-form-urlencoded'
        $LinksMatch = ($WebResponse.Links | where {$_ -like '*.appx*'} | where {$_ -like '*_neutral_*' -or $_ -like "*_"+$env:PROCESSOR_ARCHITECTURE.Replace("AMD","X").Replace("IA","X")+"_*"} | Select-String -Pattern '(?<=a href=").+(?=" r)').matches.value
        $Files = ($WebResponse.Links | where {$_ -like '*.appx*'} | where {$_ -like '*_neutral_*' -or $_ -like "*_"+$env:PROCESSOR_ARCHITECTURE.Replace("AMD","X").Replace("IA","X")+"_*"} | where {$_ } | Select-String -Pattern '(?<=noreferrer">).+(?=</a>)').matches.value
        #Create array of links and filenames
        $DownloadLinks = @()
        for($i = 0;$i -lt $LinksMatch.Count; $i++){
            $Array += ,@($LinksMatch[$i],$Files[$i])
        }
        #Sort by filename descending
        $Array = $Array | sort-object @{Expression={$_[1]}; Descending=$True}
        $LastFile = "temp123"
        for($i = 0;$i -lt $LinksMatch.Count; $i++){
            $CurrentFile = $Array[$i][1]
            $CurrentUrl = $Array[$i][0]
            #Find first number index of current and last processed filename
            if ($CurrentFile -match "(?<number>\d)"){}
            $FileIndex = $CurrentFile.indexof($Matches.number)
            if ($LastFile -match "(?<number>\d)"){}
            $LastFileIndex = $LastFile.indexof($Matches.number)
    
            #If current filename product not equal to last filename product

# Check if the app is installed
$package = Get-AppxPackage | Where-Object {$_.Name -eq $appName}

if ($package) {
    Write-Host "$appName is already installed."
} else {
    Write-Host "$appName is not installed. Do you want to install it? (Y/N)"
    $installResponse = Read-Host

    if ($installResponse -eq "Y" -or $installResponse -eq "y") {
        # Install the app
        Download-AppxPackage $appUrl $appDownload                
        if ($package) {
            Write-Host "$appName has been installed!"
        } else {
        Write-Host "$appName has failed to install!"
        }
    } else {
        Write-Host "You chose not to install $appName."
    }
}

# LOG ENDS!
Stop-Transcript
exit
