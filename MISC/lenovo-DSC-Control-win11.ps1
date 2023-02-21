<#
This is a script to install the Lenovo update for
flicking monitor issues on the Windows 11 OS.
#>

$url = "https://download.lenovo.com/km/media/attachment/DSC_Control.exe"
$filePath = "C:\Users\$env:UserName\Downloads\DSC_Control.exe"
$executablePath = "C:\Users\$env:UserName\Downloads\DSC_Control.exe"

# Download the file and show a progress bar
Invoke-WebRequest -Uri $url -OutFile $filePath -UseBasicParsing -TimeoutSec 120 -Verbose

# Validate that the file has been downloaded
if (Test-Path $filePath) {
    # Execute the downloaded file
    Start-Process $executablePath

    # Delete the downloaded file
    Remove-Item $filePath
} else {
    Write-Output "Download failed, file not found: $filePath"
}
exit
