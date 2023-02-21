## This script will alter the registry in order to put back the OUTLOOK
## Navigation bar to the bottom of the screen as it was previously

$regPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentEcs\Overrides'
$regProperty = 'Microsoft.Office.Outlook.Hub.HubBar'
$property = (Get-ItemProperty -Path $regPath).Microsoft.Office.Outlook.Hub.HubBar

Push-Location
Set-Location $regPath
Set-ItemProperty . Microsoft.Office.Outlook.Hub.HubBar "False"

Write-Host "Currently the value is set to..."
$null -ne $property
Write-Host "Closing script..."
Pop-Location
Start-Sleep 10
