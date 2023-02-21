## This script will install the secman and secman64 DLL files

$secmanPath = "C:\Program Files (x86)\Common Files\Outlook Security Manager\"
## Close Encompass
do {
  $response = Read-Host "Is it safe to close Encompass? (y/n)"
  if ($response -eq "y") {
      Stop-Process -Name "encompass" -Force
      break
  }
  elseif ($response -eq "n") {
      for ($i = 10; $i -gt 0; $i--) {
          Write-Host "Closing in $i seconds" -ForegroundColor Yellow
          Start-Sleep -Seconds 1
      }
  }
  else {
      Write-Host "Invalid response. Please enter 'y' or 'n'" -ForegroundColor Red
  }
} while ($true)
# Register secman.dll and secman64.dll
$regsvr32 = "regsvr32.exe"
$regsvr32Args = "/s ""C:\Program Files (x86)\Common Files\Outlook Security Manager\secman.dll"""
Start-Process -FilePath $regsvr32 -ArgumentList $regsvr32Args -Wait

$regsvr32Args = "/s ""C:\Program Files (x86)\Common Files\Outlook Security Manager\secman64.dll"""
Start-Process -FilePath $regsvr32 -ArgumentList $regsvr32Args -Wait

# Begin check to validate DLL were installed
Write-Host "Validating if the DLL files were successfully installed..."
# Check if secman.dll exists
if (Test-Path ($secmanPath + "secman.dll")) {
  Write-Host "secman.dll is installed!" -ForegroundColor Green
} else {
  Write-Host "secman.dll is missing!" -ForegroundColor Red
}

# Check if secman64.dll exists
if (Test-Path ($secmanPath + "secman64.dll")) {
  Write-Host "secman64.dll is installed!" -ForegroundColor Green
} else {
  Write-Host "secman64.dll is missing!" -ForegroundColor Red
}
# close the script after user input
Read-Host -Prompt "Press any key to close..."
exit
