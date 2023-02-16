## Close Encompass
Read-Host -Prompt "Please Close Encompass First then Press any Key to Continue"
## get Encompass process
$encompass = Get-Process encompass -ErrorAction SilentlyContinue
if ($encompass) {
  ## try gracefully first
  $encompass.CloseMainWindow()
  ## kill after five seconds
  Start-Sleep 5
  if (!$encompass.HasExited) {
    $encompass | Stop-Process -Force
  }
}
Remove-Variable encompass
# Register secman.dll and secman64.dll
$regsvr32 = "regsvr32.exe"
$regsvr32Args = "/s ""C:\Program Files (x86)\Common Files\Outlook Security Manager\secman.dll"""
Start-Process -FilePath $regsvr32 -ArgumentList $regsvr32Args -Wait

$regsvr32Args = "/s ""C:\Program Files (x86)\Common Files\Outlook Security Manager\secman64.dll"""
Start-Process -FilePath $regsvr32 -ArgumentList $regsvr32Args -Wait

# Display a message indicating that the registration is complete
Write-Host "DLLs registered successfully."
exit
