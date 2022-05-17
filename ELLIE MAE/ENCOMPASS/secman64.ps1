## Close Encompass
Read-Host -Prompt "Please Close Encompass First then Press any Key to Continue"
## get Encompass process
$encompass = Get-Process encompass -ErrorAction SilentlyContinue
if ($encompass) {
  ## try gracefully first
  $encompass.CloseMainWindow()
  ## kill after five seconds
  Sleep 5
  if (!$encompass.HasExited) {
    $encompass | Stop-Process -Force
  }
}
Remove-Variable encompass
## Change to Outlook Security Manager folders
set-location "C:\Program Files (x86)\Common Files\Outlook Security Manager"
## Register the DLL
Read-Host -Prompt "Press any key to install secman.dll"
Start-Process regsvr32.exe secman.dll
Read-Host -Prompt "Press any key to install secman64.dll"
Start-Process regsvr32.exe secman64.dll
write-host "DLL have been properly installed!"
sleep 3
exit
