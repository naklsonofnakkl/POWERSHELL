## Close Adobe Acrobat
Read-Host -Prompt "Please Close Adobe First then Press any Key to Continue"
## get Acrobat process
$acrobat = Get-Process Acrobat -ErrorAction SilentlyContinue
if ($acrobat) {
  ## try gracefully first
  $acrobat.CloseMainWindow()
  ## kill after five seconds
  Sleep 5
  if (!$acrobat.HasExited) {
    $acrobat | Stop-Process -Force
  }
}
Remove-Variable acrobat
## Navigates to local AppData
set-location $env:LOCALAPPDATA\Adobe\Acrobat
## Removes previous OLD folder(s)
Remove-Item -Path .\*.old -Recurse -ErrorAction SilentlyContinue
## Clear the cache for both Adobe DC and XI
Write-Output "Clearing Cache..."
Rename-Item -Path .\"DC" .\"DC.old" -ErrorAction SilentlyContinue
Rename-Item -Path .\"XI" .\"XI.old" -ErrorAction SilentlyContinu
Write-Output "Cache Cleared!"
## Close script
Sleep 3
exit
