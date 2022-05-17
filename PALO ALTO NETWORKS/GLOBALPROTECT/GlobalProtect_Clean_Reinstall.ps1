##This Script will uninstall and reinstall Global Protect
##PLEASE MAKE SURE AN ACTIVE INTERNET CONNECTION IS AVAILABLE

#Have user put in their company gateway
$GATEWAY = Read-Host "Please Enter your GlobalProtect Gateway address [ex. gp.CONTOSO.com]"
#Check to confirm GP is installed
if ($GP = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "GlobalProtect"})
{
#Uninstall GP
write-host "Uninstalling Global Protect..."
$GP.uninstall() /quiet
write-host "Uninstalled!"
sleep 1
#Clear GP cache
write-host "Clearing Global Protect Cache..."
set-location "$env:LOCALAPPDATA\Palo Alto Networks\GlobalProtect"
Remove-Item OLD -Recurse -ErrorAction -SilentlyContinue
New-Item OLD -ItemType Directory
Move-Item -Path .\*.dat .\OLD
Move-Item -Path .\*.pan .\OLD
write-host "Global Protect Cache Cleared!"
sleep 1
#Install GP
write-host "Downloading Global Protect..."
Invoke-WebRequest -Uri "https://$GATEWAY/global-protect/getmsi.esp?version=64&platform=windows" -OutFile $home\Downloads\GlobalProtect64.msi -UseDefaultCredentials
write-host "Download Complete!"
write-host "Reinstalling Global Protect..."
Start-Process $home\Downloads\GlobalProtect64.msi -ArgumentList "/quiet /passive"
write-host "Global Protect has been Reinstalled Successfully!"
Remove-Item "$home\Downloads\GlobalProtect64.msi"
sleep 3
exit
}
Else {
  Write-Host "Global Protect is Currently Not Installed!"
  sleep 1
  #Install GP
  write-host "Downloading Global Protect..."
  Invoke-WebRequest -Uri "https://$GATEWAY/global-protect/getmsi.esp?version=64&platform=windows" -OutFile $home\Downloads\GlobalProtect64.msi -UseDefaultCredentials
  write-host "Download Complete!"
  write-host "Installing Global Protect..."
  Start-Process $home\Downloads\GlobalProtect64.msi -ArgumentList "/quiet /passive"
  write-host "Global Protect has been Reinstalled Successfully!"
  Remove-Item "$home\Downloads\GlobalProtect64.msi"
  sleep 3
  exit
}
