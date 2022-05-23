##This script exists to change the users password
##Updates Credentials for Global Protect VPN
##and then perform a Group Policy update

#Necessary moduel to interact with credential manager
#Setting PSGallery as trusted to install the module and back
#to untrusted for security reasons
import-module activedirectory
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
Install-Module -Name CredentialManager -Scope CurrentUser
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Untrusted

#Prompt the user for username and old password
$user = $env:UserName
$oldPassword = Read-Host -AsSecureString -Prompt "Please enter your CURRENT Password and press enter..."
$newPassword = Read-Host -AsSecureString -Prompt "Please enter your NEW Password and press enter..."

#Change local users password
Set-ADAccountPassword -Identity $userName -OldPassword $oldPassword -NewPassword $newPassword -ErrorVariable passFail

if ($passFail)
{
  #Password changed failed. User needs to reach out to IT Support.
  Remove-Variable user
  Remove-Variable oldPassword
  Remove-Variable newPassword
  Write-Host "Password has FAILED to update!"
  sleep 1
  Write-Host "Please reach out to IT Support to have the password changed manually..."
  Write-Host "Closing..."
  sleep 5
  exit
}
Else {
  #inform user password change complete
  #change the global protect password stored in credential manager
  #restart Global Protect VPN
  #Update Group Policy with new Password
  Write-Host "Password has been updated successfully!"
  sleep 1
  Write-Host 'Updating Global Protect Credentials...'
  New-StoredCredential -Target gpcp/LatestCP -Username $user -Password $newPassword
  Write-Host 'Restarting Global Protect...'
  Remove-Variable user
  Remove-Variable oldPassword
  Remove-Variable newPassword
  Stop-Process -Name 'PanGPS' -ErrorAction SilentlyContinue
  Stop-Process -Name 'PanGPA' -ErrorAction SilentlyContinue
  start-process -FilePath "C:\Program Files\Palo Alto Networks\GlobalProtect\PanGPA.exe"
  sleep 10
  Read-Host -Prompt "Press enter to continue once Global Protect is Connected..."
  sleep 1
  Write-Host 'Updating Group Policy...'
  Invoke-GPUpdate -Force
  Write-Host "Password Change Process has been Completed Successfully! Exiting..."
  sleep 3
  exit
}
#In the event of a permission error accessing Set-ADAccountPassword
Remove-Variable user
Remove-Variable oldPassword
Remove-Variable newPassword
Write-Host "Password has FAILED to update!"
sleep 1
Write-Host "Please reach out to IT Support to have the password changed manually..."
Write-Host "Closing..."
sleep 5
exit
