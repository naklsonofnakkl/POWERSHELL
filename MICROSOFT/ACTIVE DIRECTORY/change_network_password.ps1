##This script exists to change the users password
##Updates Credentials for Global Protect VPN
##and then perform a Group Policy update

#Necessary moduel to interact with credential manager
#Setting PSGallery as trusted to install the module and back
#to untrusted for security reasons
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction -SilentlyContinue
Install-Module -Name CredentialManager -Scope CurrentUser -ErrorAction -SilentlyContinue
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Untrusted -ErrorAction -SilentlyContinue

#Prompt the user for username and new NewPassword
$user = Read-Host "Please enter your domain username [ex. JohnDoe, JDoe] and press enter..."
$newPass = Read-Host "Please enter your new Password [ex. 78!Flapy] and press enter..." -AsSecureString

#Change users password
if (Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$newPass" -Force) -ErrorAction SilentlyContinue)
{
#inform user password change complete
#change the global protect password stored in credential manager
#restart Global Protect VPN
#Update Group Policy with new Password
Write-Host "Password has been updated successfully!"
sleep 1
Write-Host 'Updating Global Protect Credentials...'
New-StoredCredential -Target gpcp/LatestCP -Username $user -Password $newPass
Write-Host 'Restarting Global Protect...'
Remove-Variable user
Remove-Variable newPass
Stop-Process -Name 'PanGPS' -ErrorAction SilentlyContinue
Stop-Process -Name 'PanGPA' -ErrorAction SilentlyContinue
start-process -FilePath "C:\Program Files\Palo Alto Networks\GlobalProtect\PanGPA.exe"
Read-Host -Prompt "Press enter to continue once Global Protect is Connected..."
sleep 1
Write-Host 'Updating Group Policy...'
Invoke-GPUpdate -Force
}
Else {
#Password changed failed. User needs to reach out to IT Support.
Remove-Variable user
Remove-Variable newPass
Write-Host "Password has FAILED to update!"
sleep 1
Write-Host "Please reach out to IT Support to have the password changed manually..."
Write-Host "Closing..."
sleep 5
exit
}
Write-Host "Password Change Process has been Completed Successfully! Exiting..."
sleep 3
exit
