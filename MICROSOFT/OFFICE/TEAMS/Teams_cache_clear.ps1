#Directories I don't want to keep retyping in
  $teamroam = "$env:APPDATA\Microsoft\Teams"
  $oldroam = "$env:APPDATA\Microsoft\Teams\OLD"
  ## Close out of TEAMS
 Stop-Process -name Teams -force
# Set the duration of the timer in seconds
$duration = 10

# Initialize the progress bar
Write-Progress -Activity "Waiting for $duration seconds while Teams closes..." -PercentComplete 0

# Loop through the timer and update the progress bar
for ($i = 1; $i -le $duration; $i++) {
    # Update the progress bar with the current progress
    $percent = ($i / $duration) * 100
    Write-Progress -Activity "Waiting for $duration seconds while Teams closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
    
    # Pause for 1 second
    Start-Sleep -Seconds 1
}

# Clear the progress bar once the timer is complete
Clear-Progress -ErrorAction SilentlyContinue

#If there is no OLD folder create one and copy files into it
#ROAMING
if ( -not ( Test-Path -Path $oldroam ) ){
  New-Item -path $teamroam -name OLD -ItemType Directory
  set-location $teamroam
  $filedest = $oldroam
  $exclude = ".\meeting-addin", ".\OLD"
  $Files = Get-ChildItem -path $teamroam | Where-object {$_.name -ne $exclude}
  foreach ($file in $files){move-item -path $file -destination $filedest -ErrorAction SilentlyContinue}
}
#If there is an OLD folder erase the OLD folder and create fresh Copy
#ROAMING
else {
  Remove-Item -Path "$teamroam\OLD" -Recurse -Force
  New-Item -path $teamroam -name OLD -ItemType Directory
  set-location $teamroam
  $filedest = $oldroam
  $exclude = ".\meeting-addin", ".\OLD"
  $Files = Get-ChildItem -path $teamroam | Where-object {$_.name -ne $exclude}
  foreach ($file in $files){move-item -path $file -destination $filedest -ErrorAction SilentlyContinue}
  }

  # Prompt the user to reopen teams before closing the script
$openApp = Read-Host "Do you want to open Microsoft Teams? (Y/N)"

# Check the user's response
if ($openApp -eq "Y" -or $openApp -eq "y") {
    # Launch Microsoft Teams with the target command
Start-Process "C:\Users\$env:USERNAME\AppData\Local\Microsoft\Teams\Update.exe" `
-ArgumentList "--processStart", "Teams.exe", "--process-start-args", "--profile=AAD"
}
else {
  # Close powershell script
  exit
}