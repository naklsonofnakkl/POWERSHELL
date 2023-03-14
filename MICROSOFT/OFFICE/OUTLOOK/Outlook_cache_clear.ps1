##This script will close Outlook and clear the cache
##Located in the Local and Roaming appdata locations

#Directories I don't want to keep retyping in
  $outroam = "$env:APPDATA\Microsoft\Outlook"
  $outlocal = "$env:LOCALAPPDATA\Microsoft\Outlook"
  $oldroam = "$env:APPDATA\Microsoft\Outlook\OLD"
  ## $oldlocal = "$env:LOCALAPPDATA\Microsoft\Outlook\OLD"
  ## Close out of OUTLOOK
 Stop-Process -name OUTLOOK -force
 # Set the duration of the timer in seconds
$duration = 10

# Initialize the progress bar
Write-Progress -Activity "Waiting for $duration seconds while Outlook closes..." -PercentComplete 0

# Loop through the timer and update the progress bar
for ($i = 1; $i -le $duration; $i++) {
    # Update the progress bar with the current progress
    $percent = ($i / $duration) * 100
    Write-Progress -Activity "Waiting for $duration seconds while Outlook closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
    
    # Pause for 1 second
    Start-Sleep -Seconds 1
}

# Clear the progress bar once the timer is complete
Clear-Progress -ErrorAction SilentlyContinue
#If there is no OLD folder create one and copy files into it
#ROAMING
if ( -not ( Test-Path -Path $oldroam ) ){
  New-Item -path $outroam -name OLD -ItemType Directory
  Move-Item -Path $outroam\*.srs $outroam\OLD
  Move-Item -Path $outroam\*.xml $outroam\OLD
}
#If there is an OLD folder erase the OLD folder and create fresh Copy
#ROAMING
else {
  Remove-Item -Path "$outroam\OLD" -Recurse -Force
  New-Item -path $outroam -name OLD -ItemType Directory
  Move-Item -Path $outroam\*.srs $outroam\OLD
  Move-Item -Path $outroam\*.xml $outroam\OLD
}
#If there are no .old folders then rename all folders to end in .old
#LOCAL
if ( -not ( Test-Path -Path "$outlocal\RoamCache.old" ) ){
  Rename-Item -Path "$outlocal\RoamCache" "$outlocal\RoamCache.old"
  Rename-Item -Path "$outlocal\Offline Address Books" "$outlocal\Offline Address Books.old"
}
#If there are .old folders, delete them and convert current foldres into .old
#LOCAL
else {
  Remove-Item -Path $outlocal\*.old -Recurse
  Rename-Item -Path "$outlocal\RoamCache" "$outlocal\RoamCache.old"
  Rename-Item -Path "$outlocal\Offline Address Books" "$outlocal\Offline Address Books.old"
}
  # Prompt the user to reopen Outlook before closing the script
  $openApp = Read-Host "Do you want to open Microsoft Outlook? (Y/N)"

  # Check the user's response
  if ($openApp -eq "Y" -or $openApp -eq "y") {
      # Open Outlook
      Start-Process -FilePath "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
  }
  else {
    # Close powershell script
    exit
  }