#Directories I don't want to keep retyping in
  $adobelocal = "$env:LOCALAPPDATA\Adobe\Acrobat"
  $oldDC = "$env:LOCALAPPDATA\Adobe\Acrobat\DC.old"
  $oldXI = "$env:LOCALAPPDATA\Adobe\Acrobat\XI.old"
  $adobeDC = "$env:LOCALAPPDATA\Adobe\Acrobat\DC"
  $adobeXI = "$env:LOCALAPPDATA\Adobe\Acrobat\XI"
  ## Close out of Adobe
  Stop-Process -name Acrobat -force 
  # Set the duration of the timer in seconds
  $duration = 10
  
  # Initialize the progress bar
  Write-Progress -Activity "Waiting for $duration seconds while Adobe closes..." -PercentComplete 0
  
  # Loop through the timer and update the progress bar
  for ($i = 1; $i -le $duration; $i++) {
      # Update the progress bar with the current progress
      $percent = ($i / $duration) * 100
      Write-Progress -Activity "Waiting for $duration seconds while Adobe closes..." -PercentComplete $percent -Status "Seconds remaining: $($duration - $i)"
      
      # Pause for 1 second
      Start-Sleep -Seconds 1
  }
  
  # Clear the progress bar once the timer is complete
  Clear-Progress -ErrorAction SilentlyContinue

#If there is no DC folder create one and copy files into it
#LOCAL
if ( -not ( Test-Path -Path $adobeDC ) ){
  if (Test-Path -Path $oldXI) {
    Remove-Item -Path "$adobelocal\*.old" -Recurse -ErrorAction SilentlyContinue
    Rename-Item -Path "$adobelocal\XI" "$adobelocal\XI.old" -ErrorAction SilentlyContinue
  }
  else {
    Rename-Item -Path "$adobelocal\XI" "$adobelocal\XI.old" -ErrorAction SilentlyContinue
  }
}
elseif ( -not ( Test-Path -Path $adobeXI ) ) {
  if (Test-Path -Path $oldDC) {
    Remove-Item -Path "$adobelocal\*.old" -Recurse -ErrorAction SilentlyContinue
    Rename-Item -Path "$adobelocal\DC" "$adobelocal\DC.old" -ErrorAction SilentlyContinue
  }
  else {
    Rename-Item -Path "$adobelocal\DC" "$adobelocal\DC.old" -ErrorAction SilentlyContinue
  }
}
else {
  Try {
Remove-Item -Path "$adobelocal\*" -Recurse -ErrorAction SilentlyContinue
  }
  catch {
    write-host "There is nothing here! Check if Adobe is currently installed!"
  }
}
