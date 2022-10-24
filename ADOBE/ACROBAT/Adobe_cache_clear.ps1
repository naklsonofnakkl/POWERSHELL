#Directories I don't want to keep retyping in
  $adobelocal = "$env:LOCALAPPDATA\Adobe\Acrobat"
  $oldDC = "$env:LOCALAPPDATA\Adobe\Acrobat\DC.old"
  $oldXI = "$env:LOCALAPPDATA\Adobe\Acrobat\XI.old"
  $adobeDC = "$env:LOCALAPPDATA\Adobe\Acrobat\DC"
  $adobeXI = "$env:LOCALAPPDATA\Adobe\Acrobat\XI"
  ## Close out of Adobe
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
