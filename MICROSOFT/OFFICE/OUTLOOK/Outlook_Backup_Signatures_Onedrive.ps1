<#
  .Export-SignatureBackup
  This will copy the Signature folder from AppData to the users OneDrive.
  It will also create a backup of a previous Signature folder in OneDrive if one
  is already present before copying over the newest version.

  .Import-RestoreSignature
  This will copy the Signature folder from OneDrive to the users AppData.
  It will also create a backup of the previous Signature folder in AppData if one
  is already present before copying over the backup from OneDrive.

#>

# the path needed to access the users OneDrive for Business directory
$company = Read-Host -Prompt "Please enter your company name as it appears on OneDrive fodler and press enter... (i.e Consco, Test Company)"
$path = "C:\Users\$env:UserName\OneDrive - $company"
$signature = "$env:APPDATA\Microsoft\Signatures"
$ms = "$env:APPDATA\Microsoft"

##This function will back up the users current signature folder into OneDrive (and keep one backup of any previous folder)
function Export-SignatureBackup {
#If there is no Signature folder copy it over directly
if ( -not ( Test-Path -Path $path\Signatures ) ){
  Copy-Item -Path "$signature" -Destination "$path" -recurse -Force
}
#If there is a Signature folder but not a .old folder, create .old and create fresh Copy
elseif ( -not ( Test-Path -Path $path\Signatures.old ) ) {
  Rename-Item -Path "$path\Signatures" "$path\Signatures.old"
  Copy-Item -Path "$signature" -Destination "$path" -recurse -Force
}
#If there is a signature folder and a .old, erase the .old, rename signature folder and create fresh Copy
else {
  Remove-Item -Path "$path\Signatures.old" -Recurse -Force
  Rename-Item -Path "$path\Signatures" "$path\Signatures.old"
  Copy-Item -Path "$signature" -Destination "$path" -recurse -Force
}
}

##This function will restore the users signature folder from OneDrive to AppData
function Import-RestoreSignature {
  ## Close out of OUTLOOK
  Stop-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
  #If there is no Signature folder in OneDrive, close script.
  if ( -not ( Test-Path -Path $path\Signatures ) ){
    Write-Output "You Don't have a signature folder in OneDrive, please run the backup function first!"
    Sleep 3
    exit
  }
  #turn the current Signature folder to .old and copy over the backup from OneDrive
  elseif ( -not ( Test-Path -Path $ms\Signatures.old ) ) {
    Rename-Item -Path "$ms\Signatures" "$ms\Signatures.old"
    Copy-Item -Path "$path\Signatures" -Destination $ms -recurse -Force
  }
  #If there is a signature folder and a .old, erase the .old, rename signature folder and create fresh Copy
  else {
    Remove-Item -Path "$ms\Signatures.old" -Recurse -Force
    Rename-Item -Path "$ms\Signatures" "$ms\Signatures.old"
    Copy-Item -Path "$path\Signatures" -Destination "$ms" -recurse -Force
  }
}
