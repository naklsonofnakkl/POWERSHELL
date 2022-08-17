<#
 .Synopsis
  Backup Outlook Signature File to OneDrive for Business

 .Description
  Asks the user to provide the name of their company as found on their OneDrive
  for Business folder and then will either Export the Signature folder to it or
  Restore the signature folder from it to the local users AppData.

  .Example
  # This will copy the Signature folder from AppData to the users OneDrive.
  # It will also create a backup of a previous Signature folder in OneDrive if one
  # is already present before copying over the newest version.
  Export-SigFolder

  .Example
  # This will copy the Signature folder from OneDrive to the users AppData.
  # It will also create a backup of the previous Signature folder in AppData if one
  # is already present before copying over the backup from OneDrive.
  Import-SigFolder
#>

# the paths needed to access the users OneDrive for Business directory
$signature = "$env:APPDATA\Microsoft\Signatures"
$ms = "$env:APPDATA\Microsoft"

##This function will back up the users current signature folder into OneDrive (and keep one backup of any previous folder)
function Export-SigFolder {
  $company = Read-Host -Prompt "Please enter your company name EXACTLY as it appears on OneDrive folder and press enter... (i.e Consco, Test Company)"
  $path = "C:\Users\$env:UserName\OneDrive - $company"
#If there is no Signature folder copy it over directly
if ( -not ( Test-Path -Path $path\Signatures ) ){
  Copy-Item -Path "$signature" -Destination "$path" -recurse -Force
  Sleep 3
  exit
}
#If there is a Signature folder but not a .old folder, create .old and create fresh Copy
elseif ( -not ( Test-Path -Path $path\Signatures.old ) ) {
  Rename-Item -Path "$path\Signatures" "$path\Signatures.old"
  Copy-Item -Path "$signature" -Destination "$path" -recurse -Force
  Sleep 3
  exit
}
#If there is a signature folder and a .old, erase the .old, rename signature folder and create fresh Copy
else {
  Remove-Item -Path "$path\Signatures.old" -Recurse -Force
  Rename-Item -Path "$path\Signatures" "$path\Signatures.old"
  Copy-Item -Path "$signature" -Destination "$path" -recurse -Force
  Sleep 3
  exit
}
}

##This function will restore the users signature folder from OneDrive to AppData
function Import-SigFolder {
  $company = Read-Host -Prompt "Please enter your company name EXACTLY as it appears on OneDrive folder and press enter... (i.e Consco, Test Company)"
  $path = "C:\Users\$env:UserName\OneDrive - $company"
  ## Close out of OUTLOOK
  Stop-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
  #If there is no Signature folder in OneDrive, close script.
  if ( -not ( Test-Path -Path $path\Signatures ) ){
    Write-Output "You Don't have a signature folder in OneDrive, please run Export-SigFolder first!"
    Sleep 3
    exit
  }
  #turn the current Signature folder to .old and copy over the backup from OneDrive
  elseif ( -not ( Test-Path -Path $ms\Signatures.old ) ) {
    Rename-Item -Path "$ms\Signatures" "$ms\Signatures.old"
    Copy-Item -Path "$path\Signatures" -Destination $ms -recurse -Force
    Sleep 3
    exit
  }
  #If there is a signature folder and a .old, erase the .old, rename signature folder and create fresh Copy
  else {
    Remove-Item -Path "$ms\Signatures.old" -Recurse -Force
    Rename-Item -Path "$ms\Signatures" "$ms\Signatures.old"
    Copy-Item -Path "$path\Signatures" -Destination "$ms" -recurse -Force
    Sleep 3
    exit
  }
}
