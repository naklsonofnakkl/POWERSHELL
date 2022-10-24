## THIS IS STILL AN ACTIVE WORK IN PROGRESS!
## PLEASE DO NOT ATTEMPT TO RUN THIS SCRIPT
## AS IT HAS NOT BEEN FINISHED AND MAY HAVE
## UNINTENTIONAL RESULTS!


Add-Type -AssemblyName PresentationFramework

# All of the Functions!
Function Clear-OutlookCache {
  #Directories I don't want to keep retyping in
    $outroam = "$env:APPDATA\Microsoft\Outlook"
    $outlocal = "$env:LOCALAPPDATA\Microsoft\Outlook"
    $oldroam = "$env:APPDATA\Microsoft\Outlook\OLD"
    $oldlocal = "$env:LOCALAPPDATA\Microsoft\Outlook\OLD"
    ## Close out of OUTLOOK
 Stop-Process -name OUTLOOK -force
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
}

Function Clear-TeamsCache {
  #Directories I don't want to keep retyping in
    $teamroam = "$env:APPDATA\Microsoft\Teams"
    $oldroam = "$env:APPDATA\Microsoft\Teams\OLD"
    ## Close out of TEAMS
   Stop-Process -name Teams -force
  #If there is no OLD folder create one and copy files into it
  #ROAMING
  if ( -not ( Test-Path -Path $oldroam ) ){
    New-Item -path $teamroam -name OLD -ItemType Directory
    set-location $teamroam
    $filedest = $oldroam
    $exclude = ".\meeting-addin", ".\OLD"
    $Files = GCI -path $teamroam | Where-object {$_.name -ne $exclude}
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
    $Files = GCI -path $teamroam | Where-object {$_.name -ne $exclude}
    foreach ($file in $files){move-item -path $file -destination $filedest -ErrorAction SilentlyContinue}
    }
  }

Function Clear-AdobeCache {
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
}

Function Install-GlobalProtect {
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
}

Function Install-PDF {
  Clear
  Write-Host "Installing the print-to-pdf feature..." -ForegroundColor Green
  Enable-WindowsOptionalFeature -Online -FeatureName "Printing-PrintToPDFServices-Features" -All
  $input = $(Write-Host "Please Press Enter Button to Restart..." -ForegroundColor Yellow -NoNewLine; Read-Host)
  Restart-Computer
  break
}

Function Remove-PDF {
  Clear
  Write-Host "Uninstalling the print-to-pdf feature..." -ForegroundColor Red
  Disable-WindowsOptionalFeature -Online -FeatureName "Printing-PrintToPDFServices-Features"
  $input = $(Write-Host "Please Press Enter Button to Restart..." -ForegroundColor Yellow -NoNewLine; Read-Host)
  Restart-Computer
  break
}

function Export-SigFolder {
  $signature = "$env:APPDATA\Microsoft\Signatures"
  $ms = "$env:APPDATA\Microsoft"
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

function Import-SigFolder {
  $signature = "$env:APPDATA\Microsoft\Signatures"
  $ms = "$env:APPDATA\Microsoft"
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

# ---------------------------------------------------------------

# where is the XAML file?
$xamlFile = "C:\Users\andrewwilson\OneDrive - Lennar Mortgage\Documents\GitHub\POWERSHELL\MISC\GUI\gui.xaml"

#create window
$inputXML = Get-Content $xamlFile -Raw
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXML

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}

# Create variables based on form control names.
# Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
Get-Variable var_*

$Null = $window.ShowDialog()
