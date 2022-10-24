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
## Closes Powershell
exit
