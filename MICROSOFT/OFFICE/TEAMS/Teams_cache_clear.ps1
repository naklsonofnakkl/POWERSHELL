#Directories I don't want to keep retyping in
  $teamroam = "$env:APPDATA\Microsoft\Teams"
  $oldroam = "$env:APPDATA\Microsoft\Teams\OLD"
  ## Close out of TEAMS
  $teams = Get-Process Teams -ErrorAction SilentlyContinue
  if ($teams) {
    ## try gracefully first
    $teams.CloseMainWindow()
    ## kill after five seconds
    Sleep 5
    if (!$teams.HasExited) {
      $teams | Stop-Process -Force
    }
  }
#If there is no OLD folder create one and copy files into it
#ROAMING
if ( -not ( Test-Path -Path $oldroam ) ){
  New-Item -path $teamroam -name OLD -ItemType Directory
  $filedest = "$teamroam\OLD"
  $exclude = "meeting-addin"
  $Files = GCI -path $env:APPDATA\Microsoft\Teams | Where-object {$_.name -ne $exclude}
  foreach ($file in $files){move-item -path $file -destination $filedest}
}
#If there is an OLD folder erase the OLD folder and create fresh Copy
#ROAMING
else {
  Remove-Item -Path "$teamroam\OLD" -Recurse -Force
  New-Item -path $teamroam -name OLD -ItemType Directory
  $filedest = "$teamroam\OLD"
  $exclude = "meeting-addin"
  $Files = GCI -path $env:APPDATA\Microsoft\Teams | Where-object {$_.name -ne $exclude}
  foreach ($file in $files){move-item -path $file -destination $filedest}
  }
## Closes Powershell
exit
