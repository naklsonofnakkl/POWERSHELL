## close teams
Stop-Process -Name 'Teams'
## change to the teams cache folder
set-location $env:APPDATA\Microsoft\Teams
## Remove the previous OLD directory
Remove-Item OLD -Recurse
New-Item OLD -ItemType Directory
## Moves all files into OLD direction except Outlook Add-on
$filedest = ".\OLD"
$exclude = "meeting-addin"
$Files = GCI -path $env:APPDATA\Microsoft\Teams | Where-object {$_.name -ne $exclude}
foreach ($file in $files){move-item -path $file -destination $filedest}
## Closes Powershell
exit
