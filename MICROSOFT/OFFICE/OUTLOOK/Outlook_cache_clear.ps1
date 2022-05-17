## Close out of OUTLOOK
Stop-Process -Name 'OUTLOOK'
## Opens AppData, removes previous OLD folder and replaces it with new one
set-location $env:APPDATA\Microsoft\Outlook
Remove-Item OLD -Recurse -ErrorAction SilentlyContinue
New-Item OLD -ItemType Directory
## Moves cache files into OLD folder
Move-Item -Path .\*.srs .\OLD
Move-Item -Path .\*.xml .\OLD
## Navigates to local AppData
set-location $env:LOCALAPPDATA\Microsoft\Outlook
## Removes previous OLD folder(s)
Remove-Item -Path .\*.old -Recurse
Remove-Item -Path .\OLD -Recurse
## Renames the cache folders as .old
Rename-Item -Path .\"RoamCache" .\"RoamCache.old"
Rename-Item -Path .\"Offline Address Books" .\"Offline Address Books.old"
## Closes Powershell
exit
