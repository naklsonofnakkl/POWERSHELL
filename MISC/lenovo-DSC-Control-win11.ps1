<#
This is a script to install the Lenovo update for
flicking monitor issues on the Windows 11 OS.
#>

##Download the update file
Invoke-WebRequest -Uri "https://download.lenovo.com/km/media/attachment/DSC_Control.exe" -OutFile C:\Users\$env:UserName\Downloads\DSC_Control.exe
##Chcek if the file has been downloaded and execute DSC_Control
if ( -not ( Test-Path -Path $Change\DSC_Control.exe ) ){
  set-location "C:\Users\$env:UserName\Downloads"
  .\DSC_Control.exe 1
}
##If the file doesn't download immediately, wait a min and execute DSC_Control
else {
  Sleep 60
  set-location "C:\Users\$env:UserName\Downloads"
  .\DSC_Control.exe 1
}
##Remove the download
remove-item C:\Users\$env:UserName\Downloads\DSC_Control.exe
exit
