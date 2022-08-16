# the path needed to access the users OneDrive for Business directory
$company = Read-Host -Prompt "Please enter your company name as it appears on OneDrive fodler and press enter... (i.e Consco, Test Company)"
$path = "C:\Users\$env:UserName\OneDrive - $company"

#If there is no Signature folder copy it over directly
if ( -not ( Test-Path -Path $path\Signatures ) ){
  Copy-Item -Path $env:APPDATA\Microsoft\Signatures -Destination $path -recurse -Force
}
#If there is a Signature folder but not a .old folder, create .old and create fresh Copy
elseif ( -not ( Test-Path -Path $path\Signatures.old ) ) {
  Rename-Item -Path "$path\Signatures" "$path\Signatures.old"
  Copy-Item -Path $env:APPDATA\Microsoft\Signatures -Destination $path -recurse -Force
}
#If there is a signature folder and a .old, erase the .old, rename signature folder and create fresh Copy
else {
  Remove-Item -Path $path\Signatures.old -Recurse
  Rename-Item -Path "$path\Signatures" "$path\Signatures.old"
  Copy-Item -Path $env:APPDATA\Microsoft\Signatures -Destination $path -recurse -Force
}
