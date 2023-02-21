# Check if ActiveX controls are enabled
$activexEnabled = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0" -Name "1001" -ErrorAction SilentlyContinue).1001 -eq 0

# If ActiveX controls are not enabled, create the registry value to enable them
if (-not $activexEnabled) {
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0" -Name "1001" -Value "0" -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0" -Name "1200" -Value "0" -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\0" -Name "1400" -Value "0" -PropertyType DWORD -Force | Out-Null
}