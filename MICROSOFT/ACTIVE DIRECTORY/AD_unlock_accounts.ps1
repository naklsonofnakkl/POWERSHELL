<#
.NOTES
    Author: Andrew Wilson
    Version: 0.1.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Unlock account in AD for all available Domains
.DESCRIPTION
    - Prompts user for SAM Username
    - Finds all domains on AD
    - Attempts to unlock user accounts

#>

$Username = Read-Host "Enter the SAM Username to unlock"
# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\AD_unlock_accounts.log"
$ErrorActionPreference = "Stop"

try {
    foreach ($w in (Get-ADDomainController -Filter *).Name) {
        Unlock-ADAccount $Username -Server $w -ErrorAction SilentlyContinue
        Get-ADUser $Username -Properties LockedOut |
        Select-Object Name, LockedOut, @{Name = 'Server'; Expression = { $w } }
        Start-Sleep 5
        $User = Get-ADUser $Username -Properties LockedOut
        $Status = if ($User.LockedOut) { "Locked" } else { "Unlocked" }
        Write-Host "$w $Status" | Out-File -FilePath $appLogs -Append
    }
    
}
catch {
    if ($_.Exception.Message -match "Insufficient access rights to perform the operation") {
        $_.CategoryInfo | Out-File -FilePath $appLogs -Append
        Write-Host "You Do Not Have Permission to Unlock Accounts on this Domain!" -ForegroundColor Red
    }
    else {
        $_.CategoryInfo | Out-File -FilePath $appLogs -Append
        Write-Host "There was an error processing the request!" -ForegroundColor Red
        Write-Host "Press any key to open the log: $appLogs" -ForegroundColor Green
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Start-Process $appLogs
    }
}
