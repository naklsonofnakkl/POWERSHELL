<#
.NOTES
    Author: Andrew Wilson
    Version: 1.3.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Clear the cache for Microsoft Teams
.DESCRIPTION
    - Checks if Teams or dependant programs are running and closes if necessary
    - Clears the temporary cache folder
    - Starts Teams back up after successful cache clear
    
#>

<#
--------------------
 VARIBALE PARADISE!
--------------------
#>

#Directories
$tempDir = $env:TEMP
$teamsCache = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"
$folders = Get-ChildItem -Directory $teamsCache | Where-Object { $_.PSIsContainer } | Foreach-Object { $_.Name }

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$appLogs = "$tempDir\MSTEAMS_Cache.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

# Function to clean up the leftover downloaded files
function Clear-Installation {
  Stop-Transcript
}

#Function to automatically close Microsoft Teams
function Remove-ProcessTree {
    $officeApps = @("ms-teams", "ms-teams-updater", "ms-teams-relauncher")
    $procs = Get-CimInstance Win32_Process |
    Select-Object Name, ProcessId, ParentProcessId
    $teams = $procs | Where-Object { $_.Name -match '^ms-teams|teams' }
    if (-not $teams) {
        Write-Warning "Teams is not running."
        return
    }
    $primaryWebView = $procs |
    Where-Object {
        $_.Name -eq 'msedgewebview2.exe' -and
        ($teams.ProcessId -contains $_.ParentProcessId)
    }

    if (-not $primaryWebView) {
        Write-Warning "No msedgewebview2.exe child found under Teams."
        return
    }
    $primaryPID = $primaryWebView[0].ProcessId
    $secondaryWebViews = $procs |
    Where-Object {
        $_.Name -eq 'msedgewebview2.exe' -and
        $_.ParentProcessId -eq $primaryPID -and
        $_.ProcessId -ne $primaryPID
    }
    foreach ($p in $secondaryWebViews) {
        try {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
            Write-Output "Stopped msedgewebview2.exe PID $($p.ProcessId) (child of $primaryPID)"
        }
        catch {
            Write-Warning "Failed to stop PID $($p.ProcessId): $_"
        }
    }
    foreach ($app in $officeApps) {
        try {
            Stop-Process -Name $app -Force -ErrorAction SilentlyContinue
            Write-Host "Application Closed: $app"
        }
        catch {
            Write-Host "Failed to close application: $app"
        }
    }

    # Output summary
    [PSCustomObject]@{
        TeamsPID           = $teams.ProcessId
        PrimaryWebView2PID = $primaryPID
        RemovedChildPIDs   = $secondaryWebViews.ProcessId
    }
    Start-Sleep -Seconds 2 
}

Function Clear-TeamsCache {
    foreach ($folder in $folders) {
        try {
            if (Test-Path $teamsCache) {
                Get-ChildItem -Path $teamsCache | Remove-Item -Confirm:$false -Recurse -Force
            }
        }
        catch [System.Exception] {
            Write-Error $_.Exception.Message
        }
    }
    Start-Sleep -Seconds 2
    Start-Process ms-teams
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Remove-ProcessTree
Clear-TeamsCache
Stop-Transcript
exit
