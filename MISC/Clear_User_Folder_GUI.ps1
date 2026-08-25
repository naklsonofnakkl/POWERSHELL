## VARIABLE PARADISE
$currentUser = Read-Host -Prompt "Please enter the AD Username of any profiles to skip (jsmith1234)"
$adminAccounts = Get-CimInstance Win32_UserProfile |
Where-Object {
    $_.LocalPath -match '^C:\\Users\\' -and
    $_.LocalPath -notmatch '(admin|srv-|a-)'
} |
Select-Object -ExpandProperty LocalPath
$blockedUsers = @(
    $currentUser,
    'Default',
    'Default User',
    'Public',
    'Administrator',
    'ADMINI~1',
    'cached',
    'ntwadmin'
)
$tempDir = $env:TEMP
$reportPath = Join-Path $tempDir "DeletedFilesReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

## FUNCTION JUNCTION
function Remove-UserFolders {
    param (
        $blockedUsers,
        $currentUser,
        $adminAccounts,
        $reportPath
    )

    $cutoff = (Get-Date).AddMonths(0)
    $deletedItemsLog = [System.Collections.Generic.List[PSCustomObject]]::new()

    $targetProfiles = $adminAccounts |
    Get-Item |
    Where-Object {
        $blockedUsers -notcontains $_.Name -and
        $_.LastWriteTime -lt $cutoff
    }

    if (-not $targetProfiles) {
        Write-Host "No qualifying profiles found to clean." -ForegroundColor Yellow
        return
    }

    Clear-Host
    [Console]::CursorVisible = $false

    try {
        foreach ($profile in $targetProfiles) {
            $size = (Get-ChildItem $profile.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
            $sizeGB = "{0:N2} GB" -f ($size / 1GB)
            $profileName = $profile.Name
            $lastModified = $profile.LastWriteTime

            [Console]::SetCursorPosition(0, 0)
            Write-Host "==================================================" -ForegroundColor Cyan
            Write-Host "               PROFILE CLEANUP DASHBOARD          " -ForegroundColor Cyan
            Write-Host "==================================================" -ForegroundColor Cyan
            Write-Host " Name          : $($profileName.PadRight(35))" -ForegroundColor Yellow
            Write-Host " Profile Size  : $($sizeGB.PadRight(35))" -ForegroundColor Yellow
            Write-Host " Last Modified : $($lastModified.ToString().PadRight(35))" -ForegroundColor Yellow
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
            Write-Host " OUTPUT:                                          " -ForegroundColor DarkCyan
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan

            try {
                $items = [System.IO.Directory]::GetFileSystemEntries($profile.FullName) |
                ForEach-Object { Get-Item -LiteralPath $_ -Force }
            }
            catch {
                Write-Host " [WARNING] Unable to enumerate $($profile.FullName)" -ForegroundColor Red
                continue
            }

            if (@($items).Count -eq 0) {
                Write-Host " [INFO] Profile folder is already empty." -ForegroundColor DarkGray

                try {
                    Get-CimInstance Win32_UserProfile |
                    Where-Object LocalPath -eq $profile.FullName |
                    Remove-CimInstance -ErrorAction Stop
                    Remove-Item -Path $profile.FullName -Force -Recurse -ErrorAction Stop
                    Write-Host " [SUCCESS] Removed empty profile folder." -ForegroundColor Green
                }
                catch {
                    Write-Host " [WARNING] Failed to delete $($profile.FullName): $($_.Exception.Message)" -ForegroundColor Red
                }
                Start-Sleep -Seconds 1
                continue
            }

            foreach ($item in $items) {
                try {
                    $subItems = Get-ChildItem $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    $targetCollection = @($item) + @($subItems)

                    Remove-Item $item.FullName -Recurse -Force -ErrorAction Stop

                    foreach ($sub in $targetCollection) {
                        $deletedItemsLog.Add([PSCustomObject]@{
                                User         = $profileName
                                Path         = $sub.FullName
                                Name         = $sub.Name
                                ItemType     = if ($sub.PSIsContainer) { "Directory" } else { "File" }
                                DeletionTime = (Get-Date)
                                Status       = "Success"
                            })
                    }
                    Write-Host " [SUCCESS] Removed: $($item.Name)" -ForegroundColor Green
                }
                catch {
                    Write-Host " [WARNING] Failed: $($item.Name)" -ForegroundColor Red
                    $deletedItemsLog.Add([PSCustomObject]@{
                            User         = $profileName
                            Path         = $item.FullName
                            Name         = $item.Name
                            ItemType     = if ($item.PSIsContainer) { "Directory" } else { "File" }
                            DeletionTime = (Get-Date)
                            Status       = "Failed: $($_.Exception.Message)"
                        })
                }
            }
            Start-Sleep -Milliseconds 800
        }
    }
    finally {
        [Console]::CursorVisible = $true
    }

    if ($deletedItemsLog.Count -gt 0) {
        Write-Host "`n"
        $deletedItemsLog | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Deleted files report exported to: $reportPath" -ForegroundColor Cyan
        Start-Process $tempDir
        Write-Host "Cleanup Complete!" -ForegroundColor Green
    }
    else {
        Write-Host "`nNo files were deleted, skipping report export." -ForegroundColor DarkYellow
    }
}

try {
    Remove-UserFolders -blockedUsers $blockedUsers -currentUser $currentUser -adminAccounts $adminAccounts -reportPath $reportPath
}
catch {
    [Console]::CursorVisible = $true
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
