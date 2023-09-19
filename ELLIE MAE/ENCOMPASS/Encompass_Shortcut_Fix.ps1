<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Sometimes we need to change where shortcuts point to
    [CURRENTLY ONLY WORKS FOR WINDOWS 10]
.DESCRIPTION
    - Changes Encompass folder location

#>

# Define the keyword to search for in shortcut names
$keyword = "Encompass - Model Office"

# Get the user's desktop folder path
$desktopFolder = "C:\Users\Public\Desktop"

# Define the taskbar folder path (replace with the actual path if necessary)
$taskbarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

# Define the new target and "Start In" paths
$newTarget = "C:\SmartClientCache\Apps\Ellie Mae\Encompass\AppLauncher.exe"
$newStartIn = "C:\SmartClientCache\Apps\Ellie Mae\Encompass\"

# Function to modify shortcut files in a given folder
function Modify-Shortcuts($folderPath) {
    Get-ChildItem -Path $folderPath -Filter "*.lnk" | ForEach-Object {
        $shortcutFile = $_.FullName
        $shortcutName = [System.IO.Path]::GetFileNameWithoutExtension($shortcutFile)

        if ($shortcutName -like "*$keyword*") {
            # Create a Shell object to manipulate the shortcut
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($shortcutFile)

            # Modify the shortcut properties
            $shortcut.TargetPath = $newTarget
            $shortcut.WorkingDirectory = $newStartIn

            # Save the modified shortcut
            $shortcut.Save()
            Write-Host "Modified shortcut: $shortcutFile"
        }
    }
}

# Modify shortcuts on the desktop
Modify-Shortcuts -folderPath $desktopFolder

# Modify shortcuts on the taskbar
Modify-Shortcuts -folderPath $taskbarFolder