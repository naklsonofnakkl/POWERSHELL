<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.6
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    This is a huge list of Excel functions necessary for various script I use in my bag of tricks
.DESCRIPTION
    - it is a lot of functions
    - Can't even justify it really, just need to call 
    - a lot of damn functions
#>

function Add-ModuleExcel {

    $moduleName = "ImportExcel"

    # Check if the module is installed
    $installedModule = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName }
    
    if ($installedModule) {
        Write-Host "The $moduleName module is installed."

        # execute the rest of the script 
        Join-JsonTable
        Format-SteamXlsx
        Clear-Installation
        Invoke-Item -Path $appDownloadPath
    }
    else {
        Write-Host "The $moduleName module is not installed."
        # Install the missing module and run the rest of the script
        Get-NuGet
        Join-JsonTable
        Format-SteamXlsx
        Clear-Installation
        Invoke-Item -Path $appDownloadPath
    }

} 

function Format-SteamXlsx {
    # Define the path to your Excel file
    $excelFilePath = "$env:TEMP\SteamLib_temp\steam_library_stats.xlsx"

    # Load the Excel data into a PowerShell object
    $condata = Import-Excel -Path $excelFilePath

    # Filter the data to only include rows where the fifth column contains "true"
    $condata = $condata | Where-Object { $_.Column5 -eq "TRUE" }

    # Export the filtered data to a new Excel file
    $filteredExcelFilePath = "$env:TEMP\SteamLib_temp\Steam_Multiplayer.xlsx"
    $condata | Export-Excel -Path $filteredExcelFilePath -NoHeader -AutoSize
}