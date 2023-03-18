<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.4
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    This is a huge list of Excel functions necessary for various script I use in my bag of tricks
.DESCRIPTION
    - it is a lot of functions
    - Can't even justify it really, just need to call 
    - a lot of damn functions
#>

function Format-SteamCsv {

# Install the ImportExcel module if it's not already installed
Install-Module -Name ImportExcel

# Load the module into the current session
Import-Module -Name ImportExcel

# Import CSV file
$data = Import-Csv -Path "$env:TEMP\SteamLib_temp\steam_library_stats.csv"

# Export data to XLSX file
$data | Export-Excel -Path "$env:TEMP\SteamLib_temp\steam_library_stats.xlsx" -AutoSize -AutoFilter

# Define the path to your Excel file
$excelFilePath = "$env:TEMP\SteamLib_temp\steam_library_stats.xlsx"

# Load the Excel data into a PowerShell object
$condata = Import-Excel -Path $excelFilePath

# Filter the data to only include rows where the fifth column contains "true"
$condata = $condata | Where-Object { $_.Column5 -eq "true" }

# Export the filtered data to a new Excel file
$filteredExcelFilePath = "$env:TEMP\SteamLib_temp\Steam_Multiplayer.xlsx"
$condata | Export-Excel -Path $filteredExcelFilePath -NoHeader -AutoSize
    } 