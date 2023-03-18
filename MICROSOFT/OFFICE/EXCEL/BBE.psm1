<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.2
    
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

    #set the output paths
    $csvFilePath = "$env:TEMP\SteamLib_temp\steam_library_stats.csv"
    $outputFilePath = "$env:TEMP\SteamLib_temp\Steam_Multiplayer.csv"

        # Import the CSV file
        $data = Import-Csv $csvFilePath

        #Filter the data to only include rows with "true" in the 5th column
        $data = $data | Where-Object { $_.Column5 -eq "true" }

        #Export the updated data to a new CSV file
        $data | Export-Csv $outputFilePath -NoTypeInformation
    } else { 
        exit
    }