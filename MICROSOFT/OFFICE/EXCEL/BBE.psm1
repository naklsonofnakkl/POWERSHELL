<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.1
    
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
    $outputFilePath = "$env:TEMP\SteamLib_temp\$yourName.csv"

    # Define the $yourName varible via a text prompt
    Add-Type -AssemblyName System.Windows.Forms

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Enter Your Username"
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Form.AutoSize = $true
    $Form.AutoSizeMode = "GrowAndShrink"

    # Create the message label
    $MessageLabel = New-Object System.Windows.Forms.Label
    $MessageLabel.Location = New-Object System.Drawing.Point(10, 10)
    $MessageLabel.Size = New-Object System.Drawing.Size(280, 20)
    $MessageLabel.Text = "Please enter your Online Alias:"
    $Form.Controls.Add($MessageLabel)

    # Create the API Key textbox
    $yourNameTextBox = New-Object System.Windows.Forms.TextBox
    $yourNameTextBox.Location = New-Object System.Drawing.Point(10, 30)
    $yourNameTextBox.Size = New-Object System.Drawing.Size(280, 20)
    $Form.Controls.Add($yourNameTextBox)

    # Create the "Submit" button
    $SubButton = New-Object System.Windows.Forms.Button
    $SubButton.Location = New-Object System.Drawing.Point(115, 70)
    $SubButton.Size = New-Object System.Drawing.Size(75, 23)
    $SubButton.Text = "Submit"
    $SubButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $SubButton
    $Form.Controls.Add($SubButton)

    # Show the form and wait for a result
    $Result = $Form.ShowDialog()

    # If the "submit" button was clicked and the API Key is 32 characters long, set the $yourName variable
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $yourName = $yourNameTextBox.Text

        # Import the CSV file
        $data = Import-Csv $csvFilePath

        #Filter the data to only include rows with "true" in the 5th column
        $data = $data | Where-Object { $_.Column5 -eq "true" }

        #Export the updated data to a new CSV file
        $data | Export-Csv $outputFilePath -NoTypeInformation
    } else { 
        exit
    }
}