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

    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Enter Your Username"
    $subForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $subForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $subForm.AutoSize = $true
    $subForm.AutoSizeMode = "GrowAndShrink"

    # Create the message label
    $subMessageLabel = New-Object System.Windows.Forms.Label
    $subMessageLabel.Location = New-Object System.Drawing.Point(10, 10)
    $subMessageLabel.Size = New-Object System.Drawing.Size(280, 20)
    $subMessageLabel.Text = "Please enter your Online Alias:"
    $subForm.Controls.Add($subMessageLabel)

    # Create the API Key textbox
    $yourNameTextBox = New-Object System.Windows.Forms.TextBox
    $yourNameTextBox.Location = New-Object System.Drawing.Point(10, 30)
    $yourNameTextBox.Size = New-Object System.Drawing.Size(280, 20)
    $subForm.Controls.Add($yourNameTextBox)

    # Create the "Submit" button
    $subButton = New-Object System.Windows.Forms.Button
    $subButton.Location = New-Object System.Drawing.Point(115, 70)
    $subButton.Size = New-Object System.Drawing.Size(75, 23)
    $subButton.Text = "Submit"
    $subButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $subForm.AcceptButton = $subButton
    $subForm.Controls.Add($subButton)

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