#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 0.0.0.2
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Script to clear the cache of applications from the BBCC module
.DESCRIPTION
    - Promps the user for an application to clear cache from
    - Pulls the list from the BBCC module
    - performs the action once 'ok' is selected
    - closes the application once the action is complete
#>


<#
--------------------
 VARIBALE PARADISE!
--------------------
#>
#Directories
$tempDir = New-Item -ItemType Directory -Path "$env:TEMP" -Name Custom_Scripts -Force

# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\Custom_Scripts
$appLogs = "$tempDir\Custom_Scripts.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

# Set the URL of the BBCC module 
$githubUrl = "https://raw.githubusercontent.com/naklsonofnakkl/POWERSHELL/main/MISC/BBCC.psm1"

# Download the module
$bbccPath = Join-Path -Path $tempDir -ChildPath "BBCC.psm1"
Invoke-WebRequest -Uri $githubUrl -OutFile $bbccPath

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

function Clear-Installation {
    Stop-Transcript
    exit
}

function show-CacheClear {
Add-Type -AssemblyName PresentationFramework
# Create the GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="NaklWilson Cache Clear" Width="400" Height="200" Topmost="True" Background="DarkBlue" Foreground="White" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>
        <TextBlock Grid.Column="0" Grid.Row="0" Text="Choose an Application:" Margin="10" TextAlignment="Center" FontSize="20" />
        <ComboBox Name="dropdown" Grid.Column="0" Grid.Row="1" Margin="10" />
        <Button Name="okButton" Grid.Row="2" Content="Ok" Margin="0,28,0,0" HorizontalAlignment="Center" VerticalAlignment="Top" Width="100" FontSize="18" >
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Import the module and retrieve the functions
Import-Module $bbccPath -Force
$moduleFunctions = Get-Command -Module BBCC -CommandType Function

# Modify the function names
$formattedFunctions = $moduleFunctions | ForEach-Object {
    # Remove "clear-" from function name for display in dropdown
    $name = $_ -replace '^clear-', ''  
    $name
}

# Add the formatted function names to the drop-down list
$dropdown = $Window.FindName("dropdown")
$dropdown.ItemsSource = $formattedFunctions

# Add functions for OK button
$okButton = $Window.FindName("okButton")
$okButton.Add_Click({
        # Get the selected function from the dropdown
        $selectedFunction = $dropdown.SelectedItem

        # Add "clear-" prefix to the selected function name
        $selectedFunction = "clear-" + $selectedFunction

        # Execute the selected function
        & $selectedFunction

        # Close the window after executing the function
        $Window.Close()
    })

# Show the window
$Window.ShowDialog() | Out-Null
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

show-CacheClear
Clear-Installation