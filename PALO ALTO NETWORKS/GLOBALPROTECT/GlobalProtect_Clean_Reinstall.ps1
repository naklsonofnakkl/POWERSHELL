#Requires -RunAsAdministrator

<#
.NOTES
    Author: Andrew Wilson
    Version: 1.0.0.0
    
.LINK
    https://github.com/naklsonofnakkl/POWERSHELL

.SYNOPSIS
    Reinstall GlobalProtect! 
.DESCRIPTION
    - prompts the user to uninstall or install GP
    - requires the user to input their portal address for install
    - can perform a full uninstall (including cache clear)
    - can perform a touchless installation (no user input required)
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
$appLogs = "$tempDir\Global_Protect.log"
$ErrorActionPreference = "Stop"
Start-Transcript -Path $appLogs -Append

$GP = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq "GlobalProtect" }

<#
--------------------
FUNCTION JUNCTION!
--------------------
#>

function Clear-Installation {
    Stop-Transcript
    exit
}

function Show-GlobalChoice {
    Add-Type -AssemblyName PresentationFramework

    $installresult = [System.Windows.MessageBox]::Show('Would you like to install GlobalProtect?', 'GlobalProtect Installer', 'YesNo', 'Question')

    if ($installresult -eq 'Yes') {
        # Insert the logic for 'Yes' choice, e.g. install GlobalProtect
        Write-Host "You selected Yes"
        Show-GlobalInstall
    }
    else {
        # Insert the logic for 'No' choice, e.g. exit or do nothing
        Write-Host "You selected No"
        Clear-Installation
    }

}

function Clear-GlobalProtect {
    try {
        ## Uninstall GP
        write-host "Uninstalling GlobalProtect..."
        $GP.uninstall()
        write-host "Uninstalled!"
        Start-Sleep 1
        ## Clear GP cache
        write-host "Clearing GlobalProtect Cache..."
        set-location "$env:LOCALAPPDATA\Palo Alto Networks\GlobalProtect"
        Remove-Item OLD -Recurse -ErrorAction -SilentlyContinue
        New-Item OLD -ItemType Directory
        Move-Item -Path .\*.dat .\OLD
        Move-Item -Path .\*.pan .\OLD
        write-host "GlobalProtect Cache Cleared!"
        Start-Sleep 3
        Show-GlobalChoice
    }
    catch {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Unable to uninstall GlobalProtect, please open the log file for more details!", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Get-GlobalProtect {
    try {
        Write-Host "INSTALL GP!"
        ## Install GP
        write-host "Downloading GlobalProtect..."
        Invoke-WebRequest -Uri "https://$GATEWAY/global-protect/getmsi.esp?version=64&platform=windows" -OutFile $home\Downloads\GlobalProtect64.msi -UseDefaultCredentials
        write-host "Download Complete!"
        write-host "Installing GlobalProtect..."
        Start-Process $home\Downloads\GlobalProtect64.msi -ArgumentList "/quiet /passive"
        write-host "GlobalProtect has been Reinstalled Successfully!"
        Remove-Item "$home\Downloads\GlobalProtect64.msi"
        Start-Sleep 3
    }
    catch {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show("Unable to install GlobalProtect, please open the log file for more details!", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Show-GlobalInstall {
    Add-Type -AssemblyName PresentationFramework
    # Create the GUI
    [xml]$xaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="GlobalProtect Installer" Width="400" Height="300" Topmost="True" Background="SlateGray" Foreground="White" WindowStartupLocation="CenterScreen">
<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />
        <RowDefinition Height="Auto" />
        <RowDefinition Height="*" />
    </Grid.RowDefinitions>
    <TextBlock Grid.Column="0" Grid.Row="0" Text="Enter your Portal Address:" Margin="10" TextAlignment="Center" FontSize="20" FontFamily="Consolas" />
    <TextBox Name="textbox" Grid.Column="0" Grid.Row="1" Margin="10,10,10,92" Grid.RowSpan="2" FontFamily="Consolas" FontSize="18" Height="36"/>
    <Button Name="okButton" Grid.Row="3" Content="Install" HorizontalAlignment="Center" VerticalAlignment="Top" Width="125" FontSize="18" Margin="0,162,0,0" FontFamily="Consolas" Height="31" >
        <Button.Effect>
            <DropShadowEffect/>
        </Button.Effect>
    </Button>
</Grid>
</Window>
"@

    $Ireader = New-Object System.Xml.XmlNodeReader $xaml
    $IWindow = [Windows.Markup.XamlReader]::Load($Ireader)

    # Reference to the textbox
    $textbox = $IWindow.FindName("textbox")

    # Add functions for OK button
    $okButton = $IWindow.FindName("okButton")
    $okButton.Add_Click({
            # Assign the value from the textbox to the variable $GATEWAY
            $GATEWAY = $textbox.Text

            # Call your function
            Get-GlobalProtect

            # Close the window after executing the function
            $IWindow.Close()
        })

    # Show the window
    $IWindow.ShowDialog() | Out-Null
}

function Show-GlobalUninstall {
    Add-Type -AssemblyName PresentationFramework
    # Create the GUI
    [xml]$xaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="GlobalProtect Installer" Width="407" Height="178" Topmost="True" Background="SlateGray" Foreground="White" WindowStartupLocation="CenterScreen">
<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />
        <RowDefinition Height="Auto" />
        <RowDefinition Height="*" />
    </Grid.RowDefinitions>
    <TextBlock Grid.Column="0" Grid.Row="0" Text="Choose an option:" Margin="10" TextAlignment="Center" FontSize="20" FontFamily="Consolas" />
    <Button x:Name="InstallButton" Grid.Row="1" Content="Install" HorizontalAlignment="Center" VerticalAlignment="Center" Width="124" FontSize="24" FontFamily="Consolas" Height="31" >
        <Button.Effect>
            <DropShadowEffect/>
        </Button.Effect>
    </Button>
    <Button x:Name="UninstallButton" Grid.Row="2" Content="Uninstall" HorizontalAlignment="Center" VerticalAlignment="Center" Width="124" FontSize="24" FontFamily="Consolas" Height="31" >
        <Button.Effect>
            <DropShadowEffect/>
        </Button.Effect>
    </Button>
</Grid>
</Window>
"@

    $Ureader = New-Object System.Xml.XmlNodeReader $xaml
    $UWindow = [Windows.Markup.XamlReader]::Load($Ureader)

    # Add functions for install button
    $installButton = $UWindow.FindName("InstallButton")
    $installButton.Add_Click({
            # Close the window after executing the function
            $UWindow.Close()
            # Call your function
            Show-GlobalInstall
        })

    # Add functions for install button
    $uninstallButton = $UWindow.FindName("UninstallButton")
    $uninstallButton.Add_Click({
            # Close the window after executing the function
            $UWindow.Close()
            # Call your function
            Clear-GlobalProtect
    

        })    

    # Show the window
    $UWindow.ShowDialog() | Out-Null
}

<#
--------------------
SCRIPTED EXECUTION!
--------------------
#>

Show-Globaluninstall
Clear-Installation

