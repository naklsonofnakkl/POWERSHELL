Add-Type -AssemblyName PresentationFramework

$messageBoxResult = [System.Windows.MessageBox]::Show("Do you wish to Install (Yes) or Uninstall (No) the Print to PDF feature?" , 'Question' , [System.Windows.MessageBoxButton]::YesNoCancel , [System.Windows.MessageBoxImage]::Question)
switch ($messageBoxResult) {
    { $_ -eq [System.Windows.MessageBoxResult]::Yes } {
        Clear
        Write-Host "Installing the print-to-pdf feature..." -ForegroundColor Green
        Enable-WindowsOptionalFeature -Online -FeatureName "Printing-PrintToPDFServices-Features" -All
        $input = $(Write-Host "Please Press Enter Button to Restart..." -ForegroundColor Yellow -NoNewLine; Read-Host)
        Restart-Computer
        break
    }

    { $_ -eq [System.Windows.MessageBoxResult]::No } {
        Clear
        Write-Host "Uninstalling the print-to-pdf feature..." -ForegroundColor Red
        Disable-WindowsOptionalFeature -Online -FeatureName "Printing-PrintToPDFServices-Features"
        $input = $(Write-Host "Please Press Enter Button to Restart..." -ForegroundColor Yellow -NoNewLine; Read-Host)
        Restart-Computer
        break
    }

    { $_ -eq [System.Windows.MessageBoxResult]::Cancel } {
        Clear
        Write-Host "Cancelled Script" -ForegroundColor Magenta
        break
    }

    default {
        # stop
        return # or EXIT
    }
}
