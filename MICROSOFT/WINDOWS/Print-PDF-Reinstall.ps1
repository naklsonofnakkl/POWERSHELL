$install = New-Object System.Management.Automation.Host.ChoiceDescription "&Install", "Enable the Printing-PrintToPDFServices-Features Windows feature."
$uninstall = New-Object System.Management.Automation.Host.ChoiceDescription "&Uninstall", "Disable the Printing-PrintToPDFServices-Features Windows feature."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($install, $uninstall)
$result = $host.ui.PromptForChoice("PrintToPDFServices Features", "Select an action:", $options, 0)

if ($result -eq 0) {
    Enable-WindowsOptionalFeature -Online -FeatureName Printing-PrintToPDFServices-Features
}
elseif ($result -eq 1) {
    Disable-WindowsOptionalFeature -Online -FeatureName Printing-PrintToPDFServices-Features
}
else {
    Write-Host "Invalid choice."
}
exit
