<#
 .Synopsis
  Lookup users in AD for specific properties based on a csv list located in OneDrive
  and export the output them to separate csv list back into OneDrive

 .Description
  Asks the user to provide the name of their company as found on their OneDrive
  for Business folder and will import the lookup list followed by exporting the
  output into another csv list

  .Example
  # This will import the csv of users needing to be looked up in AD
  # Check and see if the users are located in AD and gather the properties requested
  # Put the gathered users onto an array separated by properties
  Get-Users

  .Example
  # This will perform a general AD lookup using the email provided.
  # Will print the results to the console.
  Export-User Test@email.com
  
  .Example
  # This will convert the results array into a csv file
  # It will export the csv file to a specific location provided
  Export-Users
#>


<#
--------------------
 VARIABLE PARADISE!
--------------------
#>
# LOGS
# C:\Users\[USERNAME]\AppData\Local\Temp\
$tempDir = $env:TEMP
$appLogs = "$tempDir\AD_unlock_accounts.log"
#OneDrive Location
$company = "Consco"
# Array for the User List
$results = @()
# Output the Export Status
$exportStatus = "Export Compelted!"


<#
--------------------
FUNCTION JUNCTION!
--------------------
#>
# Function to lookup the list of users from the CSV and place them into the results array
function Get-Users {
    $global:company = Read-Host -Prompt "Please enter your company name EXACTLY as it appears on OneDrive folder and press enter... (i.e Consco, Test Company)"
    $users = Import-Csv -Path "$env:userprofile\OneDrive - $company\Documents\adusers.csv"
    foreach ($user in $users) {

        try {
            $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$($user.UserPrincipalName)'" -Properties msExchExtensionAttribute45, Enabled, DistinguishedName

        }
        catch {
            Write-Error "An error occurred: $_"
            Write-Host "$_" | Out-File -FilePath $appLogs -Append
        }
        finally {
            if ($adUser) {
                $global:results += [PSCustomObject]@{
                    UserPrincipalName          = $adUser.UserPrincipalName
                    msExchExtensionAttribute45 = $adUser.msExchExtensionAttribute45
                    Enabled                    = $adUser.Enabled
                    DistinguishedName          = $adUser.DistinguishedName
                }
            }
            else {
                $global:results += [PSCustomObject]@{
                    UserPrincipalName          = $($user.UserPrincipalName)
                    msExchExtensionAttribute45 = "NULL"
                    Enabled                    = "NULL"
                    DistinguishedName          = "NULL"
                }
            }
        }
    }      
}

#Function to export a single user (e.g. Export-User test@ucdavis.edu)
function Export-User {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
    try {
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -Properties msExchExtensionAttribute45, Enabled, DistinguishedName        
        if (!$adUser) {
            Write-Warning "$UserPrincipalName was not found in Active Directory!"
        }
        else {
            $adUser
        }
    }
    catch {
        $global:exportStatus = "An error occurred: $_"
        Write-Host "$_" | Out-File -FilePath $appLogs -Append
        Write-Host $exportStatus
    }
}

#Function to export the list of users from the result array into a CSV
function Export-Users {
    $csvexport = "$env:userprofile\OneDrive - $global:company\Documents\ADUser_ExtAttr45.csv"
    if (Test-Path -Path $csvexport) {
        try {
            Remove-Item -Path "$csvexport" -Recurse -Force
            $results | Export-Csv -Path "$csvexport" -NoTypeInformation
            $global:exportStatus = "Export Compelted Successfully!"
        }
        catch {
            $global:exportStatus = "An error occurred: $_"
            Write-Host "$_" | Out-File -FilePath $appLogs -Append
        }
        finally {
            Clear-Host
            Write-Host $exportStatus
        }
    }
    else {
        try {
            Write-Host "No Previous Export to remove..." 
            $results | Export-Csv -Path "$csvexport" -NoTypeInformation
            $global:exportStatus = "Export Compelted Successfully!"
        }
        catch {
            $global:exportStatus = "An error occurred: $_"
            Write-Host "$_" | Out-File -FilePath $appLogs -Append
        }
        finally {
            Clear-Host
            Write-Host $exportStatus
        }
    }
}