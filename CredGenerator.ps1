<#
.SYNOPSIS
Use this script to create a credential for Remedy in a specified location.

.DESCRIPTION
This script creates a credential for Remedy and exports it as an encrypted XML file at the specified location.

.PARAMETER User
Specifies the username for the credential. Default value is "SolarWinds".

.PARAMETER Password
Specifies the password for the credential. This parameter is mandatory.

.PARAMETER ExportPath
Specifies the location where the credential file will be exported. Default value is "C:\RemedyAPI\Credentials.xml".

.EXAMPLE
.\CredGenerator.ps1 -Password "<password>"

This example creates a credential for SolarWinds to access Remedy with the password "<password>" and exports it to the default location "C:\RemedyAPI\Credentials.xml".

.NOTES
Version:        1.0
Author:         Ryan Woolsey
Creation Date:  22-Mar-2023
Purpose/Change: Initial script development

.LINK

#>

[CmdletBinding()]
param(
    [string]$User = "SolarWinds",
    [Parameter(Mandatory=$true)]
    [string]$Password,
    [string]$ExportPath = "C:\RemedyAPI\Credentials_"+((whoami).replace('\','~'))+".xml"
)

# Convert the password to a secure string
try {
    $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
} catch {
    Write-Error "Failed to convert password to secure string. $($Error[0].Exception.Message)"
    break
}

# Create a new credential object
try {
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $SecurePassword
} catch {
    Write-Error "Failed to create credential object. $($Error[0].Exception.Message)"
    break
}

# Export the credential as an encrypted XML file
try {
    $Credential | Export-Clixml -Path $ExportPath
} catch {
    Write-Error "Failed to export credential. $($Error[0].Exception.Message)"
    break
}

$account = whoami

Write-Host "$account created credential successfully and exported to $ExportPath"
