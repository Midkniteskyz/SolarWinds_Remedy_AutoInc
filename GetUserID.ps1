# Get the current user's account and session information
$whoami = whoami
$quser = quser | where { $_ -match "$env:USERNAME" }
$currentUser = (Get-CimInstance win32_computersystem).UserName
$sessionInfo = Get-PSSession | Select-Object Id, Name, ComputerName, State, ConfigurationName

# Get the password information
$CredPath = "C:\RemedyAPI\Credentials_nt authority~system.xml"
$Credential = Import-Clixml -Path $CredPath
$Password = $Credential.GetNetworkCredential().Password

# Construct the output string
$outputString = "Who Am I: $whoami`nQ User: $quser`nGet-Cim Username: $currentUser`nSession Information:`n$($sessionInfo | Format-Table | Out-String)`nPassword: $Password"

# Get the current script's directory and create the output file path
$scriptDirectory = Split-Path $MyInvocation.MyCommand.Path
$outputFilePath = Join-Path $scriptDirectory "ID.txt"

# Write the output to the file
$outputString | Out-File -FilePath $outputFilePath -Encoding utf8 -Append
