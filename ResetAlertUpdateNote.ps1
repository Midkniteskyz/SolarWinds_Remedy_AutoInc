[CmdletBinding()]

#region Parameters
Param(

    [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('IncMsg')]
    [string]$IncidentMessage,

        [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('AltNote')]
    [string]$AlertNote

)
#endregion Parameters

# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -ExecutionPolicy Unrestricted -File 'C:\RemedyAPI\ResetAlertUpdateNote.ps1' -AlertNote "${N=Alerting;M=Notes}" -IncidentMessage "Alert has reset in SolarWinds."

#region Test Remedy Connection

# Test connection to the CS Remedy Server. Change to AJ if unavailable.
function Test-ServerConnection {
    <#
    .SYNOPSIS
    Tests the connection to a Remedy server.

    .DESCRIPTION
    The Test-ServerConnection function tests the connection to a Remedy server. The function accepts an optional switch parameter, -Test, which specifies whether to test the connection to the production or test server. If the -Test switch is not specified, the function tests the connection to the production server.

    .PARAMETER Test
    Optional switch parameter that specifies whether to test the connection to the production or test server.

    .OUTPUTS
    If the function successfully connects to a server, it returns the URL of the server.

    .NOTES
    This function requires the Test-Connection cmdlet.

    .LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions?view=powershell-7.1

    .EXAMPLE
    PS C:> Test-ServerConnection
    Testing connection to prodserver1 (production server)
    Connection to prodserver1 successful
    https://ProdServer1.org:443

    This example tests the connection to the production Remedy server.

    .EXAMPLE
    PS C:> Test-ServerConnection -Test
    Testing connection to testserver1 (test server)
    Connection to testserver1 successful
    https://TestServer1.org:8443

    This example tests the connection to the test Remedy server.

    #>

    [CmdletBinding()]
    param(
        [switch]$Test
    )
  
    # Define the servers
    if ($Test) {
        $servers = @(
            'https://TestServer1.org:8443'
        )
        $serverType = 'test'
    }
    else {
        $servers = @(
            'https://ProdServer1.org:443'
        )
        $serverType = 'production'
    }
  
    foreach ($server in $servers) {
        # extract the server name from the input URL
        $serverName = $server.TrimStart("https://").Split(":")[0].ToLower()
        Write-Host "Testing connection to $serverName ($serverType server)"
      
        # test connection to the server
        try {
            $null = Test-Connection -ComputerName $serverName -Count 2 -ErrorAction Stop
            Write-Host "Connection to $serverName successful"
            return $server
        }
        catch {
            Write-Warning "Connection to $serverName failed"
        }
    }
  
    # if all servers fail, write error message
    Write-Error "Unable to connect to any $serverType server"
}
#endregion Test Remedy Connection

#region Update work note
function Add-WorkNoteToIncident {
<#
.SYNOPSIS
This function adds a work note to a specified incident in BMC Remedy ITSM.

.DESCRIPTION
The function takes an incident number and a message as input, authenticates with the BMC Remedy ITSM REST API, and adds the message as a work note to the specified incident.

.PARAMETER incidentNumber
The number of the incident to which the work note should be added.

.PARAMETER message
The content of the work note to be added.

.EXAMPLE
$IncidentNumber = "INC000002099110"
$message = "Adding a new work note."
Add-WorkNoteToIncident -incidentNumber $IncidentNumber -message $message
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$incidentNumber,
        [Parameter(Mandatory=$true)]
        [string]$message,
        [Parameter(Mandatory=$false)]
        [switch]$test
    )

    try {

    # Define the base URL for the API
    if($test){
        $baseUrl = Test-ServerConnection -Test
    }else{
        $baseUrl = Test-ServerConnection
    }

    # Define the endpoint for the work log
    $workLogEndpoint = "/api/arsys/v1/entry/HPD:WorkLog"
  
        # Get the login credentials
        Write-Host "Getting credentials to login to Remedy"
        try {
            # Use this cred to running from SolarWinds
            $CredPath = "C:\RemedyAPI\Credentials_nt authority~system.xml"
            
            $Credential = Import-Clixml -Path $CredPath
            $Username = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password.ToString()
            Write-Host "Credentials successfully retrieved from $CredPath"
        }
        catch {
            Write-Host "Unable to grab credentials from $CredPath"      
            Write-Error "Failed to get Remedy credentials: $_"
            return
        }


        # Connect to the Jetty Server & get a token
        $JettyToken = Invoke-RestMethod -Uri "$baseURL/api/jwt/login" -Method 'POST' -ContentType 'application/x-www-form-urlencoded' -Body "username=$UserName&password=$Password"

    # Define the headers for the API request
    $headers = @{
        "Content-Type" = "application/json"
        "Authorization" = "AR-JWT $Jettytoken"  
    }

    # Define the body for the API request
    $body = @{
        "values" = @{
            "Incident Number" = $incidentNumber
            "Work Log Type" = "General Information"
            "Detailed Description" = $message
        }
    } | ConvertTo-Json

    # Send the API request
    $response = Invoke-RestMethod -Uri ($baseUrl + $workLogEndpoint) -Method Post -Headers $headers -Body $body

    # Return the response
    return $response
    }
    catch {
    $errorMessage = $_.Exception.Message
    $logFilePath = "C:\RemedyAPI\Logs\ResetAlertUpdateNote\ResetLog.txt"
    Add-Content -Path $logFilePath -Value ("[" + (Get-Date) + "] " + $errorMessage)
}
}
#endregion Update work note

#region Parse Incident Number
function Get-IncidentNumber {
    param(
        [Parameter(Mandatory=$true)]
        [string]$text
    )

    if ($text -match "Incident Number: (INC\d{12})") {
        return $Matches[1]
    } else {
        return $null
    }
}
#endregion Parse Incident Number

<#
$TestNote = @"
I am a test note
Incident Number: INC000002099110
ack
"@
#>
# $message = "Adding the incident number parser. Current test note is:`n$TestNote"

$IncidentNumber = Get-IncidentNumber $AlertNote
Add-WorkNoteToIncident -incidentNumber $IncidentNumber -message $IncidentMessage
