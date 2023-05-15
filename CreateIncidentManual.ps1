$filename = Get-Date -Format FileDateTime
$LogsDir = "$Env:USERPROFILE\Desktop\Logs"
$LogDirSuccess = "$LogsDir\Success"
$LogDirFailed = "$LogsDir\Failed"
Start-Transcript -Path "$LogsDir\$filename.txt" -NoClobber

function New-Incident {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('RU', 'RURL', 'RmdyURL')]
        [string] $RemedyURL,
  
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('IC', 'IncCust', 'IncCustomer')]
        [string] $IncidentCustomer,
  
        # Incident Summary
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('IS', 'IncSum', 'IncSummary')]
        [string] $IncidentSummary,
  
        # Incident Notes
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('IN', 'IncNote', 'IncNotes')]
        [string] $IncidentNote,
  
        # Incident Impact
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('II', 'IncImpct', 'IncImpact')]
        [string] $IncidentImpact,
  
        # Incident Urgency
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('IU', 'IncUrg', 'IncUrgency')]
        [string] $IncidentUrgency,
  
        # Incident Tenant
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('Tnt', 'IncTnt', 'IncTenant')]
        [string] $IncidentTenant,
  
        # Incident Environment
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('IE', 'IncEnv', 'IncEnvironment')]
        [string] $IncidentEnvironment,
  
        # Incident Template
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('Temp', 'IncTemp', 'IncTemplate')]
        [string] $IncidentTemplate,
  
        # SolarWinds Alert URL
        [Parameter(
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('AU', 'AltURL')]
        [string] $AlertURL = '',
  
        # SolarWinds KBA ID passthru 
        [Parameter(
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('AK', 'AltKBA')]
        [string] $AlertKBA = ''
  
    )
  
    begin {
        Write-Debug "In the Begin block: New-RemedyTicket"
  
        # Define API endpoints
        $GetTokenURL = $RemedyURL + "/api/jwt/login"
        $CreateIncidentURL = $RemedyURL + "/api/arsys/v1/entry/HPD:IncidentInterface_Create?fields=values(Incident Number)"
        $ReleaseTokenURL = $RemedyURL + "/api/jwt/logout"
  
        # Create an encrypted UserName and Password with this command
        #$Credential = Get-Credential
        #$Credential | Export-Clixml -Path "D:\RemedyAPI\Credentials.xml"
  
        Write-Host "Getting credentials to login to Remedy"
        try {
            # $CredPath = "D:\RemedyAPI\Credentials.xml"
            # $Credential = Import-Clixml -Path $CredPath
            # $Username = $Credential.UserName
            # $Password = $Credential.GetNetworkCredential().Password
            $UserName = "SolarWinds"
            $Password = 'Password'
            Write-Host "Credentials successfully retrieved from $CredPath"
        }
        catch {
            Write-Host "Unable to grab credentials from $CredPath"      
            Write-Error "Failed to get Remedy credentials: $_"
            return
        }
  
        # Get token value from Jetty server in Remedy
        try {          
            Write-Host "Getting token from Jetty server: $GetTokenURL"
            $Response = Invoke-RestMethod -Uri $GetTokenURL -Method 'POST' -ContentType 'application/x-www-form-urlencoded' -Body "username=$Username&password=$Password"
        }
        catch {
            Write-Error "Failed to get token from Remedy server: $_"
            return
        }
  
        # Create token
        $Token = "AR-JWT " + $Response
        Write-Host "Token Received: $Token"
  
        $Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $Headers.Add("Content-Type", "application/json")
        $Headers.Add("Authorization", $Token)
        
        # VCS tenant doesn't exist in Remedy. Change it if found.
        if ($IncidentTenant -like "VCS") {
            $IncidentTenant = "TVS"
        }
        
        # Body to be sent to Remedy
        $Body = @{
            values = [ordered]@{
                z1D_Action          = "CREATE"
                Login_ID            = $IncidentCustomer
                Description         = $IncidentSummary 
                Detailed_Decription = $IncidentNote
                Impact              = $Impact
                Urgency             = $Urgency
                Tenant              = $IncidentTenant
                chr_Environment     = $IncidentEnvironment
                TemplateID          = $IncidentTemplate
                SW_URL_AlertDetails = $AlertURL
                SW_KB_ID            = $AlertKBA
            }
        }
  
        $UpdateBody = ConvertTo-Json $Body 
  
        Write-Host "-----Request to be sent to Remedy-----"
        foreach ($key in $Body['values'].Keys) {
            Write-Host "${key}: $($Body['values'][$key])"
        }
    
    }
  
    process {
        Write-Debug "In the Process block: New-RemedyTicket"
  
        try {
            Write-Host "Sending request to Remedy."
            $CreateIncident = Invoke-RestMethod -Uri $CreateIncidentURL -Method 'POST' -Headers $Headers -Body $UpdateBody -Verbose
        }
        catch {
            Write-Error "Failed to create a new Remedy ticket: $_"
            return
        }
    }
  
    end {
        Write-Debug "In the End block: New-RemedyTicket"
        Write-Host "Ticket successfully created in Remedy."
  
        # Release Token
        Write-Host "Releasing Token"
        Invoke-RestMethod -Uri $ReleaseTokenURL -Method 'POST' -Headers $Headers -Verbose
  
        # Display the Response (Incident Number) for Acknowledgment
        $IncidentNumber = $CreateIncident.Values
  
        return $IncidentNumber    
    }
}

$RemedyHost = ""
$LoginID = "ryan.woolsey"
$RemedySummary = "This is a test Summary"
$Detailed_Description = "This is a test note"
$Impact = 3000
$Urgency = $Impact
$Environment = "on-Prd"
$Tenant = "STAMP"
$RemedyTemplate = ""
$RemedyOrganization = "STAMP"
$RemedyAssignedGroup = "Monitoring"
$SW_ApplicationKBA = ""
$SW_AlertURL = "4"

try {
    $IncidentNumber = New-Incident -RemedyURL $RemedyHost -IncCustomer $LoginID -IncSummary $RemedySummary -IncNotes $Detailed_Description -IncImpact $Impact -IncUrgency $Urgency -IncEnvironment $Environment -IncTenant $Tenant -IncTemplate $RemedyTemplate -AltKBA $SW_ApplicationKBA -AltURL $SW_AlertURL -ErrorVariable NewIncidentErrors -ErrorAction Stop

    Write-Host $IncidentNumber
} catch {
    Write-Error $_
    Stop-Transcript
    Move-Item -Path "$LogsDir\$filename.txt" -Destination $LogDirFailed
} finally {
    Stop-Transcript
    if (Test-Path -Path "$LogsDir\$filename.txt") {
        Move-Item -Path "$LogsDir\$filename.txt" -Destination $LogDirSuccess
    }
    
}




