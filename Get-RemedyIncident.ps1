function Get-Incident {
    [CmdletBinding()]
    param (
        # Specify the Remedy Host. Default will be the Non-Prod Remedy Server.
        [Parameter()]
        [string]
        $RemedyHost = "" ,

        # Specify the Incident number to look up
        [Parameter(Mandatory = $true)]
        [string]
        $IncidentNumber
    )

    begin {
        Write-Verbose "In Begin Block"

        # Creds for connecting to the Jetty Server
        $UserName = "SolarWinds"
        $Password = 'Omn1$bus'

        # Connect to the Jetty Server & get a token
        $JettyToken = Invoke-RestMethod -Uri "$RemedyHost/api/jwt/login" -Method 'POST' -ContentType 'application/x-www-form-urlencoded' -Body "username=$UserName&password=$Password"

        Write-Verbose $JettyToken

        # Format anbd add the token to the header for connection to the Remedy server
        $Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $Headers.Add("Content-Type", "application/json")
        $Headers.Add("Authorization", "AR-JWT $JettyToken")

        # Generate the query URL to Remedy
        $Query = "?q=%27Incident%20Number%27%3D%22$IncidentNumber%22"
        $URL = "$RemedyHost/api/arsys/v1/entry/HPD:IncidentInterface" + $Query

    }

    process {
        Write-Verbose "In Process Block"

        # Send the query to Remedy, Capture the response 
        $Response = Invoke-RestMethod $URL -Method 'GET' -Headers $Headers

        $IncidentInformation = [PSCustomObject]@{
            IncidentNumber              = $Response.entries.values.'Incident Number'
            Site                        = $Response.entries.values.Site
            Summary                     = $Response.entries.values.Description
            Notes                       = $Response.entries.values.'Detailed Decription'
            Submitter                   = $Response.entries.values.Submitter
            SubmitDate                  = $Response.entries.values.'Submit Date'
            LastModifiedDate            = $Response.entries.values.'Last Modified Date'
            Status                      = $Response.entries.values.Status
            AssignedSupportCompany      = $Response.entries.values.'Assigned Support Company'
            AssignedSupportOrganization = $Response.entries.values.'Assigned Support Organization'
            AssignedSupportGroup        = $Response.entries.values.'Assigned Group'
            Assignee                    = $Response.entries.values.Assignee
        }
        
    }

    end {
        Write-Verbose "End Block:"

        # Release the token
        Invoke-RestMethod -Uri "$RemedyHost/api/jwt/logout" -Method 'POST' -Headers $Headers

        # Uncomment to view the response in JSON format
        # Write-Host $Response | ConvertTo-Json

        return $IncidentInformation

    }
}

$IncidentNumber = ""
$RemedyHost = ""

Get-Incident -RemedyHost "" -IncidentNumber ""
