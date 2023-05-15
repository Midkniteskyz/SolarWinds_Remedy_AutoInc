[CmdletBinding()]

<#
    Authors: Stephen Ferrari, Ryan Woolsey

    .SYNOPSIS
    Automatically create a Remedy incident when an alert is triggered in SolarWinds.

    .DESCRIPTION
    This script will create a Remedy Incident when a SolarWinds alert triggers. The Incident inforation will be a variables populated by the Alert's custom properties. After the Incident is created, the alert note is updated with the incident number. 

    .EXAMPLE
    Add to the Trigger Action of the alert as an "Execute an External Program"

    1) Ingest the variables evaluated by SolarWinds. Update the alert note, link back to the SolarWinds Alert, and add the related KB article
        - C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -ExecutionPolicy Unrestricted -File "C:\RemedyAPI\SolarWindsConnector.ps1" -RemedyCustomer "${N=SwisEntity;M=CustomProperties.Remedy_Customer}" -RemedySummary "${N=Alerting;M=AlertDescription}" -RemedyNotes "${N=Alerting;M=AlertMessage}" -SW_Severity "${N=Alerting;M=Severity}" -RemedyTenant "${N=SwisEntity;M=CustomProperties.Remedy_Tenant}" -RemedyEnvironment "${N=SwisEntity;M=CustomProperties.Remedy_Environment}" -RemedyCompany "${N=SwisEntity;M=CustomProperties.Remedy_Company}" -RemedyOrganization "${N=SwisEntity;M=CustomProperties.Remedy_Organization}" -RemedyAssignedGroup "${N=SwisEntity;M=CustomProperties.Remedy_Assigned_Group}" -SW_AlertID "${N=Alerting;M=AlertObjectID}" -SW_AlertURL "${N=Alerting;M=AlertDetailsUrl}" -SW_ApplicationKBA "${N=SwisEntity;M=Integrations.Application.CustomProperties.Application_Remedy_KBA}"

    2) Manual entry with bare minimum required by Remedy. 
        - C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -ExecutionPolicy Unrestricted -File "C:\RemedyAPI\SolarWindsConnector.ps1" -RemedyCustomer "$Joe.Someone" -RemedySummary "Test Summary" -RemedyNotes "Test Note" -SW_Severity "Informational" -RemedyCompany "Company" -RemedyOrganization "Org1" -RemedyAssignedGroup "Group 1" -SW_AlertID "${N=Alerting;M=AlertObjectID}"

    .NOTES
    The script will accept any data types from SolarWinds, however functions may not work properly if the wrong type is provided. There is error handling to ensure data is being passed through properly and can be viewed at "C:\RemedyAPI\Logs\". Be mindful that Remedy will reject incorrect vaules sent to the API from the script. It's important to coordinate with the Remedy Admin to sync these values up with the SolarWinds Custom properties.  

    .LINK
    SWIS PowerShell / Orion SDK
    - https://github.com/solarwinds/OrionSDK

    Overview of the Remedy REST API
    - https://docs.bmc.com/docs/ars2002/overview-of-the-rest-api-909638130.html

    Remedy AR System - REST API, how to use it with Powershell's Invoke-RestMethod Commands
    - https://community.bmc.com/s/article/Remedy-AR-System-REST-API-how-to-use-it-with-Powershell-s-Invoke-RestMethod-Commands
#>

#region Parameters
Param(
    # No custom property in SolarWinds. The value is specified in the script execution. 

    [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RU', 'RURL', 'RmdyURL')]
    [string]$RemedyURL,

    # Custom property in SolarWinds is Remedy_Customer. 
    # SWIS is ${N=SwisEntity;M=CustomProperties.Remedy_Customer}
    # In Remedy, the Customer is typically the mission lead.
    [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RCu', 'RCus', 'RmdyCustomer')]
    [string]$RemedyCustomer,

    # Remedy Incident Summary field. Currently this is the SolarWinds Alert Name. 
    # SWIS is ${N=Alerting;M=AlertDescription}
    # There is some trimming that occurs with the name to reduce chars. Max Chars is 255
    [Parameter(
        Mandatory = $false,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RS', 'RSum', 'RmdySummary', 'Description')]
    [string]$RemedySummary,

    # Remedy Incident Notes field. The SolarWinds Alert Message is ingested into this. 
    # SWIS is ${N=Alerting;M=AlertMessage}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RN', 'RNot', 'RmdyNotes', 'Detailed_Description')]
    [string]$RemedyNotes,

    # SolarWinds Alert property. This is digested to become the Impact and Urgency in Remedy. 
    # SWIS is ${N=Alerting;M=Severity}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('Sev', 'RmdyImpUrg', 'Severity')]
    [string]$SW_Severity,

    # Remedy Tenant field value. Custom property in SolarWinds Remedy_Tenant. 
    # SWIS is ${N=SwisEntity;M=CustomProperties.Remedy_Tenant}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RT', 'RTen', 'RmdyTenant')]
    [string]$RemedyTenant,

    # Remedy Environment field value. Custom property in SolarWinds Remedy_Environment. 
    # SWIS is ${N=SwisEntity;M=CustomProperties.Remedy_Environment}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RE', 'REnv', 'RmdyEnvironment')]
    [string]$RemedyEnvironment,

    # Remedy Company field value. Custom property in SolarWinds Remedy_Company. 
    # SWIS is ${N=SwisEntity;M=CustomProperties.Remedy_Company}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RCo', 'RCom', 'RmdyCompany')]
    [string]$RemedyCompany,

    # Remedy Organiztion field value. Custom property in SolarWinds Remedy_Organization.
    # SWIS is ${N=SwisEntity;M=CustomProperties.Remedy_Organization}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RO', 'ROrg', 'RmdyOrganization')]
    [string]$RemedyOrganization,

    # Remedy Assigned Group field value. Custom property in SolarWinds Remedy_Assigned_Group.
    # SWIS is ${N=SwisEntity;M=CustomProperties.Remedy_Assigned_Group}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('RA', 'RAG', 'RmdyAssignedGroup')]
    [string]$RemedyAssignedGroup,

    # Custom property in SolarWinds SystemName. Used to identify the main Device triggering the Alert 
    # SWIS is ${N=SwisEntity;M=SysName;F=OriginalValue}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('SN', 'SysName', 'SysN')]
    [string]$SystemName,

    # Alert property in SolarWinds. Unique ID for the Alert Triggering.
    # SWIS is ${N=Alerting;M=AlertObjectID}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('AI', 'SWAID', 'AlertID')]
    [string]$SW_AlertID,

    # Alert property in SolarWinds. Used to pass the Alert Details page to Remedy. 
    # SWIS is ${N=Alerting;M=AlertDetailsUrl} 
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('AU', 'SWAURL', 'AlertURL')]
    [string]$SW_AlertURL,

    # Application custom property in SolarWinds. Passed to Remedy to link to the related KBA in Remedy.
    # SWIS is ${N=SwisEntity;M=Integrations.Application.CustomProperties.Application_Remedy_KBA}
    [Parameter(
        Mandatory = $false,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('AK', 'SWAKBA', 'ApplicationKBA')]
    [string]$SW_ApplicationKBA

)
#endregion Parameters

#region Execution & Logging Configuration

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force

$uid = [guid]::NewGuid().ToString("N")
$date = Get-Date -Format "yyyyMMdd-HHmmss"
$filename = "${date}~${uid}.log"
$LogsDir = "C:\RemedyAPI\Logs"
$LogDirSuccess = "$LogsDir\Success"
$LogDirFailed = "$LogsDir\Failed"

if (!(Test-Path -Path $LogsDir -PathType Container)) {
    New-Item -ItemType Directory -Path $LogsDir | Out-Null
}
if (!(Test-Path -Path $LogDirSuccess -PathType Container)) {
    New-Item -ItemType Directory -Path $LogDirSuccess | Out-Null
}
if (!(Test-Path -Path $LogDirFailed -PathType Container)) {
    New-Item -ItemType Directory -Path $LogDirFailed | Out-Null
}

Start-Transcript -Path "C:\RemedyAPI\Logs\$filename.txt" -NoClobber

$Seperator = "-" * 100

#endregion Execution  &Logging Configuration

#region Write Input Parameters

$Seperator
"Parameters Received`r`n" 
foreach($boundparam in $PSBoundParameters.GetEnumerator()) {
    "{0} : {1}" -f $boundparam.Key,$boundparam.Value
}
$Seperator

#endregion Write Input Parameters

#region functions

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
            'https://TestServer1.org:8443',
            'https://TestServer1.org:8443'
        )
        $serverType = 'test'
    }
    else {
        $servers = @(
            'https://ProdServer1.org:443',
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
  
#region Set Incident Severity
function Set-Severity {
    <#
        .SYNOPSIS
        Sets the impact and urgency based on alert severity.

        .DESCRIPTION
        The Set-Severity function sets the impact and urgency levels based on the alert severity. The function takes a severity string as input and returns an object that includes the impact and urgency values.

        .PARAMETER Severity
        Specifies the severity of the alert. The allowed values are 'Critical', 'Serious', 'Warning', 'Informational', and 'Notice'.

        .OUTPUTS
        The function returns an object that includes the impact and urgency levels.

        .NOTES
        This function is based on the Remedy Alert Management system.

        .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions?view=powershell-7.1

        .EXAMPLE
        PS C:> Set-Severity -Severity 'Warning'

        This example sets the impact and urgency levels to 3000 based on the alert severity of 'Warning'.

        .EXAMPLE
        PS C:> 'Serious', 'Informational', 'Invalid' | Set-Severity

        This example sets the impact and urgency levels to 2000 and 4000 for 'Serious' and 'Informational' alerts respectively. It throws an error for 'Invalid' severity.

    #>
    [CmdletBinding()]
    [OutputType([int])]
    PARAM (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateSet('Critical', 'Serious', 'Warning', 'Informational', 'Notice')]
        [string[]]  $Severity
  
    )
  
    BEGIN {
        # Write a message to the host to indicate the function has started
        Write-Host "Starting Set-Severity function..."
    }
    
    PROCESS {
        # Evaluate Impact and Urgency arguments based on Alert Severity
        Switch ($Severity) {
            'Critical' { $Impact = $Urgency = 1000 }
            'Serious' { $Impact = $Urgency = 2000 }
            'Warning' { $Impact = $Urgency = 3000 }
            'Informational' { $Impact = $Urgency = 4000 }
            'Notice' { $Impact = $Urgency = 4000 }
            # Throw an error if an invalid severity is provided
            default { throw "Invalid severity: $Severity" }
        }
        # Write a message to the host to indicate what severity was evaluated
        Write-Host "Evaluated impact: $Impact"
        Write-Host "Evaluated urgency: $Urgency"
  
    }
    
    END {
        # Write a message to the host to indicate the function has completed
        Write-Host "Set-Severity function complete."
        return [PSCustomObject]@{
            Impact  = $Impact
            Urgency = $Urgency
        }
    }
}
#endregion Set Incident Severity
  
#region Set Incident Template
function Set-RemedyTemplate {
    <#
        .SYNOPSIS
        Determines which Remedy template to use for a ticket based on the "Remedy_Assigned_Group".

        .DESCRIPTION
        The Set-RemedyTemplate function determines which Remedy template to use for a ticket based on the "Remedy_Assigned_Group". The function accepts the Remedy Organization and Remedy Assigned Group as input, which can be passed as parameters or taken from global variables. The function returns the GUID of the calculated Remedy template.

        .PARAMETER RemedyOrganization
        Specifies the Remedy organization.

        .PARAMETER RemedyAssignedGroup
        Specifies the Remedy assigned group.

        .OUTPUTS
        The function returns the GUID of the calculated Remedy template.

        .NOTES
        This function is based on the Remedy ticketing system.

        .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions?view=powershell-7.1

        .EXAMPLE
        PS C:> Set-RemedyTemplate -RemedyOrganization "Org1" -RemedyAssignedGroup "Assigned Group 1"

        This example calculates the Remedy template to use for a ticket based on the Remedy Organization and Remedy Assigned Group parameters.

        .EXAMPLE
        PS C:> Set-RemedyTemplate

        This example calculates the Remedy template to use for a ticket based on the global variables $global:RemedyOrganization and $global:RemedyAssignedGroup.

    #>
    [CmdletBinding()]
    <#
        .Description
        Determines which Remedy template to use for a ticket based on the "Remedy_Assigned_Group"
    #>
    param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true                                   
        )]
        [ValidateNotNullOrEmpty()]
        [string]$RemedyOrganization = $global:RemedyOrganization,
  
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true                                   
        )]
        [ValidateNotNullOrEmpty()]
        [string]$RemedyAssignedGroup = $global:RemedyAssignedGroup
    )
    
    begin {
        Write-Debug "Entering Begin block: Set-RemedyTemplate"
        Write-Host "Current Remedy Organization: $RemedyOrganization"
        Write-Host "Current Remedy Assigned Group: $RemedyAssignedGroup"
  
        # Define a hashtable containing the Remedy templates for each organization and support group
        $TemplateCollection = @{
            Organization    = @{
                "AssignedGroup" = "TemplateGUID"
            }
        }        
    }
    
    process {
        Write-Debug "Entering Process block: Set-RemedyTemplate"
        $Template = $TemplateCollection.$RemedyOrganization.$RemedyAssignedGroup
  
        # Log the calculated template
        Write-Verbose "Calculated Remedy template is: $Template"
        Write-Host "Using Remedy template $Template"
    }
    
    end {
        Write-Debug "Entering End block: Set-RemedyTemplate"
        return $Template
    }
}
#endregion Set Incident Template
  
#region Set Incident Environment
function Set-Environment {
    <#
        .SYNOPSIS
        Sets the environment abbreviation based on the full environment name.

        .DESCRIPTION
        The Set-Environment function sets the environment abbreviation based on the full environment name. The function accepts the environment name as input, which can be passed as a parameter or taken from a global variable. The function returns the abbreviation of the environment name or "UNK" if it is not found in the environment map.

        .PARAMETER Environment
        Specifies the environment name.

        .OUTPUTS
        The function returns the abbreviation of the environment name.

        .NOTES
        This function is based on the mapping of environment names to their abbreviations.

        .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions?view=powershell-7.1

        .EXAMPLE
        PS C:> Set-Environment -Environment Production

        This example sets the environment abbreviation to "PRD" based on the environment name parameter.

        .EXAMPLE
        PS C:> Set-Environment

        This example sets the environment abbreviation based on the global variable $RemedyEnvironment.

    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Environment = $RemedyEnvironment
    )
  
    begin {
        # Log the beginning of the function
        Write-Host "Starting Set-Environment function..."
  
        # Log the current values of the parameters
        Write-Verbose "Current Environment: $Environment"
  
        # Define a hashtable to map environment names to their abbreviations
        $environmentMap = @{
            'Development'      = 'DEV'
            'Production'       = 'PRD'
            'Stage'            = 'STG'
            'Test'             = 'TST'
        }
  
        # Look up the abbreviation of the environment name in the hashtable, or default to "UNK"
        if ($environmentMap.ContainsKey($Environment)) {
            $Environment = $environmentMap[$Environment].ToUpper()
        }
        else {
            $Environment = ''
        }
    }
  
    process {
        # Log the current environment value
        Write-Verbose "New Environment: $Environment"
    }
  
    end {
        # Log the end of the function
        Write-Host "Set-Environment function completed."
  
        # Return nothing
        return $Environment
    }
}
#endregion Set Incident Environment
  
#region Update SolarWinds Alert Note
function Update-SWAlert {
    <#
        .SYNOPSIS
        Updates the note of a SolarWinds alert with a Remedy ticket number.

        .DESCRIPTION
        The Update-SWAlert function updates the note of a SolarWinds alert with a Remedy ticket number. The function requires the SolarWinds alert ID and the Remedy incident number. The function connects to the SolarWinds server using the SWIS PowerShell module and updates the alert note with the formatted message.

        .PARAMETER SW_AlertID
        Specifies the SolarWinds alert ID.

        .PARAMETER IncidentNumber
        Specifies the Remedy incident number.

        .NOTES
        This function requires the SWIS PowerShell module.

        .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions?view=powershell-7.1

        .EXAMPLE
        PS C:> Update-SWAlert -SW_AlertID 123456 -IncidentNumber INC123456

        This example updates the note of the SolarWinds alert with ID "123456" with the Remedy incident number "INC123456".

        .EXAMPLE
        PS C:> Update-SWAlert -SW_AlertID 123456

        This example updates the note of the SolarWinds alert with ID "123456" with a default message that a ticket could not be created in Remedy.

    #>
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $false,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]  $SW_AlertID,
        
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]  $IncidentNumber 
    )
    
    begin {
        Write-Debug "In the Begin block: Update-SWAlert"

        # Display the incident number received
        Write-Host "AlertID: $SW_AlertID"
        Write-Host "Incident Number is: $IncidentNumber"

        # Bring in the SWIS PowerShell module to use SWIS commands
        Import-Module swispowershell

        # Set the SolarWinds DB connection 
        $SWIS = Connect-SWIS -HostName localhost -Certificate

        # Format the alert note
        $AlertObjectIds = @([int]$SW_AlertID) 

        if ($IncidentNumber -like "" -or $IncidentNumber -eq $Null) {
            # If the incident number is not provided, set a default message
            $FormattedString = "A ticket could not be created in Remedy."
        }
        else {
            # If the incident number is provided, format the message to include it
            $FormattedString = "Incident Number: " + $IncidentNumber.TrimEnd('}') -replace '@{Incident Number=', ''
        }

    }
    
    process {
        Write-Debug "In the Process block: Update-SWAlert"

        # Display the alert ID and incident number being used
        Write-Verbose "Incident Number: $IncidentNumber"
        Write-Verbose "AlertID: $SW_AlertID"

        try {
            # Update the alert note with the formatted message
            Invoke-SwisVerb $SWIS -EntityName "Orion.AlertActive" -Verb "AppendNote" -Arguments @($AlertObjectIds, $FormattedString)
        }
        catch {
            # Catch any errors and display them
            Write-Error "Failed to update alert note: $_"
            return
        }
    }
    
    end {
        Write-Debug "In the End block: Update-SWAlert"
    }
}
#endregion Update SolarWinds Alert Note
  
#region Create New Incident
function New-Incident {
    <#
        .SYNOPSIS
        Creates a new Remedy ticket with the specified incident details.

        .DESCRIPTION
        This function creates a new Remedy ticket with the specified incident details. It connects to the Remedy server using the provided URL and login credentials, retrieves a token, and uses it to create a new incident. The incident details, such as customer name, summary, notes, impact, urgency, tenant, environment, template, and SolarWinds Alert URL, are provided as parameters to the function.

        .PARAMETER RemedyURL
        Specifies the URL of the Remedy server.

        .PARAMETER IncidentCustomer
        Specifies the customer name associated with the incident.

        .PARAMETER IncidentSummary
        Specifies the summary of the incident.

        .PARAMETER IncidentNote
        Specifies the note associated with the incident.

        .PARAMETER IncidentImpact
        Specifies the impact level of the incident.

        .PARAMETER IncidentUrgency
        Specifies the urgency level of the incident.

        .PARAMETER IncidentTenant
        Specifies the tenant of the incident.

        .PARAMETER IncidentEnvironment
        Specifies the environment of the incident.

        .PARAMETER IncidentTemplate
        Specifies the template of the incident.

        .PARAMETER AlertURL
        Specifies the SolarWinds Alert URL associated with the incident.

        .PARAMETER AlertKBA
        Specifies the SolarWinds KBA ID associated with the incident.

        .EXAMPLE
        New-Incident -RemedyURL "https://remedyserver.com" -IncidentCustomer "John Doe" -IncidentSummary "Application Down" -IncidentNote "The application is down." -IncidentImpact "High" -IncidentUrgency "Critical" -IncidentTenant "TenantA" -IncidentEnvironment "Production" -IncidentTemplate "Prod App Down" -AlertURL "https://solarwinds.com/alert/123" -AlertKBA "345"

        Creates a new Remedy ticket with the specified incident details, including the customer name, incident summary, note, impact, urgency, tenant, environment, template, SolarWinds Alert URL, and KBA ID.

        .NOTES
        This function requires the "swispowershell" module to use SWIS commands. It also requires a valid username and password to connect to the Remedy server. These credentials can be saved in an encrypted form using the "Export-Clixml" cmdlet and imported using the "Import-Clixml" cmdlet.
    #>

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
  
        # Create and encrypted UserName and Password with this command
        #$Credential = Get-Credential
        #$Credential | Export-Clixml -Path "C:\RemedyAPI\Credentials.xml"
  
        # Get the login credentials
        Write-Host "Getting credentials to login to Remedy"
        try {
            # Use this cred to running from SolarWinds
            $CredPath = "C:\RemedyAPI\Credentials.xml"

            # Other creds that can be used if logged in through CAPAM
            # $CredPath = "C:\RemedyAPI\Credentials.xml"
            
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
  
        Write-Host "Request to be sent to Remedy:"
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
#endregion Create New Incident 

#endregion functions

#region Main

try {
    Write-Host "Transforming received params..."
    $RemedyHost = (Test-ServerConnection -Verbose)
    # $RemedyCustomer = "Ryan"
    # $RemedySummary = "Test Summary"
    # $RemedyNotes = "Test Notes"
    # $SW_Severity = "Warning"
    # $RemedyEnvironment = "PRD"
    # $RemedyOrganization = "Org"
    # $RemedyAssignedGroup = "Grp"
    # $RemedyTenant = "tenant"
    # $SW_AlertID = "999"
    $Impact = (Set-Severity -Severity $SW_Severity).Impact
    $Urgency = (Set-Severity -Severity $SW_Severity).Urgency
    $IncEnv = Set-Environment -Environment $RemedyEnvironment
    $RemedyTemplate = Set-RemedyTemplate -RemedyOrganization $RemedyOrganization -RemedyAssignedGroup $RemedyAssignedGroup

    $IncidentNumber = New-Incident -RemedyURL $RemedyHost -IncCustomer $RemedyCustomer -IncSummary $RemedySummary -IncNotes $RemedyNotes -IncImpact $Impact -IncUrgency $Urgency -IncEnvironment $IncEnv -IncTenant $RemedyTenant -IncTemplate $RemedyTemplate -AltKBA $SW_ApplicationKBA -AltURL $SW_AlertURL

    Write-Host "Transformed data`nRemedy Host: $RemedyHost`nImpact/Urgnecy: $Impact/$Urgency`nEnvironment: $IncEnv`nTemplate: $RemedyTemplate`nIncident Number: $IncidentNumber"

    Update-SWAlert -SW_AlertID $SW_AlertID -IncidentNumber $IncidentNumber."Incident Number"
}
catch {

    Write-Error $_

    Stop-Transcript

    Move-Item -Path "$LogsDir\$filename.txt" -Destination $LogDirFailed

}
finally {

    Stop-Transcript
    
    if (Test-Path -Path "$LogsDir\$filename.txt") {
        Move-Item -Path "$LogsDir\$filename.txt" -Destination $LogDirSuccess
    }
    
}
#endregion Main