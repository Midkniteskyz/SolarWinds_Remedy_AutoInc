<#
.SYNOPSIS
Mutes or unmutes the alerts of a node in SolarWinds Orion.

.DESCRIPTION
This script connects to the SolarWinds Orion Server to mute or unmute the alerts of a node with the specified hostname.
You can specify only one switch: 'Mute' or 'Unmute'. If no switch is specified, the script returns the status of the node alert suppression.

.PARAMETER HostName
The hostname of the node whose alerts you want to mute or unmute.

.PARAMETER Mute
Mutes the alerts of the specified node.

.PARAMETER Unmute
Unmutes the alerts of the specified node.

.PARAMETER Status
Returns the status of the specified node alert suppression.

.EXAMPLE
.\RemedySolarWindsNodeSuppression.ps1 -HostName "node1" -Mute
Mutes the alerts of the node with hostname "node1".

.EXAMPLE
.\RemedySolarWindsNodeSuppression.ps1 -HostName "node1" -Unmute
Unmutes the alerts of the node with hostname "node1".

.EXAMPLE
.\RemedySolarWindsNodeSuppression.ps1 -HostName "node1" -Status
Returns the status of the alert suppression for the node with hostname "node1".

#>
param (
    [Parameter(Mandatory = $true)]
    [string]$HostName,
  
    [Parameter(Mandatory = $false)]
    [switch]$Mute,
  
    [Parameter(Mandatory = $false)]
    [switch]$Unmute,

    [Parameter(Mandatory = $false)]
    [switch]$Status
)

#region dependencies
Import-Module SwisPowerShell
#endregion dependencies

#region Connection details
$OrionServer = ''
$Username = 'Remedy'
$Password = ''

$SwisConnection = Connect-Swis -Hostname $OrionServer -Username $Username -Password $Password
#endregion Connection details

function Get-Node {
    [CmdletBinding()]
    param (
        [string]$SystemName
    )
    
    try {
        Write-Verbose "Querying the SolarWinds database for $SystemName"

        $query = Get-SwisData $SwisConnection "SELECT NodeID, SysName, URI FROM Orion.Nodes WHERE SysName = '$SystemName'"

        if (!$query) {
            throw 
        }
        
        return $query
    }
    catch {
        Write-Host "Device not monitored"
        Exit
    }
}


function Set-NodeSuppression {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$Mute,
        
        [Parameter(Mandatory = $false)]
        [switch]$Unmute,

        [Parameter(Mandatory = $true)]
        $Node
    )

    try {
        Write-Verbose "In Set-NodeSuppression"
        Write-Verbose $Node

        if ($Unmute) {
            Write-Verbose "Unmuting $($Node.Uri)"
            Invoke-SwisVerb $SwisConnection Orion.AlertSuppression ResumeAlerts @(, $Node.Uri, [DateTime]::UtcNow) | Out-Null
            Write-Output "Device Unmuted"
        }
        elseif ($Mute) {
            Write-Verbose "Muting $($Node.Uri)"
            Invoke-SwisVerb $SwisConnection Orion.AlertSuppression SuppressAlerts @(, $Node.Uri, [DateTime]::UtcNow) | Out-Null
            Write-Output "Device Muted"
        }
        else {
            throw "Unable to mute or unmute device."
        }
    }
    catch {
        throw $_.Exception.Message
    }
}

function Get-Suppression {
    [CmdletBinding()]
    param (
        $Node
    )
    
    try {
        $nodeUri = $Node.URI

        $AllSuppressions = Get-SwisData $SwisConnection "SELECT ID, EntityUri, SuppressFrom, SuppressUntil FROM Orion.AlertSuppression Where EntityUri = '$nodeUri'" 

        if (!$AllSuppressions) {
            Write-Output "Device currently Unmuted."
        }
        else {
            Write-Output "Device currently Muted."
        } 
    }
    catch {
        throw $_.Exception.Message
    }   
}

$VerbosePreference = "SilentlyContinue"

try {
    if (($Mute -eq $true) -and ($Unmute -eq $true)) {
        throw "Please specify only one switch: 'Mute' or 'Unmute'"
    }

    $Node = Get-Node -SystemName $HostName -Verbose

    if ($Mute) {
        Set-NodeSuppression -Node $Node -Mute -Verbose
    }
    elseif ($Unmute) {
        Set-NodeSuppression -Node $Node -Unmute -Verbose
    }
    elseif ($Status) {
        Get-Suppression -Node $Node
    }
}
catch {
    throw $_.Exception.Message
}
finally {

}
