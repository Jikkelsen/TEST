#region Public functions
function Get-DCTrackItem {

    #Returns all blades in DCTrack
    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $DCtrackCredential,

        [Parameter(Mandatory = $true)]
        [String]
        $DCTrackHost,

        [Parameter()]
        [String]
        $SubClass,

		[Parameter()]
		[String]
		$id

    )

    # Statics
    $headers = @{
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    $json = @{
        columns = @(
            @{
                name   = "cmbLocation"
            }
        )
        SelectedColumns = @(
            "tiName",
            "cmbLocation",
            "tiModel",
            "tiSerialNumber",
            "tiSubClass"
        )
    }

	if ($id) # Search for a single node with known ID
	{
		$Endpoint    = "/api/v2/dcimoperations/items/$id"
		$Uri         = $DCTrackHost + $Endpoint

		$params = @{
			Uri        = $Uri
			Method     = "GET"
			Credential = $DCtrackCredential
			SkipCertificateCheck = $true
		}

		$Response     = Invoke-WebRequest @params
        $responsedata = $Response.content | ConvertFrom-Json

        return $responsedata.item

	}
	else # Search for a subclass og devices
	{
		$Endpoint    = "/api/v2/quicksearch/items?pageNumber=0&pageSize=0"
		$Uri         = $DCTrackHost + $Endpoint

		# ----  Search for the device in the payload above. Stores the HTTP response in response
		$params = @{
			Uri        = $Uri
			Method     = "POST"
			Headers    = $headers
			body       = $JSON | ConvertTo-Json
			Credential = $DCtrackCredential
			SkipCertificateCheck = $true
		}

		$Response     = Invoke-WebRequest @params
		$responsedata = $Response.content | ConvertFrom-Json -AsHashtable

		return $responsedata.searchResults.items | Where-Object {$_.tiSubclass -eq $SubClass}
	}
}


function Install-DCTrackBlade {

    # Installs a completely new node
    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $DCtrackCredential,

        [Parameter(Mandatory = $true)]
        [String]
        $DCTrackHost,

        [Parameter(Mandatory = $true)]
        [String]
        $NodeID,

        [Parameter(Mandatory = $true)]
        [String]
        $Location,

        [Parameter(Mandatory = $true)]
        [String]
        $Cabinet,

        [Parameter(Mandatory = $true)]
        [String]
        $Chassis,

        [Parameter(Mandatory = $true)]
        [ValidateSet("1","2","3","4","5","6","7","8")]
        [String]
        $Slot,

        [Parameter()]
        [String]
        $Comment,

        [Parameter()]
        [String]
        $Requester = "Powershell_DCtrack_Script",

        [Parameter()]
        [ValidateSet("urgent","high","normal","low")]
        [String]
        $Priority = "high"
    )

    # Statics
    $Endpoint    = "/api/v2/dcimoperations/items/$NodeID"
    $Uri         = $DCTrackHost + $Endpoint

    $headers = @{
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    $json = @{
        cmbRequestType    = "Install Item"
        tiRequestComments = $Comment
        tiRequestPriority = $Priority
        tiRequestedby     = $Requester
        Location          = $Location
        cmbCabinet        = $Cabinet
        cmbChassis        = $Chassis
        radioChassisFace  = "front"
        cmbSlotPosition   = $Slot
    }

    $params = @{
        Uri        = $Uri
        Method     = "PUT"
        Headers    = $headers
        body       = $JSON | ConvertTo-Json
        Credential = $DCtrackCredential
        SkipCertificateCheck = $true
    }

    Invoke-WebRequest @params

}



function Move-DCtracknode {
    # Installs a completely new node
     param
     (
         [Parameter(Mandatory = $true)]
         [pscredential]
         $DCtrackCredential,

         [Parameter(Mandatory = $true)]
         [String]
         $DCTrackHost,

         [Parameter(Mandatory = $true)]
         [String]
         $NodeID,

         [Parameter(Mandatory = $true)]
         [String]
         $Location,

         [Parameter()]
         [String]
         $Cabinet,

         [Parameter()]
         [String]
         $Chassis,

         [Parameter()]
         [ValidateSet("1","2","3","4","5","6","7","8")]
         [String]
         $Slot,

         [Parameter()]
         [String]
         $Comment,

         [Parameter()]
         [String]
         $Requester = "Powershell_DCtrack_Script"
     )

     # Statics
     $Endpoint    = "/api/v2/dcimoperations/items/$($NodeID)?returnDetails=true"
     $Uri         = $DCTrackHost + $Endpoint

     $headers = @{
         "Content-Type"  = "application/json"
         "Accept"        = "application/json"
     }

     if ($slot) # User specified a slot; enrich the rest of the chasis specific details
     {
         $json = @{
             tiRequestComments = $Comment
             tiRequestedby     = $Requester
             cmbLocation       = $Location
             cmbCabinet        = $Cabinet
             cmbChassis        = $Chassis
             radioChassisFace  = "front"
             cmbSlotPosition   = $Slot
         }
     }
     else # Move only to location
     {
         $json = @{
         tiRequestComments = $Comment
         tiRequestedby     = $Requester
         cmbLocation       = $Location
         }
     }

     $params = @{
         Uri        = $Uri
         Method     = "PUT"
         Headers    = $headers
         body       = $JSON | ConvertTo-Json
         Credential = $DCtrackCredential
         SkipCertificateCheck = $true
     }

     Invoke-WebRequest @params

 }

function New-DCtrackBlade {
    # Installs a completely new node
    # documentation: https://www.sunbirddcim.com/help/dcTrack/v810/API/en/Default.htm#APIGuide/v2_Create_a_New_Item.htm?TocPath=Using%2520the%2520REST%2520API%2520to%2520Manage%2520Items%2520%255Bv2%255D%257C_____1

    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $DCtrackCredential,

        [Parameter(Mandatory = $true)]
        [String]
        $DCTrackHost,

        [Parameter(Mandatory = $true)]
        [String]
        $Model,

        [Parameter(Mandatory = $true)]
        [String]
        $SerialNumber,

        [Parameter()]
        [String]
        $Name = $SerialNumber,

        [Parameter(Mandatory = $true)]
        [String]
        $Location,

        [Parameter(Mandatory = $true)]
        [String]
        $Cabinet,

        [Parameter(Mandatory = $true)]
        [String]
        $Chassis,

        [Parameter(Mandatory = $true)]
        [ValidateSet("1","2","3","4","5","6","7","8")]
        [String]
        $Slot,

        [Parameter()]
        [String]
        $Requester = "Powershell_DCtrack_Script",

        [Parameter()]
        [ValidateSet("urgent","high","normal","low")]
        [String]
        $Priority = "high"
    )

    # Statics
    $Endpoint    = "/api/v2/dcimoperations/items?returnDetails=true"
    $Uri         = $DCTrackHost + $Endpoint

    $headers = @{
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    $json = @{
        tiRequestComments = $Comment
        tiRequestPriority = $Priority
        tiRequestedby     = $Requester
        tiSerialNumber    = $SerialNumber
        tiName            = $name
        cmbLocation       = $Location
        cmbCabinet        = $Cabinet
        cmbChassis        = $Chassis
        cmbSlotPosition   = $Slot
        cmbMake           = "Cisco Systems"
        cmbStatus         = "Installed"
        cmbModel          = $Model
        radioChassisFace  = "front"
    }

    $params = @{
        Uri        = $Uri
        Method     = "POST"
        Headers    = $headers
        body       = $JSON | ConvertTo-Json
        Credential = $DCtrackCredential
        SkipCertificateCheck = $true
    }

    Invoke-WebRequest @params

}

function Remove-DCTracknode {

    # Comment JVM 31-10-2023: Original FunctionName is "Delete-DCTrackNode". This uses an unapproved verb, and have been changed to "remove"
    # Completely deletes a node
    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $DCtrackCredential,

        [Parameter(Mandatory = $true)]
        [String]
        $DCTrackHost,

        [Parameter(Mandatory = $true)]
        [String]
        $NodeID
    )

    # Statics
    $Endpoint    = "/api/v2/dcimoperations/items/$NodeID"
    $Uri         = $DCTrackHost + $Endpoint

    $headers = @{
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    $params = @{
        Uri        = $Uri
        Method     = "DELETE"
        Headers    = $headers
        Credential = $DCtrackCredential
        SkipCertificateCheck = $true
    }

    Invoke-WebRequest @params

}


function Unregister-DCTrackNode {

    # Note JVM 31-10-2023: Original function name is "Decomission-DCTrackNode". This uses an unapproved verb, and has been changed to unregister

    # Puts a DCTrack item in planned storage
    param
    (
        [Parameter(Mandatory)]
        [pscredential]
        $DCtrackCredential,

        [Parameter(Mandatory)]
        [String]
        $DCTrackHost,

        [Parameter(Mandatory)]
        [String]
        $NodeID,

        [Parameter()]
        [String]
        $Comment,

        [Parameter()]
        [String]
        $Requester = "Powershell_DCtrack_Script",

        [Parameter()]
        [ValidateSet("urgent","high","normal","low")]
        [String]
        $Priority = "high"
    )

    # Statics
    $Endpoint    = "/api/v2/dcimoperations/items/$NodeID"
    $Uri         = $DCTrackHost + $Endpoint

    $headers = @{
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
    }

    $json = @{
        cmbLocation       = "NETCOMPANY > GR17-LAGER"
        cmbRequestType    = "Decommission Item to Storage"
        tiRequestComments = $Comment
        tiRequestPriority = $Priority
        tiRequestedby     = $Requester
    }

    $params = @{
        Uri        = $Uri
        Method     = "PUT"
        Headers    = $headers
        body       = $JSON | ConvertTo-Json
        Credential = $DCtrackCredential
        SkipCertificateCheck = $true
    }

    Invoke-WebRequest @params

}

Function Add-ComputeDhcpReservation
{
    Param
    (
        [Parameter(Mandatory)]
        [String]
        $IPAddress,

        [Parameter(Mandatory)]
        [String]
        $MacAddress,

        [Parameter(Mandatory)]
        [String]
        $ReservationName,

        [Parameter(Mandatory)]
        [String]
        $ScopeId,

        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]
        $RemoteSession
    )



    # Make sure the mac address is in desired format
    $CleanedMacAddress = ($MacAddress -replace '\W' -replace '(..)','$1-').TrimEnd('-')

    # Check if the reservation is available
    $MACReservation = Get-ComputeDhcpReservation -RemoteSession $RemoteSession -ScopeId $ScopeId -MacAddress $CleanedMacAddress
    $IPReservation  = Get-ComputeDhcpReservation -RemoteSession $RemoteSession -ScopeId $ScopeId -IPAddress  $IPAddress

    if ($MACReservation)
    {
        # There is a reservation for the requested MAC; but its INACTIVE
        if ($MACReservation.AddressState -eq "InactiveReservation")
        {
            Write-Host "Inactive reservation detected on MAC:$($MACReservation.ClientId), tied to IP:$($MACReservation.ipaddress)"
            Write-Host "Removing inactive Reservation"

            Remove-ComputeDhcpReservation -RemoteSession $RemoteSession -ScopeId $ScopeId -MacAddress $MACReservation.ClientId
        }

        # There is a reservation for the requested MAC; but its ACTIVE
        if ($MACReservation.AddressState -eq "ActiveReservation")
        {

        }
    }

    if ($null -eq $IPReservation)
    {
        Invoke-Command -Session $RemoteSession -ScriptBlock {

            $AddParams = @{
                ScopeId         = $ScopeId
                IPAddress       = $IPAddress
                RemoteSession   = $RemoteSession.ComputerName
                MacAddress      = $MacAddress
                ReservationName = $ReservationName
            }
            Write-Host "Attemting IP Reservation with the following information:"
            $AddParams

            Add-DhcpServerv4Reservation -ScopeId $Using:ScopeId -IPAddress $Using:IPAddress -ClientId $Using:CleanedMacAddress -Name $Using:ReservationName -Type Both
            Set-DhcpServerv4OptionValue -ReservedIP $Using:IPAddress -OptionId 12 -Value $Using:ReservationName
        }
    }
    else
    {
        Write-Error -Message "IPAddress was unavailable for reservation."
    }
}


<#
   _____                            _            _____  _                _____
  / ____|                          | |          |  __ \| |              / ____|
 | |     ___  _ __  _ __   ___  ___| |_   ______| |  | | |__   ___ _ __| (___   ___ _ ____   _____ _ __
 | |    / _ \| '_ \| '_ \ / _ \/ __| __| |______| |  | | '_ \ / __| '_ \\___ \ / _ \ '__\ \ / / _ \ '__|
 | |___| (_) | | | | | | |  __/ (__| |_         | |__| | | | | (__| |_) |___) |  __/ |   \ V /  __/ |
  \_____\___/|_| |_|_| |_|\___|\___|\__|        |_____/|_| |_|\___| .__/_____/ \___|_|    \_/ \___|_|
                                                                  | |
                                                                  |_|
#>
Function Connect-ComputeDhcpServer
{
    Param
    (
        [Parameter(Mandatory)]
        [pscredential]
        $DhcpCredential,

        [Parameter()]
        [String]
        $DhcpServer = "ncop-vmwp-dhc01.prod.vmw.ncop.nchosting.dk"
    )

    try
    {
        Write-host "Establishing remote session on $dhcpServer ... " -NoNewLine

        # Prepare parameters for the connection. All custom params are needed specifically to work within NC confines
        $RemoteParams = @{
            ComputerName  = $DhcpServer
            Credential    = $DhcpCredential
            Port          = "5986"
            UseSSL        = $true
            SessionOption = @{
                SkipRevocationCheck = $true
                SkipCACheck         = $true
                SkipCNCheck         = $true
                ProxyAccessType     = "NoProxyServer"
            }
        }

        # Note JVM 29-11-2023: If you [void] this line, nothing will work
        New-PSSession @RemoteParams

        Write-host "OK"

    }
    catch
    {
        Write-Error "FAIL!"
        Write-host "Could not establish session on $dhcpServer"
        throw
    }
}
Function Disconnect-ComputeDhcpServer
{
    Param
    (
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]
        $RemoteSession
    )

    try {
        Write-Host "Removing session $($RemoteSession.ComputerName) ... " -NoNewline
        $RemoteSession | Remove-PSSession
        Write-Host "OK"

    }
    catch
    {
        Write-host "Could not remove Remote Session"
        throw
    }

}
Function Get-ComputeDhcpReservation
{
    Param
    (
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]
        $RemoteSession,

        [Parameter(Mandatory)]
        [ipaddress]
        $ScopeId,

        [Parameter()]
        [ipaddress]
        $IPAddress,

        [Parameter()]
        [String]
        $MacAddress,

        [Parameter()]
        [String]
        $ReservationName
    )

    Write-Verbose 'Getting information from $($RemoteSession.ComputerName)'
    Write-Verbose "using IPAddress:`"$IPAddress`", MacAddress:`"$MacAddress`", ReservationName:`"$ReservationName`""

    try
    {
        $Reservation = Invoke-Command -Session $RemoteSession -ScriptBlock {
            Get-DhcpServerv4Reservation -ScopeId $Using:ScopeId | Where-Object {
                $_.ClientId  -eq $Using:MacAddress      -or
                $_.Name      -eq $Using:ReservationName -or
                $_.IPAddress -eq $Using:IPAddress
            }
        }
    }
    catch
    {
        Write-Error "Could not invoke command"
        throw
    }

    if ($null -eq $Reservation)
    {
        Write-Host "The requested reservation does not exist"
        return $null
    }
    else
    {
        return $Reservation
    }
}

Function Remove-ComputeDhcpReservation
{
    Param
    (
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]
        $RemoteSession,

        [Parameter()]
        [ipaddress]
        $ScopeId,

        [Parameter()]
        [ipaddress]
        $IPAddress,

        [Parameter()]
        [String]
        $MacAddress
    )

    Write-Host "Attempting to remove reservation from $($RemoteSession.ComputerName)"
    Write-Host "using IPAddress:`"$IPAddress`", MacAddress:`"$MacAddress`", ReservationName:`"$ReservationName`""

    try
    {
        if ($IPAddress)
        {
            Invoke-Command -Session $RemoteSession -ScriptBlock {

                Remove-DhcpServerv4Reservation -IPAddress $Using:IPAddress
            }
        }
        elseif ($MacAddress)
        {
            Invoke-Command -Session $RemoteSession -ScriptBlock {

                Remove-DhcpServerv4Reservation -ScopeId $Using:ScopeId -ClientId $Using:MacAddress
            }
        }
    }
    catch
    {
        Write-Error "Could not invoke command"
        throw
    }
}
#Requires -Modules SwisPowerShell
<#
              _     _         ____       _             _   _           _
     /\      | |   | |       / __ \     (_)           | \ | |         | |
    /  \   __| | __| |______| |  | |_ __ _  ___  _ __ |  \| | ___   __| | ___
   / /\ \ / _` |/ _` |______| |  | | '__| |/ _ \| '_ \| . ` |/ _ \ / _` |/ _ \
  / ____ \ (_| | (_| |      | |__| | |  | | (_) | | | | |\  | (_) | (_| |  __/
 /_/    \_\__,_|\__,_|       \____/|_|  |_|\___/|_| |_|_| \_|\___/ \__,_|\___|

#>
#------------------------------------------------| HELP |------------------------------------------------#
<#
    .SYNOPSIS
        This script will add a node in Orion
    .PARAMETER ToolkitCredential
        Creds to import for authorization on toolkit.
    .PARAMETER SwisConnection
        Connection string gathered from Connect-Orion function
#>
Function Add-OrionNode
{
    #region---------------------------------------| PARAMETERS |---------------------------------------------#
    Param
    (
        [Parameter(Mandatory = $true)]
        [SolarWinds.InformationService.Contract2.InfoServiceProxy]
        $SwisConnection,

        [Parameter(Mandatory = $true)]
        [String]
        $NodeName,

        [Parameter(Mandatory = $true)]
        [String]
        $NodeIPAddress,

        [Parameter()]
        [String]
        $PollingEngineID = $(1..5 | Get-Random),

        [Parameter()]
        [String]
        $SNMPV2Community = "Fors23",

        [Parameter()]
        [bool]
        $Unmanaged = $true
    )
    #endregion

    #region---------------------------------------| CREATE NODE |--------------------------------------------#
    If ($null -eq $SNMPV2Community)
    {
        Write-Host "no SNMP Community supplied"
        break
    }
    Else
    {
        $NewSNMPV2NodeProperties = @{
            IPAddress     = $NodeIPAddress;
            EngineID      = $PollingEngineID;
            Caption       = $NodeName;
            ObjectSubType = "SNMP";
            Community     = $SNMPV2Community;
            SNMPVersion   = "2";
            DNS           = "$NodeName";
            SysName       = "$NodeName";
            SysObjectID   = "1.3.6.1.4.1.6876.4.1";
            }

        try
        {
            Write-Host "Creating node ... " -NoNewline
            [void]::(New-SwisObject $SwisConnection -EntityType "Orion.Nodes" -Properties $NewSNMPV2NodeProperties)
            Write-Host "OK"
        }
        catch
        {
            Write-Host "FAIL!" -BackgroundColor Red
            Write-Host "Could not create node"
            throw
        }

    }
    #endregion


    #region---------------------------------------| ADD POLLERS |--------------------------------------------#
    Do
    {
        $CreatedNode = Get-OrionNode -SwisConnection $SwisConnection -NodeIP $NodeIPAddress
        Start-Sleep -Seconds 1
    }
    Until ($Null -ne $CreatedNode)

    $NestedPollers = Get-NCNestedPollers

    Write-Host "Setting Custom Pollers ... " -NoNewline
    foreach ($Type in $NestedPollers)
    {
        # Build Poller
        $poller = @{
            NetObject     = "N:"+$CreatedNode.NodeID;
            NetObjectType = "N";
            NetObjectID   = $CreatedNode.NodeID;
            PollerType    = $Type.PollerType;
            Enabled       = $Type.Enabled;
        }

        # Add Poller
        try
        {
            Write-Verbose "`tAdding $($Poller.PollerType) "

            $Properties = @{
                "SwisConnection" = $SwisConnection
                "EntityType"     = "orion.pollers"
                "Properties"     = $poller
            }
            [void]::(New-SwisObject @Properties)

            Write-Verbose "OK"
        }
        catch
        {
            Write-Host "FAIL!"
            Write-Host "Could not add poller"
            throw
        }
    }
    Write-Host "OK"
    #endregion


    #region----------------------------------| ADD CUSTOM PROPERTIES |---------------------------------------#
    $AdditionalProperties = @{
        NC_Customer    = "Netcompany";
        NC_Category    = "Server Hosting Services";
        NC_Contact     = "OPE Ledelsesvagt";
        NC_Device      = "Server";
        NC_Environment = "PROD";
        NC_Solution    = "None";
        NC_ToolkitID   = "NCINFRA";
    }


    try
    {
        Write-Host "Updating custom properties ... " -NoNewline
        $CustomPropertiesUrl = $CreatedNode.uri+'/CustomProperties'
        $CustomPropertiesUrl = $CreatedNode.uri+'/CustomProperties'

        $Params = @{
            SwisConnection = $SwisConnection
            Uri            = $CustomPropertiesUrl
            Properties     = $AdditionalProperties
        }

        [Void]::(Set-SwisObject @Params)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not set custom properties"
        $Error[0]
        throw
    }
    #endregion
}
#-------------------------------------------------| END |------------------------------------------------#
#Requires -Modules SwisPowerShell
<#
   _____                            _           ____       _
  / ____|                          | |         / __ \     (_)
 | |     ___  _ __  _ __   ___  ___| |_ ______| |  | |_ __ _  ___  _ __
 | |    / _ \| '_ \| '_ \ / _ \/ __| __|______| |  | | '__| |/ _ \| '_ \
 | |___| (_) | | | | | | |  __/ (__| |_       | |__| | |  | | (_) | | | |
  \_____\___/|_| |_|_| |_|\___|\___|\__|       \____/|_|  |_|\___/|_| |_|


#>
#region------------------------------------------| HELP |------------------------------------------------#
<#
    .SYNOPSIS
        Establishes connection to orion. Returns the SwisConnection
    .PARAMETER OrionCredential
        Credential file to connect to Orion
    .PARAMETER OrionHostname
        Hostname of current Orion installation to point to
#>
#endregion

#---------------------------------------------| PARAMETERS |---------------------------------------------#
Function Connect-Orion
{
    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $OrionCredential,

        [Parameter()]
        [string]
        $OrionHostname = "ncop-sol01.noc02.nchosting.dk"
    )

    #-----------------------------------------| PROGRAM LOGIC |-------------------------------------------#

    # Manually set Security protocol
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

    try
    {
        Write-Host "Connecting to Orion ... " -NoNewline
        $SwisConnection = Connect-Swis -Hostname $OrionHostname -Credential $OrionCredential
        Write-Host "OK"

        return $SwisConnection
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not connect to Orion"
        throw
    }

}
#-------------------------------------------------| END |------------------------------------------------#


#Requires -Modules SwisPowerShell
#---------------------------------------------| PARAMETERS |---------------------------------------------#
Function Get-AllOrionNodes
{
    param
    (
        [Parameter()]
        [SolarWinds.InformationService.Contract2.InfoServiceProxy]
        $SwisConnection
    )

    #-----------------------------------------| PROGRAM LOGIC |-------------------------------------------#

    Write-Host "Looking up all ESX nodes in Orion ... " -NoNewline
    #Pull data from DB
    $params = @{
        Query          = "SELECT Caption,NodeID,Unmanaged,IP_Address,URI FROM Orion.Nodes WHERE Caption LIKE '%esx%';"
        SwisConnection = $SwisConnection
    }
    $Data = Get-SwisData @params
    Write-Host "OK"

    return $Data
}
#-------------------------------------------------| END |------------------------------------------------#
<#
   _____      _          _   _  _____ _   _           _           _ _____      _ _
  / ____|    | |        | \ | |/ ____| \ | |         | |         | |  __ \    | | |
 | |  __  ___| |_ ______|  \| | |    |  \| | ___  ___| |_ ___  __| | |__) |__ | | | ___ _ __ ___
 | | |_ |/ _ \ __|______| . ` | |    | . ` |/ _ \/ __| __/ _ \/ _` |  ___/ _ \| | |/ _ \ '__/ __|
 | |__| |  __/ |_       | |\  | |____| |\  |  __/\__ \ ||  __/ (_| | |  | (_) | | |  __/ |  \__ \
  \_____|\___|\__|      |_| \_|\_____|_| \_|\___||___/\__\___|\__,_|_|   \___/|_|_|\___|_|  |___/


#>
# Comment JVM 09-10-2023
# This function has no inputs, as all its paramaters are harddtyped - they have not changed for years
# Update this function when time arises
Function Get-NCNestedPollers {

    # Prepare Pollers
    $NeededPollers = @()

    # Create pollers one by one
    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486273"
        'PollerType' = "N.ResponseTime.SNMP.Native"
        'Enabled'    = "False"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486274"
        'PollerType' = "N.Status.SNMP.Native"
        'Enabled'    = "False"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486271"
        'PollerType' = "N.ResponseTime.ICMP.Native"
        'Enabled'    = "True"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486272"
        'PollerType' = "N.Status.ICMP.Native"
        'Enabled'    = "True"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486275"
        'PollerType' = "N.Details.SNMP.Generic"
        'Enabled'    = "True"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486276"
        'PollerType' = "N.Uptime.SNMP.Generic"
        'Enabled'    = "True"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486277"
        'PollerType' = "N.Cpu.SNMP.HrProcessorLoad"
        'Enabled'    = "True"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486278"
        'PollerType' = "N.Memory.SNMP.HrStorage"
        'Enabled'    = "True"
    }

    $NeededPollers += [PSCustomObject]@{
        'PollerId'   = "486279"
        'PollerType' = "N.AssetInventory.Snmp.Generic"
        'Enabled'    = "True"
    }

    # Write to Console
    return $NeededPollers | Sort-Object PollerID
}
#Requires -Modules SwisPowerShell
#---------------------------------------------| PARAMETERS |---------------------------------------------#
Function Get-OrionNode
{
    param
    (
        [Parameter()]
        [SolarWinds.InformationService.Contract2.InfoServiceProxy]
        $SwisConnection,

        [Parameter(Mandatory = $true)]
        [string]
        $NodeIP,

        [Parameter()]
        [bool]
        $Quiet = $false
    )

    #-----------------------------------------| PROGRAM LOGIC |-------------------------------------------#

    try
    {
        if (-not $Quiet)
        {
            Write-Host "Looking up node ... " -NoNewline
        }

        # Pull data from Orion Database
        $params = @{
            Query          = "SELECT Caption,NodeID,Unmanaged,IP_Address,URI FROM Orion.Nodes WHERE IP_Address = '" + $NodeIP + "';"
            SwisConnection = $SwisConnection
        }
        $Data = Get-SwisData @params

        if (-not $Quiet)
        {
            Write-Host "OK"
        }

        return $Data
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not get Orion node"
        $Error[0]
        throw
    }
}
#-------------------------------------------------| END |------------------------------------------------#
#region----------------------------------------| UNMANAGE |----------------------------------------------#
<#
    _____      _           ____       _             _   _           _       _____ _        _
  / ____|    | |         / __ \     (_)           | \ | |         | |     / ____| |      | |
 | (___   ___| |_ ______| |  | |_ __ _  ___  _ __ |  \| | ___   __| | ___| (___ | |_ __ _| |_ ___
  \___ \ / _ \ __|______| |  | | '__| |/ _ \| '_ \| . ` |/ _ \ / _` |/ _ \\___ \| __/ _` | __/ _ \
  ____) |  __/ |_       | |__| | |  | | (_) | | | | |\  | (_) | (_| |  __/____) | || (_| | ||  __/
 |_____/ \___|\__|       \____/|_|  |_|\___/|_| |_|_| \_|\___/ \__,_|\___|_____/ \__\__,_|\__\___|

#>
Function Set-OrionNodeState
{
    param (
        [Parameter(Mandatory)]
        [string]
        $NodeIP,

        [Parameter(Mandatory)]
        [pscredential]
        $OrionCredential,

        [Parameter(Mandatory)]
        [String]
        [ValidateSet("Unmanaged","Managed")]
        $State
    )

    #TODO:
    # Job Functions should not themselves log into Orion. Change input from credential to finished connection and remove connect-orion dependency
    # Manually set Security protocol
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    # Put connection to Orion in variable for future use
    $SwisConnection = Connect-Orion -OrionCredential $OrionCredential

    if ($State -eq "Unmanaged")
    {
        #Pull data from DB
        $params = @{
            NodeIP         = $NodeIP
            SwisConnection = $SwisConnection
        }
        $Data = Get-OrionNode @params

        # Exit if no return data
        if ($null -eq $data)
        {
            Write-Host "Could not find host, please add it!"
        }
        # If managed: Unmanage
        elseif ($Data.Unmanaged -eq $false)
        {
            # Build longer variables for hashtable
            $Hostname    = 'ncop-sol01.noc02.nchosting.dk'
            $UnmanageUrl = "https://$($Hostname):17778/SolarWinds/InformationService/v3/Json/Invoke/Orion.Nodes/Unmanage"
            $Now         = (Get-Date).AddHours(-2).ToString("MM'/'dd'/'yyyy hh:mm")
            $ThirtyDays  = (Get-Date).AddDays(30).ToString("MM'/'dd'/'yyyy hh:mm")

            # Build paramaters for re-management
            $UnmanageParams = @{
                uri                  = $UnmanageUrl
                Credential           = $OrionCredential
                UseBasicParsing      = $True
                Method               = "Post"
                Body                 = "[`"N:$($Data.NodeID)`",`"$($Now)`",`"$($ThirtyDays)`",`"false`"]"
                ContentType          = "application/Json"
                SkipCertificateCheck = $True
            }

            # Send the unmanage request to Orion server
            try
            {
                Write-Host "Unmanaging node ..." -NoNewline

                [void]::(Invoke-WebRequest @UnmanageParams -SkipCertificateCheck)
            }
            catch
            {
                Write-Host "FAIL!"
                throw
            }

            # Refresh the Data, and wait for desired result
            $Data    = Get-OrionNode @params -Quiet:$true
            $Counter = 0

            if ($Data.Unmanaged -eq $false)
            {
                # Continue to try, until successfully unmanaged
                Do
                {
                    # Add another "dot" to the console output
                    Write-Host "." -NoNewline

                    # Send the unmanage request
                    [void]::(Invoke-WebRequest @UnmanageParams -SkipCertificateCheck)

                    # Refresh the Data
                    Start-Sleep -Seconds 2
                    $Data = Get-OrionNode @params -Quiet:$true

                    # Exit upon too many retries
                    $Counter++
                    if ($Counter -eq 30)
                    {
                        Write-Host "FAIL!"
                        Write-Host "Node could not be unmanaged"
                        throw
                    }
                }
                Until($Data.Unmanaged -eq $true)
            }
            Write-Host " OK"
        }
        else
        {
            Write-Host "Node Already Unmanaged" -BackgroundColor "Yellow" -ForegroundColor "Black"
        }
    }
    elseif ($State -eq "Managed")
    {
        #Pull data from DB
        $params = @{
            NodeIP         = $NodeIP
            SwisConnection = $SwisConnection
        }
        $Data = Get-OrionNode @params

        # If unmanaged: Manage
        if ($Data.Unmanaged -eq $true)
        {
            Write-Host "Managing node ..." -nonewline

            # Build longer variables for hashtable
            $Hostname    = 'ncop-sol01.noc02.nchosting.dk'
            $RemanageUrl = "https://$($Hostname):17778/SolarWinds/InformationService/v3/Json/Invoke/Orion.Nodes/Remanage"
            $Now         = (Get-Date).AddHours(-2).ToString("MM'/'dd'/'yyyy hh:mm")
            $Forever     = (Get-Date).AddDays(30).ToString("MM'/'dd'/'yyyy hh:mm")

            # Build paramaters for re-management
            $RemanageParams = @{
                uri                  = $RemanageUrl
                Credential           = $OrionCredential
                UseBasicParsing      = $True
                Method               = "Post"
                Body                 = "[`"N:$($Data.NodeID)`",`"$($Now)`",`"$($Forever)`",`"false`"]"
                ContentType          = "application/Json"
                #SkipCertificateCheck = $True
            }

            # Remanaging
            $RemanageValidation = Invoke-WebRequest @RemanageParams -SkipCertificateCheck

            # Test if change went through
            if ($RemanageValidation.StatusDescription -ne 'OK')
            {
                # Continue to try, until successfully managed
                Do
                {
                    # Put another dot on the loading line
                    Write-Host "."
                    $RemanageValidation = Invoke-WebRequest $RemanageParams -SkipCertificateCheck
                }
                Until($RemanageValidation.StatusDescription -eq 'OK')
            }
            Write-Host " OK"
        }
        else
        {
            Write-Host "Node is already managed by Orion"
        }

    }
}

function Get-FormattedTime
{
    return (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
}

# This function is used, if no imput is supplied via Jenkins
Function Import-LocalKeyPath
{
    param
    (
        [Parameter()]
        [string]
        $PathApiKeyID = "$HOME\documents\ApiKeyid.txt",

        [Parameter()]
        [string]
        $PathApiKey   = "$HOME\documents\ApiKey.txt"
    )

    # Import Key file
    if (Test-Path -Path $PathApiKey)
    {
        Write-Host "Importing API Key locally ... " -NoNewline
        $ApiKey        = $PathApiKey
        $ApiKeyContent = Get-content $PathApiKey
        Write-Host "OK"
    }
    else
    {
        Write-Host "FAIL"
        Write-Host "$PathApiKey is not a valid path"
    }

    # Import ID file
    if (Test-Path -Path $PathApiKeyID)
    {
        Write-host "Importing API ID Locally  ... " -NoNewline
        $ApiKeyID = Get-content $PathApiKeyID
        Write-Host "OK"
    }
    else
    {
        Write-Host "FAIL"
        Write-Host "$PathApiKeyID is not a valid path"
    }

    # API key must have 27 lines
    if ($ApiKeyContent.Length -ne 27 -or $null -eq $ApiKey)
    {
        Write-Host "Errornous API Key!"
        Throw
    }

    # ID File must be 74 chars
    if ($ApiKeyID.Length -ne 74)
    {
        Write-Host "Errornous API Key ID!"
        Throw
    }

    # Return the two values
    return $ApiKey,$ApiKeyID
}

# This function is used, if no imput is supplied via Jenkins
Function Import-LocalKeyString
{
    param
    (
        [Parameter()]
        [string]
        $PathApiKeyID = "$HOME\documents\ApiKeyid.txt",

        [Parameter()]
        [string]
        $PathApiKey   = "$HOME\documents\ApiKey.txt"
    )

    # Import Key file
    if (Test-Path -Path $PathApiKey)
    {
        Write-Host "Importing API Key locally ... " -NoNewline
        $ApiKey        = $PathApiKey
        $ApiKeyContent = (Get-content $PathApiKey) -join "`n"
        Write-Host "OK"
    }
    else
    {
        Write-Host "FAIL"
        Write-Host "$PathApiKey is not a valid path"
    }

    # Import ID file
    if (Test-Path -Path $PathApiKeyID)
    {
        Write-host "Importing API ID Locally  ... " -NoNewline
        $ApiKeyID = Get-content $PathApiKeyID
        Write-Host "OK"
    }
    else
    {
        Write-Host "FAIL"
        Write-Host "$PathApiKeyID is not a valid path"
    }

    # API key must have 1674 chars
    if ($ApiKeyContent.Length -ne 1674 -or $null -eq $ApiKey)
    {
        Write-Host "Errornous API Key!"
        Throw
    }

    # ID File must be 74 chars
    if ($ApiKeyID.Length -ne 74)
    {
        Write-Host "Errornous API Key ID!"
        Throw
    }

    # Return the two values
    return $ApiKeyContent,$ApiKeyID
}

<#

  _   _                          _____                                           _
 | \ | |                        |  __ \                                         | |
 |  \| |  ___ __      __ ______ | |__) |__ _  ___  ___ __      __ ___   _ __  __| |
 | . ` | / _ \\ \ /\ / /|______||  ___// _` |/ __|/ __|\ \ /\ / // _ \ | '__|/ _` |
 | |\  ||  __/ \ V  V /         | |   | (_| |\__ \\__ \ \ V  V /| (_) || |  | (_| |
 |_| \_| \___|  \_/\_/          |_|    \__,_||___/|___/  \_/\_/  \___/ |_|   \__,_|


#>
# -------- HELP --------
<#
.Credit
    ALL credit goes to MAGC.
    Code commented and documented by JVM
.SYNOPSIS
    This script will generate a new secure password string or credentialobject
.PARAMETER AsString
    Specify if return object should be plaintext string
.PARAMETER Length
    Specifies the length of the
.PARAMETER ForbiddenChars
    Allows user to make specific chars forbiden
.PARAMETER MinLowerCaseChars
    Set minimum amount of required lower case chars
.PARAMETER MinUpperCaseChars
    Set minimum amount of upper case chars
.PARAMETER MinDigits
    Set minimum amount of digits
.PARAMETER MinSpecialChars
    Set minimum amount of special chars required
#>
function New-Password
{
    [CmdletBinding(PositionalBinding = $false)]
    [Alias('np')]
    [OutputType([securestring], [string])]

    #--------------------------------------------| PARAMETERS |--------------------------------------------#
    Param
    (
        [Parameter()]
        [switch]
        $AsString,

        [Parameter()]
        [ValidateRange(8, [int]::MaxValue)]
        [Int]
        $Length = 40,

        [Parameter()]
        [Alias('DisallowedChars')]
        [ArgumentCompleter(
            {
                param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
                switch ($wordToComplete -replace "`"|'")
                {
                    { 'Lowercase' -like "$_*" }
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            'abcdefghijklmnopqrstuvwxyz'.ToCharArray().ForEach({ "'$_'" }) -join ',',
                            'Lowercase',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Lowercase'
                        )
                    }
                    { 'Uppercase' -like "$_*" }
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray().ForEach({ "'$_'" }) -join ',',
                            'Uppercase',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Uppercase'
                        )
                    }
                    { 'Digits' -like "$_*" }
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            '1234567890'.ToCharArray().ForEach({ "'$_'" }) -join ',',
                            'Digits',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Digits'
                        )
                    }
                    { 'Special' -like "$_*" }
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            '/*!\"$%()=?{[]}+#-.,<_:;>~|@'.ToCharArray().ForEach({ "'$_'" }) -join ',',
                            'Special',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Special'
                        )
                    }
                    { 'Ambiguous' -like "$_*" }
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            'IlOo0'.ToCharArray().ForEach({ "'$_'" }) -join ',',
                            'Ambiguous',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Ambiguous'
                        )
                    }
                }
            }
        )]
        [char[]]
        $ForbiddenChars,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [Int]
        $MinLowercaseChars = 2,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [Int]
        $MinUppercaseChars = 2,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [Int]
        $MinDigits = 2,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [Int]
        $MinSpecialChars = 2
    )
    #---------------------------------------------| CHECK INPUT |--------------------------------------------#
    begin
    {
        # Start out by building $AllAllowedChars variable. This is all subvariables concatinated, where no forbidden chars are included
        [char[]]$AllAllowedChars = @(
            ([char[]]$AllowedLowercase = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray().Where({ $_ -cnotin $ForbiddenChars }))
            ([char[]]$AllowedUppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray().Where({ $_ -cnotin $ForbiddenChars }))
            ([char[]]$AllowedDigits = '1234567890'.ToCharArray().Where({ $_ -notin $ForbiddenChars }))
            ([char[]]$AllowedSpecial = '/*!\"$%()=?{[]}+#-.,<_:;>~|@'.ToCharArray().Where({ $_ -notin $ForbiddenChars }))
        )
        # FillerCharCount refers to the amount of characters not dictated by the required minimum of each type
        [int]$FillerCharCount = $Length - ($MinLowercaseChars + $MinUppercaseChars + $MinDigits + $MinSpecialChars)

        # For all if statements below, throw erorr if minimum requirements not met.
        if ($FillerCharCount -lt 0)
        {
            throw 'The specified length is less than the sum of the minimum character counts.'
        }
        if ($AllowedLowercase.Count -lt 1 -and $MinLowercaseChars -gt 0)
        {
            throw 'There are not enough allowed lowercase chars for the specified minimum lowercase count.'
        }
        if ($AllowedUppercase.Count -lt 1 -and $MinUppercaseChars -gt 0)
        {
            throw 'There are not enough allowed uppercase chars for the specified minimum uppercase count.'
        }
        if ($AllowedDigits.Count -lt 1 -and $MinDigits -gt 0)
        {
            throw 'There are not enough allowed digits for the specified minimum digit count.'
        }
        if ($AllowedSpecial.Count -lt 1 -and $MinSpecialChars -gt 0)
        {
            throw 'There are not enough allowed special chars for the specified minimum special count.'
        }
        # Function to generate random chars for array. Takes the chararray to populate and an amount as input
        function GetRandomChars ([char[]]$CharArray, [int]$Amount)
        {
            # Check if input is valid
            if ($CharArray.Count -gt 0 -and $Amount -gt 0)
            {
                # Fills array with random chars from input array
                for ($i = 0; $i -lt $Amount; $i++)
                {
                    $CharArray[(Get-Random -Maximum $CharArray.Count)]
                }
            }
        }
    }
    #------------------------------------------| BUILD PASSWORD |------------------------------------------#
    process
    {
        try
        {
            if ($AsString)
            {
                # User wants output as plain text string
                $StringBuilder = [System.Text.StringBuilder]::new($Length)
            }
            else
            {
                # User want output as secure string
                $SecureString = [securestring]::new()
            }

            # Get all random chars in fixed position with GetRandomChars function
            # Randomize their order with Get-Random,
            # for each char either append to secure string or plain text string depending on user choice
            @(
                GetRandomChars -CharArray $AllowedLowercase -Amount $MinLowercaseChars
                GetRandomChars -CharArray $AllowedUppercase -Amount $MinUppercaseChars
                GetRandomChars -CharArray $AllowedDigits    -Amount $MinDigits
                GetRandomChars -CharArray $AllowedSpecial   -Amount $MinSpecialChars
                GetRandomChars -CharArray $AllAllowedChars  -Amount $FillerCharCount
            ) |
            Get-Random -Count $Length | ForEach-Object -Process {
                if ($AsString)
                {
                    $null = $StringBuilder.Append($_)
                }
                else
                {
                    $SecureString.AppendChar($_)
                }
            }
            # Entire pass wword has been built
            if ($AsString)
            {
                # Return plaintext string if user asks for that
                $StringBuilder.ToString()
            }
            else
            {
                # return secure string if user did not ask for cleartext
                $SecureString
            }
        }
        catch
        {
            Write-Error $_
        }
    }
}

<#

  __          __   _ _              _    _           _    _____                           _
 \ \        / /  (_) |            | |  | |         | |  / ____|                         | |
  \ \  /\  / / __ _| |_ ___ ______| |__| | ___  ___| |_| (___   ___ _ __   ___ _ __ __ _| |_ ___  _ __
   \ \/  \/ / '__| | __/ _ \______|  __  |/ _ \/ __| __|\___ \ / _ \ '_ \ / _ \ '__/ _` | __/ _ \| '__|
    \  /\  /| |  | | ||  __/      | |  | | (_) \__ \ |_ ____) |  __/ |_) |  __/ | | (_| | || (_) | |
     \/  \/ |_|  |_|\__\___|      |_|  |_|\___/|___/\__|_____/ \___| .__/ \___|_|  \__,_|\__\___/|_|
                                                                   | |
                                                                   |_|

#>
<#
    .SYNOPSIS
        Writes to a console a formatted string, like the one below
        #------------------------------------------| SOME TEST STRING |------------------------------------------#
#>
#---------------------------------------------| PARAMETERS |---------------------------------------------#
Function Write-HostSeperator
{
    Param
    (
        # The string that will be formatted on output
        [Parameter(Mandatory = $true)]
        [String]
        $InputString,

        # Switch to $false, to not capitalize the message
        [Parameter()]
        [boolean]
        $Capitailze = $true,

        # Width of the formatted text
        [Parameter()]
        [int]
        $Width = 100,

        # Change the type of written output
        [Parameter()]
        [ValidateSet("Write-Host", "Write-Error","Write-Debug", "Write-Verbose", IgnoreCase = $false)]
        [string]
        $AsOutputType = "Write-Host"
    )

    # Always trim the input
    $InputString = $InputString.Trim()

    # Defaults to capitalized output
    if ($Capitailze -eq $true)
    {
        $InputString = $InputString.ToUpper()
    }

    # Include extra dash on odd length input strings
    if ($InputString.Length % 2 -eq 0) {$ExtraDash = ""} else {$ExtraDash = "-"}

    # Replace to each side
    $Dashcount = [math]::Floor(($Width - $InputString.Length) / 2)

    # Build formatted string with string multiplication, string addition and string manipulation :)
    $FormattedString = "`n#" + $ExtraDash + $("-" * $Dashcount) + "| " + $InputString + " |" + $("-" * $Dashcount) + "#"

    # Output as desired type
    switch ($AsOutputType)
    {
        "Write-Host"    {Write-Host    "$FormattedString"}
        "Write-Error"   {Write-Error   -Message "$FormattedString"}
        "Write-Debug"   {Write-Debug   -Message "$FormattedString"}
        "Write-Verbose" {Write-Verbose -Message "$FormattedString"}
    }
}
<#
   _____                            _          _   _      _____
  / ____|                          | |        | \ | |    |  __ \
 | |     ___  _ __  _ __   ___  ___| |_ ______|  \| | ___| |__) |_ _ _ __ ___
 | |    / _ \| '_ \| '_ \ / _ \/ __| __|______| . ` |/ __|  ___/ _` | '_ ` _ \
 | |___| (_) | | | | | | |  __/ (__| |_       | |\  | (__| |  | (_| | | | | | |
  \_____\___/|_| |_|_| |_|\___|\___|\__|      |_| \_|\___|_|   \__,_|_| |_| |_|

#>
Function Connect-NcPAM
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $PamCredential
    )

    try
    {
        # Establish Session to PAM.
        $PAMArguments = @{
            "Credential" = $PAMCredential
            "BaseURI"    = "https://pam.nchosting.dk"
            "type"       = "LDAP"
        }

        Write-Host "Connecting to PAM ... " -NoNewline
        [void]::(New-PASSession @PAMArguments)
        Write-host "OK"
    }
    catch
    {
        Write-Host 'FAIL!' -BackgroundColor Red
        Write-Host 'Could not connect to Pam. Exiting'
        throw
    }
}

function Connect-SplunkServer {
    # This function will return a SessionKey for use in future queries

    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $SplunkCredential,

        [Parameter(Mandatory = $true)]
        [String]
        $BaseURI
    )

    Write-Host "Establishing Splunk Session ... " -NoNewline

    $params = @{
        Uri         = $BaseURI + "/services/auth/login"
        Method      = "POST"
        ContentType = "application/json"
        TimeoutSec  = 15
        body        =  @{
            username    = $SplunkCredential.username
            password    = $SplunkCredential.GetNetworkCredential().Password
        }
    }

    # Run query to get key
    try
    {
        $response = Invoke-WebRequest @params
        Write-Host "OK"
        (($response.content -split "<sessionkey>")[1] -split "</sessionKey>")[0]
    }
    catch
    {
        Write-Verbose "FAIL"
        Write-error $error[0]
    }
}


function New-SplunkQuery {
    # This function will create a search query in splunk.
    # It returns the SID needed to retrieve the result
    # Function is used in conjunction with Get-SplunkQuery. In programmatic splunk, creating and retrival of searches are decoupled

    param
    (
        [Parameter(Mandatory)]
        [String]
        $BaseURI,

        [Parameter(Mandatory)]
        [String]
        $SessionKey,

        [Parameter(Mandatory)]
        [String]
        $SearchQuery
    )

    # Prepare parameters for query generation
    $params = @{
        Uri        = $BaseURI + "/services/search/jobs"
        Method     = "POST"
        TimeoutSec = "15"
        body       = @{
            "search" = $SearchQuery
        }
        Headers    = @{
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "authorization" = "Splunk $SessionKey"
        }
    }

    # Create the search
    try
    {
        Write-Host "Creating Splunk query ... " -NoNewline
        $response = Invoke-RestMethod @params
        return ($response.response.sid)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }
}

function Get-SplunkQuery {
    # This function returns a result in splunk based on input SID
    # Function is used in conjunction with Create-SplunkQuery. In programmatic splunk, creating and retrival of searches are decoupled

    param
    (
        [Parameter(Mandatory)]
        [String]
        $BaseURI,

        [Parameter(Mandatory)]
        [String]
        $SessionKey,

        [Parameter(Mandatory)]
        [String]
        $ID,

        [Parameter()]
        [System.Boolean]
        $Wait = $true,

        [parameter()]
        [System.Int32]
        $WaitSeconds = 2
    )

    # Prepare object for parameters
    $params = @{
        Uri        = $BaseURI + "/services/search/v2/jobs/$ID/results"
        Method     = "GET"
        TimeoutSec = "15"
        Headers    = @{
            "Content-Type"  = "application/json"
            "Accept"        = "application/json"
            "authorization" = "Splunk $SessionKey"
        }
        body       = @{
            "output_mode" = "json"
            "count"       = "0"
        }
    }

    # Invoke the search function
    if ($wait)
    {
        $SearchTime = Measure-Command {

            Write-Host "Retrieving results ..." -NoNewline
            # Query Splunk for results each two seconds
            do
            {
                Start-Sleep -Seconds $WaitSeconds
                $Response = Invoke-RestMethod @params
                Write-Host "." -NoNewline
            }
            while ($null -eq $Response)
        }

        Write-Host " OK"
        Write-Verbose "Found Splunk results in $($SearchTime.Totalseconds) seconds"
    }
    else
    {
        $Response = Invoke-RestMethod @params
        return($Response.results)
    }
}


function get-ToolkitList {
    param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $ToolkitCredential,

        [Parameter()]
        [String]
        $ToolkitListID,

        [Parameter()]
        [String]
        $ToolkitListName

    )

    TODO: Implement this
    Write-Host "Function not Get-ToolkitList not yet implemented"
}
<#
                                                                                                             
#>
# TODO: Await PS7 Implementation of this script
#Function Mount-ToolkitAsFolder
#{
#    Param
#    (
#        [Parameter(Mandatory)]
#        [pscredential]
#        $Toolkitcredential,
#
#        [Parameter()]
#        [String]
#        $TargetFolder = "VMware\10 - Reporting",
#
#        [Parameter()]
#        [String]
#        $Base = "\\int-goto.netcompany.com@SSL\DavWWWRoot\cases\GTO27\NCINFRA\DocumentLibrary\",
#
#        [Parameter()]
#        [String]
#        $MountName = "Toolkit"
#    )
#
#    try
#    {
#        #Write-Host 'Connecting to Toolkit ... ' -NoNewline
#        
#        $Root = $Base + $TargetFolder
#
#        Write-Host 'Connecting to Toolkit on $Root as name "Toolkit"'
#
#        $Parameters = @{
#            Name       = "Toolkit"
#            PSProvider = "FileSystem"
#            Root       = "\\goto.netcompany.com@SSL\DavWWWRoot\cases\GTO27\NCINFRA\DocumentLibrary\SAN-Storage\"
#            Credential = $ToolkitCredential
#        }
#        
#        New-PSDrive @Parameters
#        
#    }
#    catch
#    {
#        Write-Host 'FAIL!' -BackgroundColor Red
#        #Exit 1
#    }
#}

<#
   _____ _                         _____           _             _ ____  _           _      _           _          _
  / ____| |                       / ____|         | |           | |  _ \| |         | |    | |         | |        | |
 | |    | | ___  __ _ _ __ ______| |     ___ _ __ | |_ _ __ __ _| | |_) | | __ _  __| | ___| |     __ _| |__   ___| |
 | |    | |/ _ \/ _` | '__|______| |    / _ \ '_ \| __| '__/ _` | |  _ <| |/ _` |/ _` |/ _ \ |    / _` | '_ \ / _ \ |
 | |____| |  __/ (_| | |         | |___|  __/ | | | |_| | | (_| | | |_) | | (_| | (_| |  __/ |___| (_| | |_) |  __/ |
  \_____|_|\___|\__,_|_|          \_____\___|_| |_|\__|_|  \__,_|_|____/|_|\__,_|\__,_|\___|______\__,_|_.__/ \___|_|

#>
Function Clear-CentralBladeLabel
{
    Param
    (
        # Normal expected input parameter
        [Parameter()]
        [String]
        $FQDN,

        # Manual behavior; get blade based on HostName instead of FQDN
        [Parameter()]
        [String]
        $ShortName
    )

    # Convert FQDN to service profile value
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    # Get serviceprofile, depending on what was supplied
    if     ($FQDN)      {$ServiceProfile = Get-UCSCentralServiceProfile -Descr $FQDN}
    elseif ($ShortName) {$ServiceProfile = Get-UCSCentralServiceProfile -Name  $ShortName}

    # Quit if none found
    if ($null -eq $ServiceProfile) {throw}

    Write-Host "Clearing Usrlbl on $FQDN " -NoNewLine
    $Params = @{
        UsrLbl  = ""
        Confirm = $False
        Force   = $true
    }

    # Execute function in void brackets to suppress default UCS return values
    [void]::($ServiceProfile | Set-UCSCentralServiceProfile @Params)

    # Continue directly, if there is a pending reboot on the blade. If this is the case, we don't care if the blade reboots - it's supposed to do that
    if ($ServiceProfile.OperState -eq "pending-reboot")
    {
        Write-Host " OK"
        Write-Host "Blade has a pending reboot, not waiting for `"serviceprofile.ConfigState`" = `"applied`""
    }
    else
    {
        # Sleep until it's finished applying the new label
        $ExecutionTime = Measure-Command {
            do
            {
                Write-Host "." -NoNewline
                Start-Sleep -Seconds 3

                # Update ServiceProfile variable
                $ServiceProfileUpdate = Get-UCSCentralServiceProfile -Descr $ServiceProfile.Descr
            }
            until (($ServiceProfileUpdate.UsrLbl -eq $UsrLbl) -and ($ServiceProfileUpdate.ConfigState -ne "applying"))
        }
        # Then sleep additional three seconds
        Start-Sleep -Seconds 3
    }

    # Yes, the extra space is supposed to be there
    Write-Host " OK"
    Write-Host "Cleared BladeLabel in $([Math]::Round($ExecutionTime.TotalSeconds,2)) seconds"
}
<#
   _____                            _          _   _      _    _  _____  _____  _____           _             _
  / ____|                          | |        | \ | |    | |  | |/ ____|/ ____|/ ____|         | |           | |
 | |     ___  _ __  _ __   ___  ___| |_ ______|  \| | ___| |  | | |    | (___ | |     ___ _ __ | |_ _ __ __ _| |
 | |    / _ \| '_ \| '_ \ / _ \/ __| __|______| . ` |/ __| |  | | |     \___ \| |    / _ \ '_ \| __| '__/ _` | |
 | |___| (_) | | | | | | |  __/ (__| |_       | |\  | (__| |__| | |____ ____) | |___|  __/ | | | |_| | | (_| | |
  \_____\___/|_| |_|_| |_|\___|\___|\__|      |_| \_|\___|\____/ \_____|_____/ \_____\___|_| |_|\__|_|  \__,_|_|


#>
Function Connect-NcUCSCentral
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [pscredential]
        $UCSCredential,

        [Parameter()]
        [boolean]
        $Force = $false
    )

    # Don't connect if already connected
    if ($global:defaultUcsCentral -and ($Force -ne $true))
    {
       Write-Host "You're aleady connected to $(($global:defaultUcsCentral).Uri)"
    }
    else
    {
        if ($Force -eq $true)
        {
            Disconnect-UcsCentral
        }

        try
        {
            Write-Host 'Connecting to UCSCentral ... ' -NoNewline

            $Params = @{
                Name       = "ucscentral.nchosting.dk"
                Credential = $UCSCredential
            }
            [void]::(Connect-UCSCentral @Params)


            Start-Sleep -seconds 2

            # Put the connectionvariable from the function scope, and populate the global scope with it
            $ConnectionVariable       = (get-variable defaultUcsCentral).value
            $global:defaultUcsCentral = $ConnectionVariable
            Write-Host 'OK'

            Write-Host "Conneted to: $($ConnectionVariable.Uri)"

        }
        catch
        {
            Write-Host 'FAIL!' -BackgroundColor Red
            Write-Host 'Could not connect to UCS-Central. Exiting'
            throw
        }
    }
}

Function Dismount-UCSvMedia
{
    param
    (

        [Parameter()]
        [String]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'FQDN'")]
        $FQDN,

        [Parameter()]
        [Cisco.UcsCentral.ComputeBlade]
        $ServiceProfile
    )

    # Get the ServiceProfile for the requested FQDN
    if ($FQDN)
    {
        $ServiceProfile = Get-UcsCentralServiceProfile -Descr $FQDN
    }

    # Clear the vMediaPolicy
    try
    {
        Write-Host "Removing vMedia ... " -NoNewline
        $ServiceProfile | Set-UcsCentralServiceProfile -VmediaPolicyName '' -Confirm:$false -Force
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL"
        $Error
    }
}

<#
  __  __                   _          _    _  _____  _____      __  __          _ _
 |  \/  |                 | |        | |  | |/ ____|/ ____|    |  \/  |        | (_)
 | \  / | ___  _   _ _ __ | |_ ______| |  | | |    | (_____   _| \  / | ___  __| |_  __ _
 | |\/| |/ _ \| | | | '_ \| __|______| |  | | |     \___ \ \ / / |\/| |/ _ \/ _` | |/ _` |
 | |  | | (_) | |_| | | | | |_       | |__| | |____ ____) \ V /| |  | |  __/ (_| | | (_| |
 |_|  |_|\___/ \__,_|_| |_|\__|       \____/ \_____|_____/ \_/ |_|  |_|\___|\__,_|_|\__,_|


#>
Function Mount-UCSvMedia
{
    param
    (
        # FQDN to work on
        [Parameter(Mandatory)]
        [String]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'FQDN'")]
        $FQDN,

        # OS to install
        [Parameter(Mandatory)]
        [String]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'OS'")]
        $OS,

        # Path of iso to mount
        [Parameter(Mandatory)]
        [String]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'MountPathPrefix'")]
        $InputPath,

        # Name of the .iso file to mount
        [Parameter(Mandatory)]
        [String]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'ISOName'")]
        $ISOName
    )

    # Prepare vMedia
    try
    {
        Write-Host "Preparing vMedia ... " -NoNewline

        # ---- File.nchosting ----
        # Change MS notation to Unix
        $InputPath = ($InputPath.Replace("\","/")).Trim()

        # If user got path from a windows explorer window, change it to match expected notation. Example:
        # //file.nchosting.dk/Software/Microsoft/Windows/10
        # Becomes
        # /Microsoft/Windows/10
        if ($InputPath -match "//file.nchosting.dk/Software/")
        {
            $InputPath = ($inputpath -split "//file.nchosting.dk/Software/")[1]
        }

        # Correct users who does not use file endings
        if ($IsoName -notmatch ".iso")
        {
            $IsoName = $IsoName + ".iso"
        }


        # ---- UCS Central ----

        # Get the ServiceProfile for the requested FQDN
        $ServiceProfile = Get-UcsCentralServiceProfile -Descr $FQDN

        # Selecting the correct vMedia policy from OS type
        $SelectedvMediaPolicy = Get-UcsCentralCimcvmediaMountConfigPolicy  | Where-Object {$_.Name -like "*$OS*"}

        # Prepare Parameters for vmedia remediation
        $Params = @{
            "ISOName"       = $ISOName
            "ImagePath"     = $InputPath
            "Confirm"       = $false
            "Force"         = $true
        }

        # Configuring vMedia policy to the image file
        $SelectedvMediaPolicy | Get-UcsCentralCimcvmediaConfigMountEntry  | Set-UcsCentralCimcvmediaConfigMountEntry @Params

    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not prepare vMedia"
        $Error
    }


    # Mounting vMedia Policy
    try
    {
        Write-Host "Mounting vMedia to $FQDN ... " -NoNewline

        $MountParams = @{
            "VmediaPolicyName" = $SelectedvMediaPolicy.Name
            "Confirm"          = $false
            "Force"            = $true
        }
        $ServiceProfile | Set-UcsCentralServiceProfile @MountParams

        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not mount vMedia"
        $error
    }

    #-----------------------------------------------| Cleanup |----------------------------------------------#
    if ($error)
    {
        Write-Host "Script finished with errors:" -BackgroundColor Red
        $Error
        Exit 1
    }
}
<#
   _____      _           _____           _             _ ____  _           _      _           _          _
  / ____|    | |         / ____|         | |           | |  _ \| |         | |    | |         | |        | |
 | (___   ___| |_ ______| |     ___ _ __ | |_ _ __ __ _| | |_) | | __ _  __| | ___| |     __ _| |__   ___| |
  \___ \ / _ \ __|______| |    / _ \ '_ \| __| '__/ _` | |  _ <| |/ _` |/ _` |/ _ \ |    / _` | '_ \ / _ \ |
  ____) |  __/ |_       | |___|  __/ | | | |_| | | (_| | | |_) | | (_| | (_| |  __/ |___| (_| | |_) |  __/ |
 |_____/ \___|\__|       \_____\___|_| |_|\__|_|  \__,_|_|____/|_|\__,_|\__,_|\___|______\__,_|_.__/ \___|_|


#>
Function Set-CentralBladeLabel
{
    Param
    (
        # Normal expected input parameter
        [Parameter()]
        [String]
        $FQDN,

        # Manual behavior; get blade based on HostName instead of FQDN
        [Parameter()]
        [String]
        $ShortName,

        # Use "Clear-CentralBladeLabel" to remove label
        [Parameter(Mandatory)]
        [String]
        $UsrLbl
    )

    # Convert FQDN to service profile value
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    # Get serviceprofile, depending on what was supplied
    if ($FQDN)          {$ServiceProfile = Get-UCSCentralServiceProfile -Descr $FQDN}
    elseif ($ShortName) {$ServiceProfile = Get-UCSCentralServiceProfile -Name  $ShortName}

    # Quit if none found
    if ($null -eq $ServiceProfile) {throw}

    Write-Host "Setting Usrlbl to `"$UsrLbl`" on $FQDN ..." -NoNewLine
    $Params = @{
        UsrLbl  = $UsrLbl
        Confirm = $False
        Force   = $true
    }
    # Execute function in void brackets to suppress default UCS return values
    [void]::($ServiceProfile | Set-UCSCentralServiceProfile @Params)

    # Continue directly, if there is a pending reboot on the blade. If this is the case, we don't care if the blade reboots - it's supposed to do that
    if ($ServiceProfile.OperState -eq "pending-reboot")
    {
        Write-Host " OK"
        Write-Host "Blade has a pending reboot: Not waiting for `"serviceprofile.ConfigState`" = `"applied`""
    }
    else
    {
        # Sleep until it's finished applying the new label
        $ExecutionTime = Measure-Command {
            do
            {
                Write-Host "." -NoNewline
                Start-Sleep -Seconds 3

                # Update ServiceProfile variable
                $ServiceProfileUpdate = Get-UCSCentralServiceProfile -Descr $ServiceProfile.Descr
            }
            until (($ServiceProfileUpdate.UsrLbl -eq $UsrLbl) -and ($ServiceProfileUpdate.ConfigState -ne "applying"))
        }
        # Then sleep additional three seconds
        Start-Sleep -Seconds 3

        # Yes, the extra space is supposed to be there
        Write-Host " OK"
        Write-Host "Applied BladeLabel in $([Math]::Round($ExecutionTime.TotalSeconds,2)) seconds"
    }
}
<#
   _____ _             _           _____           _             ___      ____  __ _               _
  / ____| |           | |         / ____|         | |           | \ \    / /  \/  | |             | |
 | (___ | |_ __ _ _ __| |_ ______| |     ___ _ __ | |_ _ __ __ _| |\ \  / /| \  / | |__   ___  ___| |_
  \___ \| __/ _` | '__| __|______| |    / _ \ '_ \| __| '__/ _` | | \ \/ / | |\/| | '_ \ / _ \/ __| __|
  ____) | || (_| | |  | |_       | |___|  __/ | | | |_| | | (_| | |  \  /  | |  | | | | | (_) \__ \ |_
 |_____/ \__\__,_|_|   \__|       \_____\___|_| |_|\__|_|  \__,_|_|   \/   |_|  |_|_| |_|\___/|___/\__|


#>
Function Start-CentralVMhost
{
    param
    (
        [Parameter(Mandatory)]
        [Cisco.UcsCentral.LsServer]
        $ServiceProfile,

        [Parameter(Mandatory)]
        [String]
        $PowerOnLabel
    )

    # Get the Blade
    $Blade = Get-UcsCentralBlade | Where-object {$_.AssignedToDN -match $ServiceProfile.Dn}

    # Continue only if you found a single blade
    if ($Blade.Count -eq 1)
    {
        # State which blade is being worked on
        Write-HostSeperator "$($Blade.Serial), $($ServiceProfile.Rn)" -Width 45

        # only continue if you find a single blade that is powered off
        if ($Blade.OperPower -eq "off")
        {
            try
            {
                # Sets the usrlabel of the blade to match poweron state
                Set-CentralBladeLabel -FQDN ($ServiceProfile.Descr) -UsrLbl $PowerOnLabel

                # Power On the blade
                Write-Host "Sending PowerON signal ... " -NoNewline
                do
                {
                    [void]($ServiceProfile | Get-UcsCentralLsServerOperation | Set-UcsCentralLsServerOperation -State 'admin-up'        -Confirm:$false -Force -ErrorAction SilentlyContinue)
                    [void]($ServiceProfile | Get-UcsCentralLsServerOperation | Set-UcsCentralLsServerOperation -State 'cycle-immediate' -Confirm:$false -Force -ErrorAction SilentlyContinue)

                    Start-Sleep -Seconds 2
                    $PowerOnTest = Get-UcsCentralServiceProfile -Dn $ServiceProfile.Dn
                }
                Until($PowerOnTest.OperState -ne "power-off")
                Write-Host "OK"

                # Sets state in vCenter
                Write-Host "Setting Attributes in vCenter ... " -NoNewline
                $VMhost = Get-VMhost $ServiceProfile.Descr
                $VMhost.ExtensionData.setCustomValue('State', $PowerOnLabel)
                $VMhost.ExtensionData.setCustomValue('StateDate', $(Get-FormattedTime))
                Write-Host "OK"
            }
            catch
            {
                Write-Host "FAIL!" -BackgroundColor Red
                throw
            }
        }
        else
        {
            Write-Host "$($ServiceProfile.Descr) is not powered off!" -BackgroundColor Red

            $VMhost = Get-vmhost -name $ServiceProfile.Descr

            if ($VMhost.ConnectionState -in "Maintenance")
            {
                Write-Host "Blade state is `"$PowerOfflabel`", but is actually VMhost is in Maintenance. Resetting label"
                Set-CentralBladeLabel -FQDN ($ServiceProfile.Descr) -UsrLbl $PowerOnLabel
            }
            elseif ($VMhost.ConnectionState -in "Connected")
            {
                Write-Host "Blade state is `"$PowerOfflabel`", but is acctually VMhost is connected. Resetting label"
                Set-CentralBladeLabel -FQDN ($ServiceProfile.Descr) -UsrLbl "live"
            }
            else
            {
                Write-Host "Blade state is `"$PowerOfflabel`", Blade is powered ON, but VMhost is disconnected"
                Resolve-UCSPowerOnProgress -VMhost $VMhost
            }
        }
    }
    else
    {
        Write-Host "$($ServiceProfile.Dn) is not single!" -BackgroundColor Red
        throw
    }
}
function Add-IntersightvLANs
{
    param
    (
        [Parameter(Mandatory)]
        [string[]]
        $vlanObject
    )

    #Should be changed into Toolkit custom integration
    #$vlanObject = @("sdfi-daf-test-app,2193","sdfi-daf-test-dmz,2194","sdfi-daf-test2-db,2195","sdfi-daf-test2-app,2196","sdfi-daf-test2-dmz,2197","sdfi-daf-dev-db,2198","sdfi-daf-dev-app,2199","sdfi-daf-dev-dmz,2200","sdfi-daf-lbap,2201","sdfi-daf-lbanp,2202","sdfi-daf-dnoc,2203","sdfi-daf-transit01,286","sdfi-daf-conv,2391")

    #Loops through all VLANs in the VLANObject array
    Foreach ($vlan in $vlanObject)
    {

        # Splits the VLAN name from the VLAN ID on comma ","
        $vlanname = $vlan.Split(',')[0]

        # Splits VLAN ID from VLAN Name on ","
        $vlanid = $vlan.Split(',')[1]

        # Validating if VLAN ID Exists as a Fabric VLAN
        $vlanIdCheck = Get-IntersightFabricVlan -VlanId $vlanID

        Write-Host "Validating if a VLAN with ID: $vlanID exists"
        # If VLAN doesn't exist, keep moving through loop
        If ($null -eq $vlanIdCheck)
        {
            Write-Host "No VLAN detected with ID: $vlanID"
        }
        elseif ($null -ne $vlanIdCheck)
        {
            #If VLAN exists, skip VLAN (Continue)
            Write-Host "Vlan with ID $vlanID located on VLAN $($VlanIdCheck.Name) skipping creation"
            Continue
        }
        # Validating if VLAN Name exists as a Fabric VLAN
        $vlanNameCheck = Get-IntersightFabricVlan -Name $vlanName

        # If VLAN Name is Free keep going through loop
        if ($null -eq $vlanNameCheck)
        {
            Write-Host "No VLAN detected with Name: $vlanName"
        }
        elseif ($null -ne $vlanNameCheck)
        {
            # If VLAN Name is taken, skip VLAN creation (continue)
            Write-Host "Vlan with Name $vlanName located on $VLAN $($VlanNameCheck.Name) skipping creation"
            Continue
        }

        # Selecting UCS Multicast policy (Default policy) for vlan creation
        $UCSMulticastPolicy = Get-IntersightFabricMulticastPolicy -Name UCS-Multicast-Policy
        Write-Host "Multicast Policy $($UCSMulticastPolicy.Name) Selected"

        # Selecting UCS VLAN Configuration Policy (Stretched UCS VLAN Configuration for all UCS Domains)
        $FabricVlanConfiguration = Get-IntersightFabricEthNetworkPolicy -Name UCS-VLAN-Configurations-Stretched
        Write-Host "Fabric VLAN Allow configuration selected $($FabricVlanConfiguration.Name)"

        # Creates a new fabric VLAN with the VLAN ID and VLAN Name from previous split, adds it to the stretched UCS VLAN Configuration
        [void]($NewVlan = New-IntersightFabricVlan -Name $vlanName -VlanId $vlanId -EthNetworkPolicy $FabricVlanConfiguration -MulticastPolicy $UCSMulticastPolicy)
        Write-Host "VLAN $($NewVlan.Name) with ID: $($NewVlan.vlanid) has been created"

        # Gets the ESXi NetworkGroupPolicy - this is used for all ESXi hosts
        $ESXiHostDataVLAN = Get-IntersightFabricEthNetworkGroupPolicy -Name UCS-Ethernet-VMW-DATA-Network-Global-Group-VLAN

        # Gets the entire string of VLAN IDs, splits the entire string on Comma "," and validates if the policy contains the VLAN ID for addition
        if ($NewVlan.Id -in $ESXiHostDataVLAN.VlanSettings.AllowedVlans.Split(','))
        {
            Write-Host 'Vlan already exists in UCS-Ethernet-VMW-DATA-Network-Global-Group-VLAN skipping addition'
            Continue
        }
        else
        {
            # Creates a new string with the currently allowed VLANs and adds the newly allowed VLAN
            $AllowedVlans = $ESXiHostDataVLAN.VlanSettings.AllowedVlans + ',' + $NewVlan.vlanId

            # Writes a simple count of VLANs before and after addition of the VLAN
            Write-Host "Old VLAN Count $($ESXiHostDataVLAN.VlanSettings.AllowedVlans.Split(',').Count)"
            Write-Host "New VLAN Count $($AllowedVlans.Split(',').Count)"

            # Sets the new string of VLANs as allowed VLANs
            [void](Set-IntersightFabricEthNetworkGroupPolicy -Moid $ESXiHostDataVlan.Moid -VlanSettings (Initialize-IntersightFabricVlanSettings -AllowedVlans $AllowedVlans))
        }

        #Starts fetching fabric interconnects with pending deployments, does this untll amount of Fabrics with pending activites are 4 Fabrics or more - should be changed to fetch number of UCS Domains / Number of fabrics prior to entering the do until
        Write-Host 'Fetching Fabric Interconnects with pending deployments'
        do
        {
            #Remove this start-sleep? Adjust do until instead of $FabricsWithChanges.Count -ge 4, do -gt count of pod+locaiton keys
            #Silly sleep for 5 seconds...
            Start-Sleep -Seconds 5

            # Gets all pending changes on fabric interconnects
            $AllPendingChanges = Get-IntersightFabricConfigChangeDetail

            # Gets the Moid for all the fabrics with pending tasks
            $FabricsWithChanges = $AllPendingChanges.Parent.ActualInstance.Moid | Select-Object -Unique

            # This Until should be changed from hardcode to a dynamic resolve of number of Fabrics in Intersight
        }until($FabricsWithChanges.Count -ge 4)

        # Walks through each single fabric and deploys changes
        Foreach ($Fabric in $FabricsWithChanges)
        {
            $FabricSwitchProfile = Get-IntersightFabricSwitchProfile -Moid $Fabric
            Write-Host "Deploying vlan on $($FabricSwitchProfile.Name)"

            # Sets fabric state as deploy (Deploys changes)
            [void](Set-IntersightFabricSwitchProfile -Moid $FabricSwitchProfile.Moid -Action 'Deploy')
        }
        #Replace this start-Sleep with a get pending activities on Fabric Interconnects
        Start-Sleep -Seconds 120
    }
    Write-Host 'Deploying vlan on all ESXi hosts'
    Start-Sleep -Seconds 120
    #Hardcore move - gets all intersight server profiles that contains "esx" in the name and deploys changes.
    #Important! Change this to fetch all ESXi hosts containing the streched policy that has been modified, and then deploy on those only.
    #Deploying on all hosts with esx in the name can lead to deploying on wrong hosts / servers
    Get-IntersightServerProfile | Where-Object {$_.Name -match 'esx'} | Set-IntersightServerProfile -Action 'Deploy'

    Add-Compute-Vlan -vlanObject $VLANObject
}
<#
    _____             __ _                    _____       _                _       _     _    _____                            _   _             
  / ____|           / _(_)                  |_   _|     | |              (_)     | |   | |  / ____|                          | | (_)            
 | |     ___  _ __ | |_ _ _ __ _ __ ___ ______| |  _ __ | |_ ___ _ __ ___ _  __ _| |__ | |_| |     ___  _ __  _ __   ___  ___| |_ _  ___  _ __  
 | |    / _ \| '_ \|  _| | '__| '_ ` _ \______| | | '_ \| __/ _ \ '__/ __| |/ _` | '_ \| __| |    / _ \| '_ \| '_ \ / _ \/ __| __| |/ _ \| '_ \ 
 | |___| (_) | | | | | | | |  | | | | | |    _| |_| | | | ||  __/ |  \__ \ | (_| | | | | |_| |___| (_) | | | | | | |  __/ (__| |_| | (_) | | | |
  \_____\___/|_| |_|_| |_|_|  |_| |_| |_|   |_____|_| |_|\__\___|_|  |___/_|\__, |_| |_|\__|\_____\___/|_| |_|_| |_|\___|\___|\__|_|\___/|_| |_|
                                                                             __/ |                                                              
                                                                            |___/                                                                                                                                                                                                   
  #>
Function Confirm-IntersightConnection
{
    # enable -verbose flag, but don't take any inputs
    [cmdletbinding()]
    param()

    # Check if current connection
    if($null -eq (get-intersightconfiguration))
    {
        Write-Host "Not connected to Intersight. Ending"
        Exit 1
    }
    else 
    {
        Write-Verbose -Message "You are connected to Intersight!"
    }
}

Function Connect-Intersight
{
    param
    (
        # Expect string, containing ID
        [Parameter(Mandatory)]
        [string]
        $ApiKeyID,

        # Expect string, containing Key file Path
        [Parameter()]
        [string]
        $ApiKeyFilePath,

        # Expect Key File contents
        [Parameter()]
        [String]
        $ApiKeyString,

        # Parameter help description
        [Parameter()]
        [String]
        $Intersight = "ncop-ucsp-int01.prod.ucs.ncop.nchosting.dk"
    )

    Write-Verbose "Validating keys from input ... "
    # Check validit of filelength
    if ($ApiKeyID.Length -eq 74)
    {
        Write-Verbose "OK"
    }
    else
    {
        Write-Error "FAIL - No keys input"
        throw
    }

    Write-Verbose "Validating connection to Intersight ... "
    [void]::($ConnectionCheck = test-netconnection $Intersight -port 443)

    if($ConnectionCheck.TcpTestSucceeded -eq "true")
    {
        Write-Verbose "OK"
        Write-Verbose "ComputerName:     $($ConnectionCheck.ComputerName)"
        Write-Verbose "RemotePort:       $($ConnectionCheck.RemotePort)"
        Write-Verbose "TcpTestSucceeded: $($ConnectionCheck.TcpTestSucceeded)"
    }
    else
    {
        Write-Error "FAIL!" -BackgroundColor Red
        throw "No valid network connection to Intersight, check DNS and Firewall"
    }

    # Build connection paramaters from documentation

    $onprem = @{
        BasePath             = "https://$Intersight"     # $Intersight contains FQDN of On-prem VM
        ApiKeyId             = $ApiKeyID                                                # String containing ID
        HttpSigningHeader    = @('(request-target)', 'Host', 'Date', 'Digest')          # Required header from Documentation
        SkipCertificateCheck = $true                                                    # Get around HTTP error
        debug                = $true
    }

    if ($ApiKeyFilePath)
    {
        $onprem["ApiKeyFilePath"] = $ApiKeyFilePath
        Write-Verbose "Logging in using Keyfile Path"
    }
    elseif ($ApiKeyString)
    {
        $onprem["ApiKeyString"] = $ApiKeyString
        Write-Verbose "Logging in using KeyString"
    }

    # Don't continue on error
    if ($onprem)
    {
        # Finally finish the connection
        try
        {
            Write-Host 'Connecting to Intersight ... ' -NoNewline
            Set-IntersightConfiguration @onprem
            Start-sleep -seconds 2
            Write-Host "OK"
        }
        catch
        {
            # Exit on Failure
            Write-Host "FAIL"
            throw
        }
    }
    else
    {
        Write-Host "No `$Onprem variable"
    }
}

<#
   _____  _                                 _        _____       _                _       _     _         __  __          _ _
 |  __ \(_)                               | |      |_   _|     | |              (_)     | |   | |       |  \/  |        | (_)
 | |  | |_ ___ _ __ ___   ___  _   _ _ __ | |_ ______| |  _ __ | |_ ___ _ __ ___ _  __ _| |__ | |___   _| \  / | ___  __| |_  __ _
 | |  | | / __| '_ ` _ \ / _ \| | | | '_ \| __|______| | | '_ \| __/ _ \ '__/ __| |/ _` | '_ \| __\ \ / / |\/| |/ _ \/ _` | |/ _` |
 | |__| | \__ \ | | | | | (_) | |_| | | | | |_      _| |_| | | | ||  __/ |  \__ \ | (_| | | | | |_ \ V /| |  | |  __/ (_| | | (_| |
 |_____/|_|___/_| |_| |_|\___/ \__,_|_| |_|\__|    |_____|_| |_|\__\___|_|  |___/_|\__, |_| |_|\__| \_/ |_|  |_|\___|\__,_|_|\__,_|
                                                                                    __/ |
                                                                                   |___/
#>

#region---------------------------------------| PARAMETERS |---------------------------------------------#
# Set parameters for the script here
Function Dismount-IntersightvMedia
{
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'FQDN'")]
        [string]
        $FQDN,

        [Parameter()]
        [Boolean]
        $Sleep = $false
    )
    #endregion

    # Don't continue without a connection to Intersight
    Confirm-IntersightConnection

    # Preperation
    try
    {
        Write-Host "Preparing to remove vMedia ... " -NoNewline

        # Gets the server profile, and quit if it does not exist
        if ($null -eq ($SelectedServerProfile = Get-IntersightServerProfile -Name $FQDN))
        {
            Write-Host "FAIL!" -BackgroundColor Red
            Write-Host "Could Not find Serviceprofile for $FQDN"
            $error
            Exit 1
        }

        # EstablishType
        $Type = $serverprofile.assignedserver.ActualInstance.ObjectType

        # Get current mapped vMedia, and quit if none
        $VmediaPolicyMoRef = $SelectedServerProfile.PolicyBucket.ActualInstance | Where-Object {$_.ObjectType -eq "VmediaPolicy"}
        if($null -eq $vmediaPolicyMoRef)
        {
            Write-Host "FAIL!" -BackgroundColor Red
            Write-Host "No vmedia policy mounted, ending script"
            $Error
            Exit 1
        }

        # Prepare input shared for Racks and Blades
        $ActualInstance = $SelectedServerProfile.PolicyBucket.ActualInstance
        $KvmPolicy      = $ActualInstance | Where-Object {$_.ObjectType -eq "KvmPolicy"}

        # Racks needs to have only their KVM policy and no vmedia attached
        if ($Type -eq "ComputeRackUnit")
        {

            $PolicyBucket = @(
                $KvmPolicy
                )

        }
        # Blades needs their entire config minus the vmedia
        elseif ($Type -eq "ComputeBlade")
        {
            $BootPolicy    = $ActualInstance | Where-Object {$_.ObjectType -eq "BootPrecisionPolicy"}
            $vNicPolicy    = $ActualInstance | Where-Object {$_.ObjectType -eq "VnicLanConnectivityPolicy"}
            $vNicSanPolicy = $ActualInstance | Where-Object {$_.ObjectType -eq "VnicSanConnectivityPolicy"}
            $AccessPolicy  = $ActualInstance | Where-Object {$_.ObjectType -eq "AccessPolicy"}
            $BiosPolicy    = $ActualInstance | Where-Object {$_.ObjectType -eq "BiosPolicy"}


            $PolicyBucket = @(
                $KvmPolicy,
                $BootPolicy,
                $vNicPolicy,
                $VnicSanPolicy,
                $AccessPolicy,
                $BiosPolicy )

        }
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could Not prepare vMedia"

        Write-Host "Manual error"
        $error
        Exit 1
    }

    # Unmount vMedia
    try
    {
        Write-Host "Unmounting vMediaPolicy ... " -NoNewline

        $DismountParams = @{
            Moid                 = ($SelectedServerProfile.Moid)
            Action               = "Deploy"
            PolicyBucket         = $PolicyBucket
            ServerAssignmentMode = "None"
        }

        [void]::(Set-IntersightServerProfile @DismountParams)
        Write-Host "OK"

        if ($Sleep)
        {
            # Wait for the to remove
            Write-Host "Sleeping for 15 seconds, to make sure policy is unmounted ... " -NoNewline
            Start-Sleep -Seconds 15
            Write-Host "OK"
        }

        # Remove Policy - only if unmount is successful
        try
        {
            Write-Host "Removing vMediaPolicy ... " -NoNewline
            [void]::(Remove-IntersightVmediaPolicy -moid $VmediaPolicyMoRef.Moid)
            Write-host "OK"

        }
        catch
        {
            Write-Host "FAIL!"
            Write-Host "Could not remove vMediaPolicy"
            $Error
            Exit 1
        }
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not remove vMedia Policy"
        $Error
    }
}
<#
   __  __                   _        _____       _                _       _     _         __  __          _ _
 |  \/  |                 | |      |_   _|     | |              (_)     | |   | |       |  \/  |        | (_)
 | \  / | ___  _   _ _ __ | |_ ______| |  _ __ | |_ ___ _ __ ___ _  __ _| |__ | |___   _| \  / | ___  __| |_  __ _
 | |\/| |/ _ \| | | | '_ \| __|______| | | '_ \| __/ _ \ '__/ __| |/ _` | '_ \| __\ \ / / |\/| |/ _ \/ _` | |/ _` |
 | |  | | (_) | |_| | | | | |_      _| |_| | | | ||  __/ |  \__ \ | (_| | | | | |_ \ V /| |  | |  __/ (_| | | (_| |
 |_|  |_|\___/ \__,_|_| |_|\__|    |_____|_| |_|\__\___|_|  |___/_|\__, |_| |_|\__| \_/ |_|  |_|\___|\__,_|_|\__,_|
                                                                    __/ |
                                                                   |___/
#>

#region---------------------------------------| PARAMETERS |---------------------------------------------#
# Set parameters for the script here
Function Mount-IntersightvMedia
{
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'FQDN'")]
        [string]
        $FQDN,

        [Parameter(Mandatory = $true)]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'InputPath'")]
        [string]
        $InputPath,

        [Parameter(Mandatory = $true)]
        [ValidateScript({![System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_)}, ErrorMessage="WILDCARD in variable 'IsoName'")]
        [string]
        $ISOName,

        [Parameter()]
        [Boolean]
        $Sleep = $false
    )
    #endregion

    #region------------------------------------------| SETUP |-----------------------------------------------#

    # Don't continue without a connection to Intersight
    Confirm-IntersightConnection

    try
    {
        Write-Host "Preparing vMedia ... " -NoNewline

        # Gets the server profile where a mapping is required
        if ($null -eq ($SelectedServerProfile = Get-IntersightServerProfile -Name $FQDN))
        {
            Write-Host "Could Not find Serviceprofile for $FQDN"
        }

        # EstablishType
        $Type = $serverprofile.assignedserver.ActualInstance.ObjectType

        # Generate random number for profile name
        $RandomNumber = Get-Random -Minimum 1 -Maximum 10

        # Generates a vmedia policy name
        $NewVMediaName = "vMediaPolicy-Temp-$($SelectedServerProfile.Name)-$RandomNumber"

        # Gets the default Org
        $Org = Get-IntersightOrganizationOrganization -Name default

        # Change MS notation to Unix
        $InputPath = ($InputPath.Replace("\","/")).Trim()

        # If user got path from file, change it
        if ($InputPath -match "//file.nchosting.dk/Software/")
        {
            $InputPath = ($inputpath -split "//file.nchosting.dk/Software/")[1]
        }

        # Correct users who does not use file endings
        if ($IsoName -notmatch ".iso")
        {
            $IsoName = $IsoName + ".iso"
        }

        # Modifies vmedia mappings
        $vMediaParams = @{
            VolumeName             = "$($SelectedServerProfile.Name)-$RandomNumber"
            FileLocation           = "nfs://5.44.142.53/SoftwareNFS/$InputPath/$ISOName"
            MountOptions           = "Retry=10"
            DeviceType             = "Cdd"
            MountProtocol          = "Nfs"
            AuthenticationProtocol = "None"
            ClassId                = "VmediaMapping"
            ObjectType             = "VmediaMapping"
        }
        $NewVmediaSettings = Initialize-IntersightVmediaMapping @vMediaParams

        # Creates a new random vmedia policy
        $PolicyParams = @{
            Name         = $NewVMediaName
            Organization = $Org
            Mappings     = $NewVmediaSettings
        }
        $NewVmediaPolicy = New-IntersightVmediaPolicy @PolicyParams

        # Generates moref for new vmedia policy
        $vMediaPolicyMoref = $NewVmediaPolicy | Get-IntersightMoMoRef

        # Deploys new vMediaPolicy
        $ActualInstance = $SelectedServerProfile.PolicyBucket.ActualInstance

        $KvmPolicy     = $ActualInstance | Where-Object {$_.ObjectType -eq "KvmPolicy"}
        $BiosPolicy    = $ActualInstance | Where-Object {$_.ObjectType -eq "BiosPolicy"}

        # Racks servers don't need to have all polices populated
        if ($Type -eq "ComputeRackUnit")
        {
            $PolicyBucket = @(
                $KvmPolicy,
                $vMediaPolicyMoref
            )

            $MountParams = @{
                Moid                 = ($SelectedServerProfile.Moid)
                Action               = "Deploy"
                PolicyBucket         = $PolicyBucket
                ServerAssignmentMode = "None"
            }

        }
        elseif ($type -eq "ComputeBlade")
        # But Blade servers needs to be populated with their entire existing config
        {
            $vNicPolicy    = $ActualInstance | Where-Object {$_.ObjectType -eq "VnicLanConnectivityPolicy"}
            $vNicSanPolicy = $ActualInstance | Where-Object {$_.ObjectType -eq "VnicSanConnectivityPolicy"}
            $BootPolicy    = $ActualInstance | Where-Object {$_.ObjectType -eq "BootPrecisionPolicy"}
            $AccessPolicy  = $ActualInstance | Where-Object {$_.ObjectType -eq "AccessPolicy"}

            $PolicyBucket = @(
                $BootPolicy,
                $vNicPolicy,
                $VnicSanPolicy,
                #$vMediaPolicyMoref,
                $AccessPolicy,
                $KvmPolicy,         # From line 111
                $BiosPolicy         # From line 112

            )

            $MountParams = @{
                Moid                 = ($SelectedServerProfile.Moid)
                Action               = "Deploy"
                PolicyBucket         = $PolicyBucket
                ServerAssignmentMode = "Static"
            }
        }
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not prepare vMedia"
        $Error
    }

    # Push the server profile to the server
    try
    {
        Write-Host "Mounting .iso ... " -NoNewline
        [void]::(Set-IntersightServerProfile @MountParams)
        Write-Host "OK"

    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        $Error
    }

    if ($Sleep)
    {
        # Wait for the to remove
        Write-Host "Sleeping for 15 seconds, to make sure policy is mounted ... " -NoNewline
        Start-Sleep -Seconds 15
        Write-Host "OK"
    }
}
<#
  _   _                      _   _      _____       _                _       _     _    _____ _                   _     _____            __ _ _
 | \ | |                    | \ | |    |_   _|     | |              (_)     | |   | |  / ____| |                 (_)   |  __ \          / _(_) |
 |  \| | _____      ________|  \| | ___  | |  _ __ | |_ ___ _ __ ___ _  __ _| |__ | |_| |    | |__   __ _ ___ ___ _ ___| |__) | __ ___ | |_ _| | ___
 | . ` |/ _ \ \ /\ / /______| . ` |/ __| | | | '_ \| __/ _ \ '__/ __| |/ _` | '_ \| __| |    | '_ \ / _` / __/ __| / __|  ___/ '__/ _ \|  _| | |/ _ \
 | |\  |  __/\ V  V /       | |\  | (__ _| |_| | | | ||  __/ |  \__ \ | (_| | | | | |_| |____| | | | (_| \__ \__ \ \__ \ |   | | | (_) | | | | |  __/
 |_| \_|\___| \_/\_/        |_| \_|\___|_____|_| |_|\__\___|_|  |___/_|\__, |_| |_|\__|\_____|_| |_|\__,_|___/___/_|___/_|   |_|  \___/|_| |_|_|\___|
                                                                        __/ |
                                                                       |___/
#>
Function New-NCIntersightChassisProfile
{
    # Don't continue without a connection to Intersight
    Confirm-IntersightConnection

    # Retrieve a list of all chassis and chassisprofiles in Intersight
    $AllIntersightChassisProfiles = Get-IntersightChassisProfile
    $AllIntersightChassis         = Get-IntersightEquipmentChassis

    # Filter to include only thouse that have been unconfigured
    $UnconfiguredChassis = $AllIntersightChassis | Where-Object {$_.Moid -notin $AllIntersightChassisProfiles.AssignedChassis.ActualInstance.Moid}

    if ($UnconfiguredChassis.count -gt 0)
    {
        Write-Host "Found $($UnconfiguredChassis.Count) unconfigured chassis with the serials below:"
        $UnconfiguredChassis.Serial
        Write-Host ""

        # Gather policies identical to all chassis
        $PowerPolicy   = Get-IntersightPowerPolicy   -Name "UCS-Power-Profile-X9508" | Get-IntersightMoMoRef
        $ThermalPolicy = Get-IntersightThermalPolicy -Name "UCS-Thermal-Policy"      | Get-IntersightMoMoRef
        $Organization  = Get-IntersightOrganizationOrganization -name "default"      | Get-IntersightMoMoRef

        # Configure each remaining chassis
        foreach ($Chassis in $UnconfiguredChassis)
        {
            Write-Host "Configuring $($Chassis.Serial) ... " -NoNewline

            # Prepare Parameters for the deployment of Chassis Profile
            $ProfileParams = @{
                AssignedChassis = $Chassis | Get-IntersightMoMoRef
                Name            = $Chassis.Name
                Organization    = $Organization
                Action          = "Deploy"
                TargetPlatform  = "FIAttached"
                PolicyBucket    = @(
                    $PowerPolicy,
                    $ThermalPolicy
                )
            }

            # Then apply it to the chassis
            try
            {
                # Create the Profile
                $Creation = new-IntersightChassisProfile @ProfileParams
                Start-Sleep -seconds 2

                # And manually push it. This should not be necessary
                [void]::(set-IntersightChassisProfile @ProfileParams -Moid $Creation.moid)

                Write-Host "OK"
            }
            catch
            {
                Write-Host "FAIL!" -BackgroundColor Red
                Write-Host "Could not create or push profile"
                $Error
                Exit 1
            }
        }
    }
}

<#
  _   _                      _____                          _____            __ _ _
 | \ | |                    / ____|                        |  __ \          / _(_) |
 |  \| | _____      _______| (___   ___ _ ____   _____ _ __| |__) | __ ___ | |_ _| | ___
 | . ` |/ _ \ \ /\ / /______\___ \ / _ \ '__\ \ / / _ \ '__|  ___/ '__/ _ \|  _| | |/ _ \
 | |\  |  __/\ V  V /       ____) |  __/ |   \ V /  __/ |  | |   | | | (_) | | | | |  __/
 |_| \_|\___| \_/\_/       |_____/ \___|_|    \_/ \___|_|  |_|   |_|  \___/|_| |_|_|\___|


#>
Function New-ServerProfile
{

    #region------------------------------------------| HELP |------------------------------------------------#
    <#
        .Synopsis
            Creates a new intersight service profile
        .PARAMETER FQDN
            Fully qualified domain name of the server you wish to create
        .PARAMETER BladeSerial
            Serial Number of the physical blade you are working on
        .PARAMETER OS
            ESXi, Linux or Windows
        .PARAMETER VLAN
            VLAN of the physical server. Will no do anything for ESXi Hosts
    #>
    #endregion

    #region---------------------------------------| PARAMETERS |---------------------------------------------#
    # Set parameters for the script here
    param
    (
        [Parameter(Mandatory = $true)]
        [String]
        $FQDN,

        [Parameter(Mandatory = $true)]
        [String]
        $BladeSerial,

        [Parameter(Mandatory = $true)]
        [String]
        [ValidateSet("vmware","linux","windows")]
        $OS,

        [Parameter(Mandatory = $false)]
        [String]
        $VLAN = "n/a"
    )
    #endregion
    #---------------------------------------------| For Humans |---------------------------------------------#

    Write-Host "Building with the following details:"

    $InitialInfo = [pscustomobject]@{
        FQDN        = $FQDN
        BladeSerial = $BladeSerial
        OS          = $OS
        vLAN        = $Vlan
    }

    # Write info to the user
    $InitialInfo | Format-Table -Wrap
    Write-Host ""

    # ---- Odd / Even ----
    if ($FQDN -Match "esx")
    {
        $OddEven = if (($FQDN -replace "[^0-9]" , '') % 2 -eq 0 ) {"even"} else {"odd"}
    } else
    {
        $OddEven = "all"
    }

    Write-HostSeperator "Paramater Preperation"

    # ---- Customer / Solution ----
    $Customer = $FQDN.Split("-")[0]
    $Solution = $FQDN.Split("-")[1]
    Write-Host "Customer: $Customer"
    Write-Host "Solution: $Solution"

    # ---- Select blade ----
    $AllBlades       = Get-IntersightComputeBlade
    $SelectedBlade   = $AllBlades     | Where-Object {$_.Serial -match $BladeSerial -and $_.MgmtIPAddress -eq ""}
    $SelectedBladeMo = $SelectedBlade | Get-IntersightMoMoRef

    if ($SelectedBladeMo)
    {
        Write-Host "Found Blade with Serial $($SelectedBlade.Serial)"
    }
    else
    {
        Write-Host "Could not find blade; exiting"
        throw
    }

    #region ---- Get boot policy and template ----
    $Splits = $SelectedBlade.Name.Split("-")
    $IntersightChassis = get-IntersightEquipmentChassis -name ($Splits[0..2] -join "-")

    # Get location tags
    $LocationTag   = ($IntersightChassis).Tags
    $Templates     = Get-IntersightServerProfileTemplate | Where-Object {$_.Name -match "$OS"}
    $IntersightOrg = Get-IntersightOrganizationOrganization -name "default"

    Write-Host "Template: $($Templates.Name) Selected"

    # Select boot policy
    $DCLocation  = ($LocationTag | Where-Object {$_.Key -eq "location"}).Value
    $PodLocation = ($LocationTag | Where-Object {$_.Key -eq "pod"}).Value

    Write-Host "DC Location: $DCLocation"
    Write-Host "PodLocation: $PodLocation"

    $BootPolicySelection = Get-IntersightBootPrecisionPolicy | Where-Object {
        $_.Tags.Value -contains $DCLocation  -and
        $_.Tags.Value -contains $PodLocation -and
        $_.Tags.Value -contains $OddEven     -and
        $_.Tags.Value -contains $os}

    $BootPolicyMoRef = $BootPolicySelection | Get-IntersightMoMoRef

    if(($null -eq $BootPolicySelection) -or ($null -eq $Templates))
    {
        Write-Host "Unable to locate a boot policy or template, ending script"
        throw
    }
    else
    {
        Write-Host "Found Boot Policy: $($BootPolicySelection.Name) and template: $($Templates.Name)"
    }
    #endregion

    # Select vNics
    if($FQDN -match "esx")
    {
        Write-Host "ESXi detected - fetching ESXi LAN Connectivity policy"
        $LanConnectivityPolicySelection = Get-IntersightVnicLanConnectivityPolicy | Where-Object {
            $_.Tags.Value -contains $DCLocation  -and
            $_.Tags.Value -contains $PodLocation -and
            $_.Tags.Value -contains $OS }
    }
    else
    { # Physical blade

        # VLAN Creation
        try
        {
            Write-Host "$OS Server detected, creating network settings"
            $VlanSettings = Initialize-IntersightFabricVlanSettings -AllowedVlans $Vlan -NativeVlan $Vlan
            Write-Host "Vlan settings created for vlan: $VLAN"
        }
        catch
        {
            Write-Host "Unable to create VLAN settings, ending script"
            Throw
        }

        # vNic creation
        try
        {
            Write-Host "$OS Server detected, creating network settings"
            $LanConnectivityPolicySelection = New-IntersightVnicLanConnectivityPolicy -SwitchId "A" -
            Write-Host "Vlan settings created for vlan: $VLAN"
        }
        catch
        {
            Write-Host "Unable to create VLAN settings, ending script"
            Throw
        }

        $NetworkPolicy      = New-IntersightFabricEthNetworkGroupPolicy   -Name $FQDN -Organization $IntersightOrg -VlanSettings $VlanSettings
        $NetworkPolicyMoref = $NetworkPolicy | Get-IntersightMoMoRef
        $AdapterPolicy      = Get-IntersightVnicEthAdapterPolicy          -Name "UCS-ETH-ADAPTER-POLICY-GLOBAL"
        $ControlPolicy      = Get-IntersightFabricEthNetworkControlPolicy -Name "UCS-Network-Control-Policy"
        $QosPolicy          = Get-IntersightVnicEthQosPolicy              -Name "UCS-ETH-QOS-Policy-global"
        $MacPool            = Get-IntersightMacpoolPool | Where-Object {$_.Name -match "$dclocation-$podlocation"}
        $PlacementSettings  = Initialize-IntersightVnicPlacementSettings  -PciLink 0 -Id "MLOM" -ObjectType "VnicPlacementSettings" -ClassId "VnicPlacementSettings" -SwitchId "A" -

        $LanConnectivityPolicySelection = Get-IntersightVnicLanConnectivityPolicy -Moid "64d50106b714fd1a1e9c808f"

        # Create eth01 vnic
        $vNicParams = @{
            Name                          = "eth1"
            EthAdapterPolicy              = $AdapterPolicy
            EthQosPolicy                  = $QosPolicy
            FabricEthNetworkGroupPolicy   = $NetworkPolicyMoref
            MacPool                       = $MacPool
            Placement                     = $PlacementSettings
            LanConnectivityPolicy         = $LanConnectivityPolicySelection
            FabricEthNetworkControlPolicy = $ControlPolicy
            FailoverEnabled               = $true
            Order                         = 4
        }

        [void]::(New-IntersightVnicEthIf @vNicParams)
    }

    if($null -eq $LanConnectivityPolicySelection)
    {
        Write-Host "Unable to locate LAN Policy, ending script"
    }
    else
    {
        Write-Host "Found Lan Connectivity Policy: $($LanConnectivityPolicySelection.Name)"
    }

    $LanConnectivityPolicyMoref = $LanConnectivityPolicySelection | Get-IntersightMoMoRef

    # Select SAN config
    if($FQDN -match "esx")
    {
        $SANConfigurationSelection  = Get-IntersightVnicSanConnectivityPolicy | Where-Object {
        $_.Tags.Value -contains $DCLocation -and
        $_.Tags.Value -contains $PodLocation -and
        $_.Tags.Value -contains "vmware"}
    }
    else
    {
        $SANConfigurationSelection  = Get-IntersightVnicSanConnectivityPolicy | Where-Object {
        $_.Tags.Value -contains $DCLocation -and
        $_.Tags.Value -contains $PodLocation -and
        $_.Tags.Value -contains "all"}
    }

    $SANConfigurationMoref = $SANConfigurationSelection | Get-IntersightMoMoRef
    if ($null -ne $SANConfigurationMoref)
    {
        Write-Host "Found SAN Connectivity Policy: $($SANConfigurationSelection.Name)"
    }

    # Select vMedia
    $vMediaPolicy = Get-IntersightVmediaPolicy | Where-Object {
        $_.Tags.Value -contains $PodLocation -and
        $_.Tags.Value -contains $OS}

    $vMediaPolicyMoref = $vMediaPolicy | Get-IntersightMoMoRef
    if ($null -ne $vMediaPolicyMoref)
    {
        Write-Host "Found vMedia Policy: $($vMediaPolicy.Name)"
    }
    else
    {
        Write-Host "Could not find vMedia Policy"
    }

    # Select KVM
    $KvmPolicy         = Get-IntersightKvmPolicy
    $KvmPolicyMoref    = $KvmPolicy | Get-IntersightMoMoref

    # Select IMC
    $IMCPolicy         = Get-IntersightAccessPolicy -Name "UCS-IMC-Policy"
    $IMCPolicyMoref    = $IMCPolicy | Get-IntersightMoMoref

    #Select UUID Pool
    $UUIDPool          = Get-IntersightUuidpoolPool -Name "UCS-UUID-POOL"
    $UUIDPoolMoref     = $UUIDPool | Get-IntersightMoMoref

    #Select Bios policy
    $BiosPolicy        = Get-IntersightBiosPolicy -Name "default-bios-policy"
    $BiosPolicyMoref   = $BiosPolicy | Get-IntersightMoMoref

    # Create policy Bucket
    $PolicyBucket = @(
        $BootPolicyMoRef,
        $SANConfigurationMoref,
        $LanConnectivityPolicyMoref,
        $KvmPolicyMoref,
        $IMCPolicyMoref,
        $BiosPolicyMoref
    )

    # And add the vMediaPolicy, if its an esxhost
    if($FQDN -match "esx")
    {
        $PolicyBucket += $vMediaPolicyMoref
    }

    # Write data to console
    [pscustomobject]@{
        FQDN         = $FQDN
        Serial       = $SelectedBlade.Serial
        Vlan         = $Vlan
        DCLocation   = $DCLocation
        PodLocation  = $PodLocation
        OddEven      = $OddEven
        OS           = $OS
        BootPolicy   = $($BootPolicySelection.name)
        Org          = $($IntersightOrg.name)
        SANSelection = $SANConfigurationSelection.Tags.Value -join " "
    }

    Write-HostSeperator "Attempt serverprofile creation"
    try
    {
        $Params = @{
            Name                 = $FQDN
            PolicyBucket         = $PolicyBucket
            ConfigContext        = $Templates.ConfigContext
            Uuidpool             = $UUIDPoolMoref
            Organization         = $IntersightOrg
            Tags                 = $LocationTag
            #AssignedServer       = $SelectedBladeMo
            #ServerAssignmentMode = "Static"
            ServerAssignmentMode = "None"
            TargetPlatform       = "FIAttached"
            Action               = "Deploy"
        }

        Write-Host "Creating Server profile ... " -NoNewline
        [void]::(New-IntersightServerProfile @params)
        Write-Host "Ok"

        Write-Host "Allowing Server profile to propergate .." -NoNewline
        1..20 | ForEach-Object {
            Write-host "." -NoNewline
            Start-sleep -Seconds 15
        }
        Write-Host " OK"

        Write-host "Deploying Server profile to $BladeSerial ... " -NoNewline
        [void]::(get-intersightserverprofile -name $FQDN | Set-IntersightServerProfile -Action "Deploy" -ServerAssignmentMode "Static" -AssignedServer $SelectedBladeMO)
        Write-Host "Done"


        Write-Host "Waiting for $BladeSerial to get Server profile assigned .." -NoNewline
        1..20 | ForEach-Object {
            Write-host "." -NoNewline
            Start-sleep -Seconds 15
        }
        Write-Host " OK"

    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        $Error[0]
    }
    #endregion

    #region---------------------------------------| DISCONNECT |---------------------------------------------#
    Write-Host "The script has finished running: Closing"

    If ($Error)
    {
        Throw
    }
    #endregion
    #-------------------------------------------------| END |------------------------------------------------#
}
<#
   _____      _        _____       _                _       _     _   ____  _           _      _           _          _
  / ____|    | |      |_   _|     | |              (_)     | |   | | |  _ \| |         | |    | |         | |        | |
 | (___   ___| |_ ______| |  _ __ | |_ ___ _ __ ___ _  __ _| |__ | |_| |_) | | __ _  __| | ___| |     __ _| |__   ___| |
  \___ \ / _ \ __|______| | | '_ \| __/ _ \ '__/ __| |/ _` | '_ \| __|  _ <| |/ _` |/ _` |/ _ \ |    / _` | '_ \ / _ \ |
  ____) |  __/ |_      _| |_| | | | ||  __/ |  \__ \ | (_| | | | | |_| |_) | | (_| | (_| |  __/ |___| (_| | |_) |  __/ |
 |_____/ \___|\__|    |_____|_| |_|\__\___|_|  |___/_|\__, |_| |_|\__|____/|_|\__,_|\__,_|\___|______\__,_|_.__/ \___|_|
                                                       __/ |
                                                      |___/
#>
Function Set-IntersightBladeLabel
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [String]
        $FQDN,

        [Parameter(Mandatory = $true)]
        [String]
        $UsrLbl
    )

    # Don't continue without a connection to Intersight
    Confirm-IntersightConnection

    # Gather values for the user label
    $ServerProfile = Get-IntersightServerProfile -Name $FQDN
    $Tags          = Initialize-IntersightMoTag -Key State -Value $UsrLbl
    $CompleteTag   = ($ServerProfile.Tags | Where-Object {$_.Key -ne 'State'}) + $Tags

    # Push label to blade
    try
    {
        Write-Host "Setting Intersight State tag to `"$UsrLbl`" on $FQDN ... " -NoNewLine
        [void](Set-IntersightServerProfile -Moid $ServerProfile.Moid -Tags $CompleteTag)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor "Red"
        Write-Host "Could not push label to blade"
        $Error
        Exit 1
    }
}
<#
   _____ _             _        _____       _                _       _     ___      ____  __ _               _
  / ____| |           | |      |_   _|     | |              (_)     | |   | \ \    / /  \/  | |             | |
 | (___ | |_ __ _ _ __| |_ ______| |  _ __ | |_ ___ _ __ ___ _  __ _| |__ | |\ \  / /| \  / | |__   ___  ___| |_
  \___ \| __/ _` | '__| __|______| | | '_ \| __/ _ \ '__/ __| |/ _` | '_ \| __\ \/ / | |\/| | '_ \ / _ \/ __| __|
  ____) | || (_| | |  | |_      _| |_| | | | ||  __/ |  \__ \ | (_| | | | | |_ \  /  | |  | | | | | (_) \__ \ |_
 |_____/ \__\__,_|_|   \__|    |_____|_| |_|\__\___|_|  |___/_|\__, |_| |_|\__| \/   |_|  |_|_| |_|\___/|___/\__|
                                                                __/ |
                                                               |___/
#>
Function Start-IntersightVMhost
{
    param
    (
        [Parameter(Mandatory = $true)]
        [Intersight.Model.ServerProfile]
        $ServerProfile,

        [Parameter(Mandatory = $true)]
        [String]
        $PowerOnLabel
    )

    if ($ServerProfile.Name -notmatch "esx")
    {
        Write-Error "esx server profile was not input"
        throw
    }

    $Name = ($serverprofile.name -split "\.")[0]
    Write-HostSeperator "Powering On $Name" -width 45

    $Tags        = Initialize-IntersightMoTag -Key "State" -Value $PowerOnLabel
    $CompleteTag = ($ServerProfile.Tags | Where-Object {$_.Key -ne 'State'}) + $Tags
    $Blade       = Get-IntersightComputeBlade -Moid $ServerProfile.AssignedServer.ActualInstance.moid

    # 1: Set the tags on the server profile itself
    try
    {
        Write-Host "Updating tags on ServerProfile in Intersight ... " -NoNewline
        [Void]::(Set-IntersightServerProfile -Moid $ServerProfile.Moid -Tags $CompleteTag)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
    }

    # 2: Sets state in VMware
    try
    {
        $VMhost = Get-VMhost $ServerProfile.Name

        Write-Host "Updating tags on VMhost in vCenter ... " -NoNewline
        $VMhost.ExtensionData.setCustomValue('State', $PowerOnLabel)
        $VMhost.ExtensionData.setCustomValue('StateDate', $(Get-FormattedTime))
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not update Tag on VMhost."
        throw $error
    }


    # 3: Power on the blade
    if ($Blade.OperPowerState -ne "on")
    {
        try
        {
            Write-Host "Sending PowerON signal ... " -NoNewline
            [Void]::(Set-IntersightServerProfile -Moid $ServerProfile.Moid -Action "Deploy")
            Write-Host "OK"
        }
        catch
        {
            Write-Host "FAIL!" -BackgroundColor Red
            Write-Host "Could not send powerON signal"
            $Error[0]
            throw
        }
    }
}
function Add-NCDatastore
{
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ComputeCluster,

        [Parameter(Mandatory = $true)]
        [string]
        $AllDatastoreNamesString,

        [Parameter()]
        [int]
        $BootLunSizeGB = 10,

        [Parameter()]
        [int]
        $VMFSVersion = 6,

        [Parameter(,
        ParameterSetName="hip")]
        [boolean]
        $hip,

        [Parameter(ParameterSetName="hip")]
        [string]
        $FQDN
        
    )   
    
    #Splits DS Names into an array
    $AllDatastoreNames = $AllDatastoreNamesString.Split(",")

    #Validates if a connection to vCenters exists
    Confirm-vCenterConnection
    Write-Host "Searching for host in $ComputeCluster"
    #Loops through all datastores
    Foreach ($Datastore in $AllDatastoreNames.Trim())
    {
        Write-Host "---- Trying to create $Datastore"
        
        #Gets first host that isn't a buffer and has most CPU capacity
        $VMHost = Get-Cluster $ComputeCluster      | 
            Get-VMHost -State "Connected"          | 
            Where-Object {$_.Name -notmatch "buf"} | 
            Sort-Object CpuUsageMhz                |  
            Select-Object -First 3                 | 
            Get-Random                             
        Write-Host "VMHost: $VMHost selected"
        Write-Host "Rescanning host to find newly created datastores"
        [void]($VMHost | Get-VMHostStorage -RescanAllHba -RescanVmfs)
        #Gets all currently added datastores from the host
        $AllCreatedDatastores = $VMHost | Get-Datastore

        #Extract LUN names from datastores
        $LUNnames = $AllCreatedDatastores.ExtensionData.Info.vmfs.Extent.DiskName
    
        #Collects and matches all LUNs against already created datastores - selects all not created LUNs where size is larger than boot lun size
        $AllUnallocatedLuns = Get-ScsiLun -VMHost $VMHost | Where-Object {{$_.CanonicalName -notin $LUNnames} -and {$_.CapacityGB -gt $BootLunSizeGB}}
        
        Write-Host "$($AllUnallocatedLuns.count) LUNs connected to the VMhost"
        Write-Host "Attempting to locate LUN matching to label $Datastore"
        
        #Matches the LUN against the disk label and selects the correct LUN
        $SelectedLUN = $AllUnallocatedLuns | Where-Object {$_.CanonicalName -match $Datastore.Split("-")[1].Substring(1,4)}
        
        #Valides if a LUN was selected
        if($null -ne $SelectedLUN)
        {
            Write-Host "LUN deteched with CanonicalName:"
            Write-Host $($SelectedLun.CanonicalName)
            Write-Host "Desired label: $Datastore"
        }
        else
        {
            Write-Host "Unable to detected a LUN matching $Datastore"
            Write-Host "Ending script"
            Exit
        }
    
        #Creates a new datastore
        try 
        {
            if($hip -eq $true)
            {
            Write-Host "High performance LUN"
                $DatastoreParams = @{
                    Name              = $Datastore 
                    Path              = $SelectedLUN.CanonicalName 
                    FileSystemVersion = $VMFSVersion 
                    Vmfs              = $true
                    Confirm           = $false
                }
            }
            else{
            Write-Host "Creating Datastore $Datastore ... " -NoNewline
            $DatastoreParams = @{
                Name              = $Datastore 
                Path              = $SelectedLUN.CanonicalName 
                FileSystemVersion = $VMFSVersion 
                Vmfs              = $true
                Confirm           = $false
            }
            
            }
            [void]($VMHost | New-Datastore @DatastoreParams)
            #Validates if datastore was created
            $CreatedDatastore = Get-Datastore $Datastore
        
            if($null -ne $CreatedDatastore)
            {
                Write-Host "OK"
            }
            else
            {
                Write-Host "FAIL!" -BackgroundColor "Red"
                Write-Host "Unable to create DS with label $Datastore"
                Exit
            }
    
        }
        catch
        {
            Write-Host "FAIL!" -BackgroundColor "Red"
            Write-Host "$Datastore Could not be created"
            $error 
            Exit
        }
    } 

    #Moves datastore(s) into specified cluster
    try 
    {
        if($hip -eq $true)
        {
            Write-Host "hip : true"
            Write-Host "Fetching Datastore cluster for $FQDN"            
            #Datastore clusters are named the same as their compute counterparts
            $TargetCluster = Get-DatastoreCluster $FQDN -ErrorAction SilentlyContinue
            if($TargetCluster.Name -ne $FQDN)
            {
                Write-Host "No HIP Cluster detected with name: $FQDN"
                Write-Host "Creating HIP Cluster: $FQDN"
                New-DatastoreCluster -Name $FQDN -Confirm:$false -Location (Get-Folder "HighPerformanceDatastores")
            }
                Write-Host "Moving datastores into HIP Datastore Cluster $FQDN ... " -NoNewline
                $TargetCluster = Get-DatastoreCluster $FQDN
                [void](Get-Datastore $AllDatastoreNames | Move-Datastore -Destination $TargetCluster -Confirm:$false)
                Write-Host "OK"
        }
        else{
            Write-Host "Moving datastores into $ComputeCluster ... " -NoNewline
            #Datastore clusters are named the same as their compute counterparts
            $TargetCluster = Get-DatastoreCluster $ComputeCluster
            
            [void](Get-Datastore $AllDatastoreNames | Move-Datastore -Destination $TargetCluster -Confirm:$false)
            Write-Host "OK"
        }    
    }
    catch 
    {
        Write-Host "FAIL!" -BackgroundColor "Red"
        $Error
        Exit
    }
    
    #Rescans all connected hosts in cluster
    Write-Host "Rescanning $ComputeCluster ... " -NoNewline
    [void](Get-Cluster $ComputeCluster | Get-VMHost -State Connected | Get-VMHostStorage -RescanAllHba -RescanVmfs)
    Write-Host "OK"
        
}

<#
   _____                      _      _              __  __       _       _                                  __  __           _
  / ____|                    | |    | |            |  \/  |     (_)     | |                                |  \/  |         | |
 | |     ___  _ __ ___  _ __ | | ___| |_ ___ ______| \  / | __ _ _ _ __ | |_ ___ _ __   __ _ _ __   ___ ___| \  / | ___   __| | ___
 | |    / _ \| '_ ` _ \| '_ \| |/ _ \ __/ _ \______| |\/| |/ _` | | '_ \| __/ _ \ '_ \ / _` | '_ \ / __/ _ \ |\/| |/ _ \ / _` |/ _ \
 | |___| (_) | | | | | | |_) | |  __/ ||  __/      | |  | | (_| | | | | | ||  __/ | | | (_| | | | | (_|  __/ |  | | (_) | (_| |  __/
  \_____\___/|_| |_| |_| .__/|_|\___|\__\___|      |_|  |_|\__,_|_|_| |_|\__\___|_| |_|\__,_|_| |_|\___\___|_|  |_|\___/ \__,_|\___|
                       | |
                       |_|
#>
function Complete-MaintenanceMode
{
    param (
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost,

        [Parameter()]
        [int]
        $MaintenanceTimer = 15
    )

    # Refresh Custom Attributes
    $CustomAttributes = Get-VMhostCustomAttributes -VMHost $VMhost
    $CurrentDate      = Get-FormattedTime
    $ShortName        = ($VMhost.Name -split "\.")[0]

    $MaintenanceTaskTime = [Math]::Abs($(New-TimeSpan -Start $($CustomAttributes.StateDate) -End $CurrentDate).TotalMinutes)
    $MaintenanceTask     = Get-Task -id $CustomAttributes.MaintenanceTaskId -ErrorAction SilentlyContinue

    # There is a maintenance task entirely
    if ($MaintenanceTask.State -eq "Running")
    {
        # ... and VMhost has attempted to enter maintenance for more than 15 minutes, help it along
        if ($MaintenanceTaskTime -ge $MaintenanceTimer)
        {
            Write-Verbose "Time is: $CurrentDate. $ShortName has current State Date: $($CustomAttributes.StateDate)   -   TimeSpan is $MaintenanceTaskTime"
            Write-Host "$Shortname has tried entering maintenance mode for $([Math]::Round($MaintenanceTaskTime,2)) minutes out of tolerated $MaintenanceTimer minutes"

            # Get all VMs on the host that are still remaining. Start with the smallest
            $RemainingVMsOnHost = Get-VMHost -Name $VMhost.Name | Get-VM | Sort-Object -Property "MemoryGB"

            if ($RemainingVMsOnHost)
            {
                # Notify user of action
                Write-Host "$ShortName has attempted to enter maintenance for more than $MaintenanceTimer minutes, assisting with VM Migration of the last $($RemainingVMsOnHost.count) VMs:"

                # Get all *other* hosts in the cluster, that should accept VMs at this time.
                $ClusterHosts = $RemainingVMsOnHost[0]  | Get-Cluster | Get-VMHost | Where-Object {
                    $_.ConnectionState -eq       "connected"  -and
                    $_.name            -ne       $VMhost.name -and
                    $_.name            -notmatch "buf"        -and
                    $_.Customfields["State"] -eq "live"
                }

                # Move all VMs to other hosts in the cluster
                Foreach ($SingleVM in ($RemainingVMsOnHost))
                {
                    # Move single VM. Void output to not clutter console
                    # Move to random applicable VMhost to better utilize cluster resources and minimize additional vMotions
                    try
                    {
                        $Destination = ($ClusterHosts | Get-Random).Name

                        $MoveVariables = @{
                            VM               = $SingleVM.Name
                            Destination      = $Destination
                            VMotionPriority  = "High"
                            RunAsync         = $true
                        }
                        Write-Host "  $($SingleVM.Name) will be moved to $Destination"
                        [void]::(Move-VM @MoveVariables)

                    }
                    # Could not move the VM. Notify of the error
                    catch
                    {
                        Write-Host "Could not move $SingleVM"
                    }
                }
            }
            else
            {
                Write-Host "$ShortName has been emptied of VMs"
            }
        }
        # VMhost has spent less than 15 minutes entering maintenance. Give it more time
        else
        {
            # Notify user of Aciton
            Write-Host "Still waiting for $ShortName to enter Maintenancemode"
        }
    }
    elseif ($MaintenanceTask.State -eq "Failed")
    {
        Write-Error "An error occurred on $Shortname!"
    }
    else
    {
        Write-Error "There is not Maintenance Task found on $Shortname. Setting to Live state"
        Write-Host "Correcting tags ... " -NoNewline
        $VMhost.ExtensionData.setCustomValue('State', "live")
        $VMhost.ExtensionData.setCustomValue('StateDate', (Get-FormattedTime))
        Write-Host "OK"
    }
}
<#
   _____             __ _                             _____           _             _____                            _   _
  / ____|           / _(_)                           / ____|         | |           / ____|                          | | (_)
 | |     ___  _ __ | |_ _ _ __ _ __ ___ ________   _| |     ___ _ __ | |_ ___ _ __| |     ___  _ __  _ __   ___  ___| |_ _  ___  _ __
 | |    / _ \| '_ \|  _| | '__| '_ ` _ \______\ \ / / |    / _ \ '_ \| __/ _ \ '__| |    / _ \| '_ \| '_ \ / _ \/ __| __| |/ _ \| '_ \
 | |___| (_) | | | | | | | |  | | | | | |      \ V /| |___|  __/ | | | ||  __/ |  | |___| (_) | | | | | | |  __/ (__| |_| | (_) | | | |
  \_____\___/|_| |_|_| |_|_|  |_| |_| |_|       \_/  \_____\___|_| |_|\__\___|_|   \_____\___/|_| |_|_| |_|\___|\___|\__|_|\___/|_| |_|

  #>
Function Confirm-vCenterConnection
{
    # Check if current connection
    if($null -eq $global:DefaultVIServer)
    {
        Write-Host "Not connected to any vCenters, ending"
        throw
    }
    else
    {
        $VMhostTest = [void]::(Get-VMHost | get-random -ErrorAction SilentlyContinue)
        if ($VMhostTest)
        {
            Write-Verbose "vCenterConnection Confirmed"
        }
        else
        {
            Write-Host "vCenterConnection cannot be confirmed. Attempting reconnect"
        }
    }
}

<#
   _____                            _          _   _             _____           _
  / ____|                          | |        | \ | |           / ____|         | |
 | |     ___  _ __  _ __   ___  ___| |_ ______|  \| | _____   _| |     ___ _ __ | |_ ___ _ __
 | |    / _ \| '_ \| '_ \ / _ \/ __| __|______| . ` |/ __\ \ / / |    / _ \ '_ \| __/ _ \ '__|
 | |___| (_) | | | | | | |  __/ (__| |_       | |\  | (__ \ V /| |___|  __/ | | | ||  __/ |
  \_____\___/|_| |_|_| |_|\___|\___|\__|      |_| \_|\___| \_/  \_____\___|_| |_|\__\___|_|


  #>
Function Connect-NCvCenter
{
    Param
    (
        [Parameter(Mandatory = $True)]
        [pscredential]
        $vCenterCredential,

        [Parameter()]
        [String]
        $vCenter = 'vcenter01.nchosting.dk',

        [Parameter()]
        [System.Boolean]
        $Force = $false
    )

    #Connects to the vCenter
    try
    {
        Write-Host 'Connecting to vCenter(s) ... ' -NoNewline

        # Manually set Security protocol
        [void]::([System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12)

        # Check if current connection
        if(($null -eq $global:DefaultVIServer) -or ($force))
        {
            $DefaultParams = @{
                Scope   = "User"
                Confirm = $false
            }

            # Don't display the Customer Experience Improvement Programme textbox, always allow multipe vCenters,
            # Don't care about invalid certificates, and always show deprecated commands warning
            Set-PowerCLIConfiguration @DefaultParams -ParticipateInCEIP:$false           | Out-Null
            Set-PowerCLIConfiguration @DefaultParams -DefaultVIServerMode "Multiple"     | Out-Null
            Set-PowerCLIConfiguration @DefaultParams -InvalidCertificateAction "Ignore"  | Out-Null
            Set-PowerCLIConfiguration @DefaultParams -DisplayDeprecationWarnings:$true   | Out-Null

            # Build connection paramaters for connection
            $Params = @{
                Server     = $vCenter
                AllLinked  = $True
                Credential = $vCenterCredential
            }

            # Finish the connection
            [void]::(Connect-VIserver @Params)
            Write-Host 'OK'
        }
        else
        {
            Write-Host "FAIL" -BackgroundColor "Yellow"
            Write-Host "You're already connected to the viServers below:"
            ($global:DefaultVIServers).name
        }
    }
    catch
    {
        Write-Host 'FAIL' -BackgroundColor Red
        Write-Host 'Could not connect to vCenter. Exiting'
        throw
    }
}

Function Get-AllVMhosts
{
    Param
    (
        [Parameter()]
        [String]
        $State
    )

    # Get all nodes from vCenter
    try
    {
        Write-Host "Getting all VMhosts from vCenter(s) ... " -NoNewline
        $OutVariable = Get-VMhost | Where-Object {$_.ConnectionState -eq $State}
        Write-Host "OK"

        return $OutVariable
    }
    catch
    {
        Write-Host "FAIL!" -BackgroundColor Red
        Write-Host "Could not get VMhosts from vCenter"
        throw
    }
}

<#
   _____      _     ______  _______   ___   _      _                            __                              _____           _
  / ____|    | |   |  ____|/ ____\ \ / (_) | |    (_)                          / _|                            / ____|         | |
 | |  __  ___| |_  | |__  | (___  \ V / _  | |     _  ___ ___ _ __  ___  ___  | |_ _ __ ___  _ __ ___   __   _| |     ___ _ __ | |_ ___ _ __
 | | |_ |/ _ \ __| |  __|  \___ \  > < | | | |    | |/ __/ _ \ '_ \/ __|/ _ \ |  _| '__/ _ \| '_ ` _ \  \ \ / / |    / _ \ '_ \| __/ _ \ '__|
 | |__| |  __/ |_  | |____ ____) |/ . \| | | |____| | (_|  __/ | | \__ \  __/ | | | | | (_) | | | | | |  \ V /| |___|  __/ | | | ||  __/ |
  \_____|\___|\__| |______|_____//_/ \_\_| |______|_|\___\___|_| |_|___/\___| |_| |_|  \___/|_| |_| |_|   \_/  \_____\___|_| |_|\__\___|_|


#>
Function Get-ESXiLicenseFromvCenter
{
    Param
    (
        [Parameter()]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMHost
    )

    Write-Debug 'Stepping into Get-ESXiLicenseFromvCenter function'
    
    # Extract all VMhosts from same vCenter as provided host, 
    # select the first host that is not the selected one,
    # Retrieve that hosts licensekey
    $LicenseKey = Get-VMHost -Server (Get-vCenterFromVMhost -VMhost $VMhost) |
        Where-Object { $_.Name -ne $VMhost.Name } |
        Select-Object -First 1 |
        Select-Object -Property LicenseKey

    return $LicenseKey.LicenseKey
}

function Get-PodFromCluster
{
    Param
    (
        [Parameter()]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.ClusterImpl]
        $Cluster
    )

    # Attempt quick method
    $Pod = ($Cluster.CustomFields)["Pod"]

    Return ($Pod).ToUpper()
}
<#
   _____      _                 _____           _              ______                          _           _
  / ____|    | |               / ____|         | |            |  ____|                        | |         | |
 | |  __  ___| |_ ________   _| |     ___ _ __ | |_ ___ _ __  | |__ _ __ ___  _ __ ___     ___| |_   _ ___| |_ ___ _ __
 | | |_ |/ _ \ __|______\ \ / / |    / _ \ '_ \| __/ _ \ '__| |  __| '__/ _ \| '_ ` _ \   / __| | | | / __| __/ _ \ '__|
 | |__| |  __/ |_        \ V /| |___|  __/ | | | ||  __/ |    | |  | | | (_) | | | | | | | (__| | |_| \__ \ ||  __/ |
  \_____|\___|\__|        \_/  \_____\___|_| |_|\__\___|_|    |_|  |_|  \___/|_| |_| |_|  \___|_|\__,_|___/\__\___|_|


#>
Function Get-vCenterFromCluster
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.ClusterImpl]
        $VMHostCluster
    )

    $vCenterSeletion = $VMHostCluster.Uid.Split('@').Split(':')[1]
    return $vCenterSeletion
}

<#
   _____      _                 _____           _            ______               __      ____  __ _               _   
  / ____|    | |               / ____|         | |          |  ____|              \ \    / /  \/  | |             | |  
 | |  __  ___| |_ ________   _| |     ___ _ __ | |_ ___ _ __| |__ _ __ ___  _ __ __\ \  / /| \  / | |__   ___  ___| |_ 
 | | |_ |/ _ \ __|______\ \ / / |    / _ \ '_ \| __/ _ \ '__|  __| '__/ _ \| '_ ` _ \ \/ / | |\/| | '_ \ / _ \/ __| __|
 | |__| |  __/ |_        \ V /| |___|  __/ | | | ||  __/ |  | |  | | | (_) | | | | | \  /  | |  | | | | | (_) \__ \ |_ 
  \_____|\___|\__|        \_/  \_____\___|_| |_|\__\___|_|  |_|  |_|  \___/|_| |_| |_|\/   |_|  |_|_| |_|\___/|___/\__|
                                                                                                                       
#>
Function Get-vCenterFromVMhost
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost
    )

    $vCenterSelection = $VMhost.Uid.Split('@').Split(':')[1]
    
    Return $vCenterSelection
}
Function Set-VMHostPowerStateOff
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost,

        [Parameter()]
        [pscredential]
        $OrionCredential
    )

    # Continue work on expected state
    if ($VMhost.ConnectionState -eq 'Maintenance')
    {

        # Pull VMhostAttributes into common variable
        $VMhostAttributes = Get-VMhostCustomAttributes -VMhost $VMhost
        Write-Host "Starting to power off $($VMhostAttributes.Hostname)"

        # Start by unmanaging the host
        try
        {
            $UnmanageParams = @{
                NodeIP          = $VMhostAttributes.ManagementIP
                OrionCredential = $OrionCredential
                State           = "Unmanaged"
            }
            Set-OrionNodeState @UnmanageParams
        }
        catch
        {
            Write-Host "Could not unmanage host. Exiting"
            throw
        }

        # THIS IS OUTCOMMENTED (25-10-2023), BECAUSE INTERSIGHT.POWERSHELL MODULE HAS ISSUES WITH CERTIFICATES
        # AND CANNOT LOG INTO INTERSGIHT AND CENTRAL AT THE SAME TIME
        # THIS FUNCTIONALITY IS MOVED OUT OF FUNCTION, HANDLED BY PARENT SCRIPT AND HANDLED IN OTHER RUNSPACE

        # ---- 1: First set the label on the blade itself
        #if ($VMhost.Model -notmatch 'm6') #Legacy UCS
        #{
        #    Write-Host "$($VMhostAttributes.Hostname) is a UCS Central Blade"
        #    Set-CentralBladeLabel -UsrLbl 'PoweredOffDueToSchedule' -FQDN $VMhost.Name
#
        #}
        #elseif ($VMhost.model -match 'm6') #Intersight
        #{
        #    Write-Host "$($VMhostAttributes.Hostname) is a UCS Intersight Blade"
        #    Set-intersightBladeLabel -UsrLbl 'PoweredOffDueToSchedule' -FQDN $VMhost.Name
        #}


        # ---- 2: Then set tags in vCenter
        try
        {
            Write-Host "Setting tags on VMhost in vCenter ... " -NoNewline
            $VMhost.ExtensionData.setCustomValue('State', 'PoweredOffDueToSchedule')
            $VMhost.ExtensionData.setCustomValue('StateDate', $(Get-FormattedTime))
            Write-Host "OK"
        }
        catch
        {
            Write-Host "FAIL!"
            Write-Host "Could not set tags in vCenter"
            throw
        }


        # ---- 3: Finally Power Off the host from vCenter
        try
        {
            Write-Host "Sending Shutdown command ... " -NoNewline
            Start-Sleep -seconds 10
            [void]::(Stop-VMHost -VMHost $VMhostAttributes.Name -Reason "Scheduled Power off" -Confirm:$false)
            Write-Host "OK"
        }
        catch
        {
            Write-Host "FAIL!" -BackgroundColor Red
            throw
        }
    }
    else
    {
        Write-Host "Node is NOT in maintenance mode. EXIT" -BackgroundColor Red
        throw
    }
}

<#

 __      ____  __                           _____             __ _                        _    _           _
 \ \    / /  \/  |                         / ____|           / _(_)                      | |  | |         | |
  \ \  / /| \  / |_      ____ _ _ __ ___  | |     ___  _ __ | |_ _  __ _ _   _ _ __ ___  | |__| | ___  ___| |_ ___
   \ \/ / | |\/| \ \ /\ / / _` | '__/ _ \ | |    / _ \| '_ \|  _| |/ _` | | | | '__/ _ \ |  __  |/ _ \/ __| __/ __|
    \  /  | |  | |\ V  V / (_| | | |  __/ | |___| (_) | | | | | | | (_| | |_| | | |  __/ | |  | | (_) \__ \ |_\__ \
     \/   |_|  |_| \_/\_/ \__,_|_|  \___|  \_____\___/|_| |_|_| |_|\__, |\__,_|_|  \___| |_|  |_|\___/|___/\__|___/
                                                                    __/ |
                                                                   |___/

#>
Function Enable-VMhost
{
    #------------------------------------------------| HELP |------------------------------------------------#
    <#
        .Synopsis
            This script will configure a VMhost, that is about to go into production. It will config the host in VMware, position it in the correct cluster, set its password in PAM
        .PARAMETER vCenter
            vCenter to put the Host into
        .PARAMETER vCenterCredential
            Credentials for logging into vCenter
        .PARAMETER TargetCluster
            Destination cluster of the VMhost
        .PARAMETER VMHost
            Name of the VMhost to work on
        .PARAMETER SafeName
            Name of safe to look in in PAM
        .PARAMETER PAMURI
            URI of PAM instance
        .PARAMETER PAMCredential
            Credentials for logging into PAM
    #>

    Param
    (
        [Parameter()]
        [String]
        $vCenterServer,

        [Parameter()]
        [pscredential]
        $vCenterCredential = (Get-Credential),

        [Parameter()]
        [pscredential]
        $DefaultVMhostCredential,

        [Parameter()]
        [String]
        $TargetCluster,

        [Parameter()]
        [String]
        $TargetVMHost,

        [Parameter()]
        [String]
        $SafeName = "ncop-vmw-prod_emergency",

        [Parameter()]
        [String]
        $PAMURI = "https://pam.nchosting.dk",

        [Parameter()]
        [pscredential]
        $PAMCredential = (Get-Credential),

        [Parameter()]
        [pscredential]
        $OrionCredential

    )

    #------------------------------------------------| SETUP |-----------------------------------------------#

    Write-HostSeperator "Setup" -Width 40
    Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential
    #
    $Cluster   = Get-Cluster -name $TargetCluster
    $VIServer  = Get-vCenterFromCluster -VMHostCluster $Cluster

    #$VIServer = $ClusterSelection.Uid.Split("@").Split(":")[1]
    Write-Host "$VIServer selected"

    # Change to jenkins parameter
    $SNMPCommunity  = "Fors23"
    $SSHPolicy      = "On"
    #$VMHostUser     = "root"
    #$VMHostPassword = "myp@ssw0rd"
    #$LogNFSHost     = "100.64.24.210"
    #$LogNFSPath     = "/VMware_logs"
    #$LogNFSName     = "log"
    #$LogSettingPath = "[VMware_logs] ESXiLogs"

    $NTPAddresses = @(
        [IPAddress]"5.44.143.71",
        [IPAddress]"5.44.143.72",
        [IPAddress]"5.44.143.73",
        [IPAddress]"5.44.143.74"
    )

    $Arguments = @{
        "Length"            = 21
        "ForbiddenChars"    = @('I','l','O','o','0','1','*','<','>','@','[','\',']','{','|','}','~','$',':',';')
        "MinLowercaseChars" = 3
        "MinUppercaseChars" = 3
        "MinDigits"         = 3
        "MinSpecialChars"   = 3
        "AsString"          = $true
        }

    $NewPassword = New-Password @Arguments


    #----------------------------------------| VALIDATE TARGET HOST |----------------------------------------#


    # Check if it's possible to login on Target VMhost
    $TestParams = @{
        TargetVMhost      = $TargetVMHost
        VMhostCredential  = $DefaultVMhostCredential
        SetMaintenance    = $True
    }
    Test-VMhostLogin @TestParams

    # Check if the VMhost is in vCenter, has VMs or templates
    Write-Host "Validating VMhost ... " -NoNewline
    $VMHostCheck   = Get-VMHost $TargetVMHost    -ErrorAction SilentlyContinue

    if ($null -eq $VMHostCheck)
    {
        Write-Host "OK"
    }
    else
    {
        Write-Host "FAIL!"
        Write-Host "Checking VMhost compliancy"
        Search-VMhost -VMhost $VMHostCheck
    }


    #---------------------------| PUT NEW VMHOST INTO vCENTER AND ASSIGN LICENSE |---------------------------#

    $CreationFolder = Get-Folder -server $VIServer | Where-Object {$_.Name -eq "Creation" -and $_.Type -eq "HostAndCluster"}
    try
    {
        Write-Host "Adding $TargetVMHost to $VIServer ... " -NoNewline
        $AddParams = @{
            Name       = $TargetVMHost
            Location   = $(get-folder -Id $CreationFolder.Id)
            Credential = $DefaultVMhostCredential
            Force      = $true
            Server     = $VIServer
        }
        [void]::(Add-VMHost @AddParams)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }

    # Put VMhsot in variable
    $VMhost = Get-VMHost $TargetVMHost
    Write-Host "Gathering correct license from $VIServer"

    $LicenseKey = Get-VMHost -Server $VMhost.Uid.Split("@").Split(":")[1] | Where-Object {$_.Name -ne $VMHost} | Select-Object -First 1  | Select-Object -Property LicenseKey
    Write-Host $LicenseKey.LicenseKey "Selected"

    try
    {
        Write-Host "Assigning Licensekey to VMhost ... " -NoNewline
        [void]::(Set-VMHost -VMHost $VMHost -LicenseKey $LicenseKey.LicenseKey)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }


    #---------------------------------------| DATASTORE CONFIGURATION |--------------------------------------#

    try
    {
        Write-Host "Configuring Datastores ... " -NoNewline
        [void]::(Get-VMHost $VMHost | Get-ScsiLun -LunType disk | Where-Object { $_.MultipathPolicy -notlike "RoundRobin" } | Set-Scsilun -MultiPathPolicy "RoundRobin")
        Write-Host "OK"
        Write-Host "All datastores set to multipath policy RoundRobin`n"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }


    #----------------------------------------| SERVICE CONFIGURATION |---------------------------------------#

    # NTP
    Write-HostSeperator "Setting up NTP" -Width 40
    $NTPCheck = Get-VMHostNtpServer -VMHost $VMHost

    try
    {
        if ($NTPCheck -notcontains ($NTPAddresses.IPAddressToString)[0])
        {
            Set-VMhostNTPservers -VMhost $VMhost -NTPServer $NTPAddresses
        }
        else
        {
            Write-Host "NTP is already set up with information:"
            $NTPCheck
        }
    }
    finally
    {
        Write-Host "Making sure NTP service is turned ON ... " -NoNewline
        [void]::(Get-VmHostService -VMHost $VMHost | Where-Object { $_.key -eq "ntpd" } | Start-VMHostService)
        [void]::(Get-VmHostService -VMHost $VMHost | Where-Object { $_.key -eq "ntpd" } | Set-VMHostService -policy "automatic")
        Write-Host "OK"
    }

    # SSH
    Write-HostSeperator "SSH" -Width 40
    try
    {
        Write-Host "Setting up SSH to `"$SSHPolicy`" ... " -NoNewline
        [void]::(Get-VmHostService -VMHost $VMHost | Where-Object { $_.Key -eq "TSM-SSH" } | Start-VMHostService)
        [void]::(Get-VmHostService -VMHost $VMHost | Where-Object { $_.Key -eq "TSM-SSH" } | Set-VMHostService -Policy $SSHPolicy)
        Write-host "OK"
        Write-Host ""
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }

    # SNMP
    Write-HostSeperator "SNMP" -Width 40
    try
    {
        Write-Host "Connecting directly to ESXi host $TargetVMhost ... " -NoNewline
        [void]::(Connect-VIServer $VMHost -Credential $DefaultVMhostCredential)
        Write-Host "OK"

        Write-Host "Setting SNMP on VMhost ... " -NoNewline
        $VMHostSNMP = Get-VMHostSnmp -Server $TargetVMHost
        $VMHostSNMP = Set-VMHostSnmp $VMHostSNMP -Enabled:$true -ReadOnlyCommunity $SNMPCommunity
        Write-Host "OK"

        Write-Host "Disconnecting from ESXi host $VMHost ... " -NoNewline
        [void]::(Disconnect-VIServer -Server $TargetVMHost -Confirm:$false)
        Write-Host "OK"

        Write-Host "Starting SNMP Services ... " -NoNewline
        [void]::(Get-VmHostService -VMHost $VMHost | Where-Object { $_.Key -eq "snmpd" } | Start-VMHostService)
        [void]::(Get-VmHostService -VMHost $VMHost | Where-Object { $_.Key -eq "snmpd" } | Set-VMHostService -Policy "On")
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }


    #------------------------------------------| NIC CONFIGURATION |-----------------------------------------#
    Write-HostSeperator "Nic Configuration" -Width 40

    $Pod = Get-PodFromCluster -Cluster $Cluster

    if ($Pod -eq "POD01")
    {
        $ManagementPortGroup = "vmware-esxi-mgmt-pod01"
    }
    elseif ($Pod -eq "POD02")
    {
        $ManagementPortGroup = "vmware-esxi-mgmt-pod02"
    }

    try
    {
        $VMNics = Get-VMHost $VMHost | Get-VMHostNetworkAdapter | Where-Object { $_.Name -notlike "vmk*" }
        $VMks   = Get-VMHost $VMHost | Get-VMHostNetworkAdapter | Where-Object { $_.Name -like "*vmk*" }
    }
    catch
    {
        Write-Host "FAIL"
        throw
    }

    #----------------------------------------| HOSTING / PRODUCTION |----------------------------------------#
    try
    {
        Write-Host "Setting up hosting NICs ..." -NoNewline
        do {
            if (($VIServer -like "*vCenter01*") -or ($VIServer -like "*vCenter03*"))
            {
                [void]::(Get-VDSwitch -Name "Hosting" -Server $VIServer | Add-VDSwitchVMHost -VMHost $VMHost -Confirm:$false -ErrorAction SilentlyContinue)
                $VDSwitch = "Hosting"
            }
            elseif ($VIServer -like "*vCenter02*")
            {
                [void]::(Get-VDSwitch -Name "Produktion" -Server $VIServer  | Add-VDSwitchVMHost -VMHost $VMHost -Confirm:$false -ErrorAction SilentlyContinue)
                $VDswitch = "Produktion"
            }

            $VMHostView = Get-View -ViewType HostSystem -filter @{"Name" = $TargetVMhost}
            [void]::(Get-VMHost $VMHost | Get-VDSwitch -Name $VDSwitch | Add-VDSwitchPhysicalNetworkAdapter $VMNics[3] -Confirm:$false -ErrorAction SilentlyContinue)
            [void]::(Get-VMHost $VMHost | Get-VDSwitch -Name $VDSwitch | Add-VDSwitchPhysicalNetworkAdapter $VMNics[4] -Confirm:$false -ErrorAction SilentlyContinue)

            $VMNicCheck1  = $VMNICs[3].Name
            $VMNicCheck2  = $VMNICs[4].Name
            $NICs         = $VMHostView.Config.Network.ProxySwitch
            $HostingNIC   = $NICs | Where-Object { $_.DvsName -eq $VDSwitch }
            $HostingCheck = ($HostingNIC.Pnic -like "*$VMNicCheck1") -and ($HostingNIC.Pnic -like "*$VMNicCheck2")

            Write-Host "." -NoNewline
            Start-Sleep -Seconds 1
        }
        until ($HostingCheck -eq "true")
        Write-Host " OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }


    #---------------------------------------------| MANAGEMENT |---------------------------------------------#
    Write-Host "Setting up Management NICs ..." -NoNewline
    do {
        $VMNicCheck3   = $VMNICs[1].Name
        $VMNicCheck4   = $VMNICs[2].Name
        $VMHostView    = Get-View -ViewType HostSystem -filter @{"Name" = $TargetVMhost}
        $NICs          = $VMHostView.Config.Network.ProxySwitch
        $ManagementNIC = $NICs | Where-Object { $_.DvsName -like "Management*" }

        #Write-Host "DvSwitch $($ManagementNIC.DvsName) selected in $VIServer"
        $ManagementCheck = ($ManagementNIC.Pnic -like "*$VMNicCheck3") -and ($ManagementNIC.Pnic -like "*$VMNicCheck4")
        $ManagementDVSwitch = Get-VDSwitch | Where-Object Name -match "Management"

        if (-not $ManagementCheck)
        {
            [void]::($ManagementDVSwitch = Get-VDPortgroup -name $ManagementPortGroup -VDSwitch $ManagementDVSwitch -Server $VIServer)
            [void]::(Get-VDSwitch -Name "Management*" -Server $VIServer | Add-VDSwitchVMHost -VMHost $VMHost -Confirm:$false -ErrorAction SilentlyContinue)
            [void]::(Get-VMHost $VMHost | Get-VDSwitch -Name "Management*" -Server $VIServer | Add-VDSwitchPhysicalNetworkAdapter $VMNics[2] -Confirm:$false -ErrorAction Ignore)
            [void]::(Set-VMHostNetworkAdapter -PortGroup $ManagementDVSwitch -VirtualNic $VMKs -Confirm:$false)

            Start-Sleep -Seconds 10
            [void]::(Get-VMHost $VMHost | Get-VDSwitch -Name "Management*" -Server $VIServer | Add-VDSwitchPhysicalNetworkAdapter $VMNics[1] -Confirm:$false -ErrorAction Ignore)
            [void]::(Get-VMHost $VMHost | Get-VirtualSwitch -Server $VIServer | Where-Object { $_.Name -like "*vSwitch*" } | Remove-VirtualSwitch -Confirm:$false)
            Write-Host "." -NoNewline
        }
    }
    until ($ManagementCheck -eq "true")
    Write-Host " OK"

    #-----------------------------------------------| vMOTION |----------------------------------------------#
    Write-Host "Setting up vMotion NICs .." -NoNewline
    try
    {
        do {
            $VMNicCheck5  = $VMNICs[0].Name
            $NICs         = $VMHostView.Config.Network.ProxySwitch
            $vMotionNIC   = $NICs | Where-Object { $_.DvsName -like "vMotion*" }
            $vMotionCheck = $vMotionNIC.Pnic -like "*$VMNicCheck5"

            if ($null -ne $vMotionCheck)
            {
                $vMotionCheck = "true"
            }
            else
            {
                $vMotionCheck = ""
            }

            $VMHostView      = Get-View -ViewType HostSystem -filter @{"Name" = $TargetVMhost}
            $vMotionVDSwitch = Get-VDSwitch      -Server $VIServer | Where-Object Name -match "vMotion"
            $VirtualSwitch   = Get-VirtualSwitch -Server $VIServer | Where-Object Name -match "vMotion"

            [void]::($vMotionVDSwitch | Add-VDSwitchVMHost -VMHost $VMHost -Confirm:$false -ErrorAction SilentlyContinue)
            [void]::(Get-VMHost $VMHost -Server $VIServer | Get-VDSwitch -Name "vMotion*" | Add-VDSwitchPhysicalNetworkAdapter $VMNics[0] -Confirm:$false -ErrorAction SilentlyContinue)

            if ($VIServer -like "*vCenter01*" -or $VIServer -like "*vCenter03*")
            {
                [void]::(New-VMHostNetworkAdapter -VMHost $VMHost -VirtualSwitch $vMotionVDSwitch -PortGroup "vmware-esxi-vmotion" -VMotionEnabled 1 -Server $Viserver)
            }
            else
            {
                [void]::(New-VMHostNetworkAdapter -VMHost $VMHost -VirtualSwitch $VirtualSwitch -PortGroup "vmware-esxi-vmotion" -VMotionEnabled 1)
            }
            Write-Host "." -NoNewline

        }
        until ($vMotioNCheck -eq "true")
        Write-Host " OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }


    #---------------------------------------| ADVANCED CONFIGURATION |---------------------------------------#
    Write-HostSeperator "Advanced Configuration" -Width 40
    try
    {
        Write-Host "Supressing shell log warnings ... " -NoNewline
        [void]::(Get-VMHost $VMHost | Get-AdvancedSetting -Name "UserVars.SuppressShellWarning" | Set-AdvancedSetting -Value 1 -Confirm:$false)
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }

    #----------------------------------------| ADD TO TARGET CLUSTER |---------------------------------------#
    if ($TargetCluster)
    {
        Write-Host "Moving $TargetVMhots to $TargetCluster ... " -NoNewline
        [void]::(Move-VMHost $VMHost -Destination $TargetCluster)
        Write-Host "OK"
    }


    #-----------------------------------| UPDATE ROOT PASSWORD ON VMHOST |-----------------------------------#
    try
    {
        Write-Host "Resetting ESXi password ..." -NoNewline
        [void]::(Disconnect-VIServer * -Confirm:$false -ErrorAction "SilentlyContinue")
        [void]::(Connect-VIServer -Server $TargetVMHost -Credential $DefaultVMhostCredential)
        [void]::(Set-VMHostAccount -UserAccount "root" -Password $NewPassword -Confirm:$false)
        [void]::(Disconnect-VIServer -Server $TargetVMHost -Confirm:$false)
        Write-Host "OK"

        Write-Host "$TargetVMhost Ready for use" -BackgroundColor Green
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }


    #----------------------------------------| PAM PASSWORD ADDITION |---------------------------------------#
    Write-HostSeperator "PAM" -Width 40
    Connect-NcPAM -PamCredential $PAMCredential

    # Only create new PAS account, if it does not already exist
    if ($Null -eq (Get-PASAccount -Search $VMHost))
    {
        try
        {
            # Set arguments for adding account
            $Arguments = @{
                "address"                    = $VMHost
                "UserName"                   = "Root"
                "Secret"                     = $NewPassword
                "SafeName"                   = $SafeName
                "secretType"                 = "Password"
                "platformID"                 = "NCESXirootAccount"
                "automaticManagementEnabled" = $False
                "manualManagementReason"     = "Rotated by Compute team"
            }

            # Add the account
            Write-Host "Adding $VMHost to PAM ... " -NoNewline
            Add-PASAccount @Arguments
            Write-Host "OK"
        }
        catch
        {
            Write-Host "FAIL"
        }
        finally
        {
            Close-PASSession
        }
    }
    else
    {
        Write-Host "Account already exists in PAM"
    }

    #------------------------------------------------| ORION |-----------------------------------------------#
    Write-HostSeperator "Orion" -Width 40

    #TODO:
    #Get-IPfromReservation -ReservationName $TargetVMhost

    #$AddNodeParams = @{
    #    SwisConnection = (Connect-Orion -OrionCredential $OrionCredential)
    #    NodeName       = $TargetVMHost
    #    $NodeIPAddress = ""
    #}
    #TODO Implement this feature
    #Add-OrionNode @AddNodeParams
    #-------------------------------------------------| END |------------------------------------------------#
}
function Search-VMhost
{
    param
    (
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost
    )

    # query local VMs and templates
    $VMCheck       = $VMHostCheck | Get-VM       -ErrorAction SilentlyContinue
    $TemplateCheck = $VMHostCheck | Get-Template -ErrorAction SilentlyContinue

    # Complete checks
    if ($VMhost.ConnectionState -eq "connected")
    {
        Write-Host $VMhost.Name "Connected, stopping script" -BackgroundColor Red
        throw
    }
    else
    {
        Write-Host $VMhost "Not connected, proceeding" -BackgroundColor Green
    }

    if ($VMhost.PowerState -eq "PoweredOn")
    {
        Write-Host $VMhost.Name "Powered on, stopping script" -BackgroundColor Red
        throw
    }
    else
    {
        Write-Host $VMhost.Name "Not powered on, proceeding" -BackgroundColor Green
    }

    if ($null -ne $VMCheck)
    {
        Write-Host "No VMs located on " $VMhost.Name " proceeding" -BackgroundColor Green
    }
    else
    {
        Write-Host "VMs located on " $VMhost.Name " stopping script" -BackgroundColor Red
        throw
    }

    if ($null -ne $TemplateCheck)
    {
        Write-Host "No Templates located on " $VMhost.Name " proceeding" -BackgroundColor Green
    }
    else
    {
        Write-Host "Templates located on " $VMhost.Name "stopping script" -BackgroundColor Red
        throw
    }

    Write-Host ""
    Write-Host "Removing "$VMhost.Name " from vCenter" -BackgroundColor Green
    Write-Host ""
    $VMhost | Remove-VMHost -Confirm:$false
}
function Set-VMhostNTPservers
{
    param
    (
        [Parameter()]
        [IPAddress[]]
        $NTPServer,

        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost
    )

    # Loop over each Address in input, assign them
    Foreach ($Address in $NTPServer)
    {
        try
        {
            Write-Host "Adding NTP Server $Address to $VMhost ... " -NoNewline

            [void](Add-VmHostNtpServer -VMHost $VMHost -NtpServer $Address)
            Write-Host "OK"
        }
        catch
        {
            Write-Host "FAIL!"
            throw
        }
    }

    # And make sure the NTP service is running and will start at reboot
    try
    {
        Write-Host "Starting NTP service ... "  -NoNewline
        [void](Get-VmHostService -VMHost $VMHost | Where-Object { $_.key -eq "ntpd" } | Start-VMHostService)
        [void](Get-VmHostService -VMHost $VMHost | Where-Object { $_.key -eq "ntpd" } | Set-VMHostService -policy "automatic")
        Write-Host "OK"
    }
    catch
    {
        Write-Host "FAIL!"
        throw
    }
}
function Test-VMhostLogin
{
    param
    (
        [Parameter()]
        [pscredential]
        $VMhostCredential,

        [Parameter()]
        [String]
        $TargetVMHost,

        [Parameter()]
        [Boolean]
        $SetMaintenance
    )

    # Prepare params for loop
    $ConnectionParams = @{
        Server      = $TargetVMHost
        Credential  = $VMhostCredential
        ErrorAction = "SilentlyContinue"
    }

    $VMhostName = ($TargetVMHost -split "\.")[0]
    $Counter    = 0

    Write-Host "Testing connectivity to $VMhostName ." -NoNewline

    # Repeat until connection is made
    # This is done, becuase it takes up to two hours for the DNS entries to allow this call to go through
    # TODO: Refresh DNS on demand
    do {

        # Do something to the console
        Write-Host "." -NoNewline

        # Check connection
        $ESXiUpTest = Connect-VIServer @ConnectionParams

        # if no connection, just try again
        if($null -eq $ESXiUpTest)
        {
            Write-Host "."
            Start-Sleep -Seconds 30
            $Counter++
        }

        # three hours have passed
        if ($Counter -eq 360)
        {
            Write-Host " FAIL!"
            Write-Host "Could not verify connectivity. Check DNS"
            throw
        }

    }
    Until($null -ne $ESXiUpTest)

    Write-Host " OK"
    Write-Host "Connection to $VMhostName verified"

    if ($SetMaintenance)
    {
        Write-host "Setting $VMhostName in Maintenance mode ... " -NoNewline
        $VMhost = Get-VMhost -server $ESXiUpTest
        [void]::(Set-VMHost -VMHost $VMhost -State "Maintenance")
        $VMhost = $null
        Write-Host "OK"
    }

    # Disconnect, to not have lingering connections
    [void]::(Disconnect-VIServer -Server $TargetVMHost -Confirm:$false)
}
function Get-VMhostCustomAttributes {

    param
    (
       [Parameter(Mandatory = $true)]
       [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
       $VMhost
    )

    # Create custom object for iteration
    $VMhostObject = @{
        Name                = ($VMhost.Name)
        Hostname            = ($VMhost.Name -split "\.")[0]
        ConnectionState     = ($VMhost.ConnectionState)
        ManagementIP        = ($vmhost | Get-VMHostNetworkAdapter -ErrorAction SilentlyContinue | Where-Object { $_.ManagementTrafficEnabled }).IP
        State               = ($VMhost.Customfields)["State"]
        StateDate           = ($VMhost.Customfields)["StateDate"]
        ESXiVersion         = ($VMhost.Customfields)["ESXiVersion"]
        FirmwareVersion     = ($VMhost.Customfields)["FirmwareVersion"]
        System              = ($VMhost.Customfields)["System"]
        UpgradeState        = ($VMhost.Customfields)["UpgradeState"] | Sort-Object 'StateDate' -Descending
        StateMachineIgnore  = ($VMhost.Customfields)["StateMachineIgnore"]
        MaintenanceTaskDate	= ($VMhost.Customfields)["MaintenanceTaskDate"]
        MaintenanceTaskId   = ($VMhost.Customfields)["MaintenanceTaskId"]
    }

    # And return the object
    return $VMhostobject
}
function Resolve-StateMachineOutlierStates
{
    #Set parameters for the script here
    param
    (
        [Parameter()]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost,

        [Parameter()]
        [String]
        $VMhostName
    )
    #endregion

    # VMhost is offline, but state is live
    if ($VMhost.ConnectionState -eq "NotResponding" -and $VMhostAttributes.State -eq "live")
    {
        Write-error "Host is not responding, but in state live!"
        #Send-NotResponding -VMhost $VMhost
    }


    # VM is for whatever reason missing custom attributes; enrich the VMhost
    if ("" -eq $VMhostAttributes.State)
    {
        # Refresh all variables to current values
        Set-VMhostCustomAttributes -VMhostName $VMhost.name
    }
}
<#
  _    _  _____  _____   _____                        ____          _____
 | |  | |/ ____|/ ____| |  __ \                      / __ \        |  __ \
 | |  | | |    | (___   | |__) |____      _____ _ __| |  | |_ __   | |__) | __ ___   __ _ _ __ ___  ___ ___
 | |  | | |     \___ \  |  ___/ _ \ \ /\ / / _ \ '__| |  | | '_ \  |  ___/ '__/ _ \ / _` | '__/ _ \/ __/ __|
 | |__| | |____ ____) | | |  | (_) \ V  V /  __/ |  | |__| | | | | | |   | | | (_) | (_| | | |  __/\__ \__ \
  \____/ \_____|_____/  |_|   \___/ \_/\_/ \___|_|   \____/|_| |_| |_|   |_|  \___/ \__, |_|  \___||___/___/
                                                                                     __/ |
                                                                                    |___/
#>
Function Resolve-UCSPowerOnProgress
{
    param (
        [Parameter()]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost,

        [Parameter()]
        [String]
        $VMhostName
    )

    #region--------------------------------------| PROGRAM LOGIC |-------------------------------------------#

    # Builds the custom varialbes as seen below for the statemachine.
    Set-StateMachineVariables
    <#
        $PoweringOnLabel
        $PowerOffLabel
        $ReadyForScheduledPowerDownLabel
        $LiveLabel
        $EnteringMaintenancePowerDownLabel
        $MarkedForUpgradeLabel
        $SelectedForUpgradeLabel
        $EnteringMaintenanceUpgradelabel
    #>

    # Used when dealing with deserialized objects in other threads
    if ($VMhostName)
    {
        $VMhost = get-VMhost -name $VMhostName
    }

    # Set variables for script
    $VMhostAttributes = Get-VMhostCustomAttributes -VMhost $VMhost
    $CurrentDate      = Get-FormattedTime
    $TimeSpan         = [Math]::Abs($(New-TimeSpan -Start $VMhostAttributes.StateDate -End $CurrentDate).TotalMinutes)

    Write-Host "Time is: $CurrentDate. $($VMhostAttributes.Hostname) has current State Date: $($VMhostAttributes.StateDate)   -   TimeSpan is $TimeSpan"

    # If blade is a legacy UCS blade, power via UCS Central
    if ($VMhost.Model -notmatch 'm6')
    {
        if ($TimeSpan -ge $AwaitTimer)
        {
            Write-Host "Waited for $($VMhostAttributes.Hostname) to PowerON for $TimeSpan minutes"
            $Blade = Get-UCSCentralServiceProfile -Descr $VMhost.Name

            # Throw error if no blade is returned
            if ($null -eq $Blade)
            {
                Write-Host "No blade found!" -BackgroundColor Red
                throw
            }

            Write-Host 'Checking if blade is in the correct state'
            if ($Blade.Usrlbl -in $PoweringOnLabel,$PowerOffLabel -and $VMhost.ConnectionState -notin "Maintenance","Connected")
            {
                Write-Host 'blade validated, checking for major errors'
                Write-Host "$($Blade.descr) took too long to power on, checking hosts for critical faults."
                $BladeFaults = Get-UcsCentralFaultDomainInst -OrigSeverity "major" | Where-Object {($_.CentralAffectedObject -match $Blade.PnDn -and $_.cause -match 'vif-down') -or ($_.cause -match 'link-down')}

                if ($null -ne $BladeFaults)
                {
                    Write-Host "Multiple uplink errors detected, power cycling $($Blade).descr"
                    [void]($Blade | Get-UcsCentralLsServerOperation | Set-UcsCentralLsServerOperation -State "cycle-immediate" -Confirm:$false -Force -ErrorAction SilentlyContinue)

                    # Update VMhostAttributes.StateDate on VMhost
                    Write-Host 'Setting states on ESXi hosts'
                    $VMhost.ExtensionData.setCustomValue('StateDate', $(Get-FormattedTime))
                }
            }
            elseif ($VMhost.ConnectionState -eq "Maintenance")
            {
                # Update VMhostAttributes.StateDate on VMhost
                Write-Host "Blade is marked $($blade.UsrLbl), but VMhost is maintenance. Update tags"

                Set-CentralBladeLabel -FQDN $VMhost.Name -UsrLbl "live"

                Write-Host "Updating VMhost tags ... " -NoNewline
                $VMhost.ExtensionData.setCustomValue('StateDate', $(Get-FormattedTime))
                Write-Host "OK"
            }
        }
        else
        {
            Write-Host "Waiting for $($VMhostAttributes.Hostname) to power on ... giving it some more time"
        }
    }

    #TODO: Implement Intersight poweronhandler
    # NOTE JVM: Intersight blades seems to power on just fine. Don't spend time on this right now
    if ($VMhost.Model -match 'm6')
    {
        Write-Host "Function not yet impelmented! Contact JVM and call him lazy"
    }
}
# Note JVM: Yes, correct, using a function that sets variables will hurt both discoverability and readability of the code in which it participates,
# because none of the variables can be jumped to and read in cleartext anymore. On the plusside, it is absolutely *essential* that the state machine variables
# are 100% case sensitive spelled correct for the State machine to work. All in all, I think it's a worthwhile, although expensive transaction.
Function Set-StateMachineVariables
{
    param
    (
        [Parameter()]
        [System.Boolean]
        $Quiet
    )

    if (-not $Quiet){Write-Host "Setting statemachine Variables ... " -NoNewline}

    $DefaultOptions = @{
        Option = "AllScope"
        Scope  = "Global"
        Force  = $true
    }

    try
    {
        # Strings
        New-Variable @DefaultOptions -name 'PoweringOnLabel'                   -Value 'PoweringON'
        New-Variable @DefaultOptions -name 'PowerOffLabel'                     -Value 'PoweredOffDueToSchedule'
        New-Variable @DefaultOptions -name 'ReadyForScheduledPowerDownLabel'   -Value 'readyforscheduledpowerdown'
        New-Variable @DefaultOptions -name 'LiveLabel'                         -Value 'live'
        New-Variable @DefaultOptions -name 'EnteringMaintenancePowerDownLabel' -Value 'entering-maintenanace-scheduled-power-down'
        New-Variable @DefaultOptions -name 'MarkedForUpgradeLabel'             -Value 'MarkedForAutoUpgrade'
        New-Variable @DefaultOptions -name 'SelectedForUpgradeLabel'           -Value 'SelectedForUpgrade'
        New-Variable @DefaultOptions -name 'EnteringMaintenanceUpgradelabel'   -Value 'entering-maintenanace-scheduled-upgrade'

        # Ints
        New-Variable @DefaultOptions -name 'AwaitTimer' -Value 15
    }
    catch
    {
        Write-host "FAIL!"
        throw
    }

    if (-not $Quiet){Write-Host "OK"}
}
function Set-VMhostCustomAttributes
{
    #Set parameters for the script here
    param
    (
        [Parameter()]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $VMhost,

        [Parameter()]
        [String]
        $VMhostName
    )
    #endregion

    # Used when dealing with deserialized object in other threads
    if ($VMhostName)
    {
        $VMhost = get-VMhost -name $VMhostName
    }

    # Builds the custom varialbes as seen below for the statemachine.
    Set-StateMachineVariables -Quiet:$true

    # Prepare state variable
    if ($VMhost.ConnectionState -in "connected")
    {
        $state = $global:LiveLabel
    }
    else
    {
        Write-Error "Could not determine current state. Investigate"
        #Send-MissingStateCaseMail -VMhost $VMhost
    }

    # Models
    if ($VMhost.Model -match "m4","m5"){
        $System = "ucscentral"
    }
    else{
        $System = "intersight"
    }

    # Rest of the variables
    $Time               = Get-FormattedTime
    $Firmware           = $vmhost.extensiondata.Hardware.BiosInfo.BiosVersion
    $UpgradeState       = ""
    $StatemachineIgnore = $false

    # Set values and write to host.
    # Note JVM: Don't mess with the tabs. The width is correct.
    Write-Host "Setting custom variables on $VMhost"

    Write-Host " - State`t`t$State"
    $VMhost.ExtensionData.SetCustomValue("State", $State)

    Write-Host " - StateDate`t`t$Time"
    $VMhost.ExtensionData.SetCustomValue("StateDate", $Time)

    Write-Host " - System`t`t$System"
    $VMhost.ExtensionData.SetCustomValue("System", $System)

    Write-Host " - FirmwareVersion`t$Firmware"
    $VMhost.ExtensionData.SetCustomValue("FirmwareVersion", $Firmware)

    Write-Host " - UpgradeState`t$UpgradeState"
    $VMhost.ExtensionData.SetCustomValue("UpgradeState", $UpgradeState)

    Write-Host " - StateMachineIgnore`t$StatemachineIgnore"
    $VMhost.ExtensionData.SetCustomValue("StateMachineIgnore", $StatemachineIgnore)
    Write-Host "Done"
}


# Function cannot be indented correctly due to HTML calls
function Send-MissingStateCaseMail {

param
(
    [parameter()]
    $VMhostName
)

$MailBody = @"

$VMhostName does not contain required custom attributes for Statemachine to work on in.
Please create them manually


<br>
---
<br>

Jenkins jobname: VMware State Machine
@@SITE=NCINFRA@@
@@CASETYPE=Event@@
@@TEAM=Infrastruktur@@
@@APPLICATION=NCOP0013 Servers@@
"@.Replace("`n", '<br>')


$SendMail = @{
    SMTPServer  = "smtp.nchosting.dk"
    From        = "noreply@jenkins.noc01.nchosting.dk"
    To          = "su_GOTOprod_Alarms@netcompany.com"
    Subject     = "Alert Jenkins: VMware | $VMhostname missing custom attributes for Statemachine to work"
    Body        = $Mailbody
}

try
{
    Send-MailMessage @SendMail
}
catch {
    Write-Host "Could not send mail"
    $Error[0]
}

}

function Send-NotResponding
{
    param
    (
        [Parameter()]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
        $Vmhost
    )

    #TODO: Impelment this function
    Write-Host "Function not yet implemented"
}
<#
   ______  _______   ___   _    _                           _         _____ _               _
 |  ____|/ ____\ \ / (_) | |  | |                         | |       / ____| |             | |
 | |__  | (___  \ V / _  | |  | |_ __   __ _ _ __ __ _  __| | ___  | |    | |__   ___  ___| | _____ _ __
 |  __|  \___ \  > < | | | |  | | '_ \ / _` | '__/ _` |/ _` |/ _ \ | |    | '_ \ / _ \/ __| |/ / _ \ '__|
 | |____ ____) |/ . \| | | |__| | |_) | (_| | | | (_| | (_| |  __/ | |____| | | |  __/ (__|   <  __/ |
 |______|_____//_/ \_\_|  \____/| .__/ \__, |_|  \__,_|\__,_|\___|  \_____|_| |_|\___|\___|_|\_\___|_|
                                | |     __/ |
                                |_|    |___/
#>
function ESXiUpgradeChecker
{
     param (
          [Parameter()]
          [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
          $VMhost
      )

     $AllRunningTasks = Get-Task

     if ($AllRunningTasks.ObjectId -contains $VMhost.Id)
     {
          Write-Host "Upgrade task detected for $($VMhost.Id)"

          $HostTask = $AllRunningTasks | Where-Object { $_.ObjectId -match $VMhost.Id }
          if ($HostTask.PercentComplete -lt 100)
          {
               Write-Host 'Host still in progress'
          }
     }
     elseif ($AllRunningTasks.ObjectId -notcontains $VMhost.Id)
     {
          $VMhost.ExtensionData.setCustomValue('UpgradeState', '')
          $VMhost.ExtensionData.setCustomValue('State', 'live')

          Set-VMHost -VMHost $VMhost.Name -State Connected -Confirm:$false -RunAsync
          Write-Host 'Remanaging Host'
          $RemanageParams = @{
              NodeIP          = $VMhostObject.ManagementIP
              OrionCredential = $OrionCredential
              State           = "Managed"
          }
          Set-OrionNodeState @RemanageParams

          $ClusterBuffers = Get-VMHost $VMhost.Name | Get-Cluster | Get-VMHost | Where-Object { $_.Name -match 'buf' }
          $BorrowedBuffer = $ClusterBuffers | Get-Annotation | Where-Object { $_.Name -eq 'UpgradeBuffer' -and $Value -eq 'yes' }
          $BorrowedBuffer.AnnotatedEntity   | Set-VMHost -State "Maintenance" -Confirm:$false -Evacuate
          $BorrowedBuffer.AnnotatedEntity   | Move-VMHost -Destination (Get-Cluster ncop-vmclu-res-prod-sr6-21) -Confirm:$false
          $BorrowedBuffer.AnnotatedEntity.ExtensionData.setCustomValue('UpgradeBuffer', 'no')
     }
}

<#
    _____      _     _    _                           _       _     _       __      ____  __ _               _
  / ____|    | |   | |  | |                         | |     | |   | |      \ \    / /  \/  | |             | |
 | |  __  ___| |_  | |  | |_ __   __ _ _ __ __ _  __| | __ _| |__ | | ___   \ \  / /| \  / | |__   ___  ___| |_ ___
 | | |_ |/ _ \ __| | |  | | '_ \ / _` | '__/ _` |/ _` |/ _` | '_ \| |/ _ \   \ \/ / | |\/| | '_ \ / _ \/ __| __/ __|
 | |__| |  __/ |_  | |__| | |_) | (_| | | | (_| | (_| | (_| | |_) | |  __/    \  /  | |  | | | | | (_) \__ \ |_\__ \
  \_____|\___|\__|  \____/| .__/ \__, |_|  \__,_|\__,_|\__,_|_.__/|_|\___|     \/   |_|  |_|_| |_|\___/|___/\__|___/
                          | |     __/ |
                          |_|    |___/
#>
Function Get-UpgradeAbleVMHosts
{
    param (
        [Parameter()]
        $AllVMhostObjects
    )

    $StateDate = Get-Date -Format 'dd/MM/yyyy'
    $Global:UpgradeReadyVMHosts = $AllVMhostObjects | Where-Object { 
        $_.State           -eq 'enteringmaintenanace' -and 
        $_.ConnectionState -eq 'Maintenance' } | Select-Object -First 1

    if ($null -eq $Global:UpgradeReadyVMHosts)
    {
        Write-Host 'No upgrade ready hosts detected'
    }
    else
    {
        Write-Host 'The following hosts are upgrade ready'
        Write-Host "$($Global:UpgradeReadyVMHosts.Name)"
        $Global:UpgradeAbleVMHost.ExtensionData.setCustomValue('State', 'upgradeready')
        $Global:UpgradeAbleVMHost.ExtensionData.setCustomValue('StateDate', $StateDate)
    }
}

function Set-VMHostMarkedForUpgrade
{
    $AllUpgradeAbleClusters = (Get-Cluster | Get-Annotation | Where-Object {$_.Name -eq 'AutoUpgradeAble' -and $_.Value -eq 'yes'}).AnnotatedEntity
    $AllBaselines           = Get-Baseline -TargetType Host -BaselineType Upgrade -BaselineContentType Static

    Foreach ($Cluster in $AllUpgradeAbleClusters)
    {
        $vCenterSpecificBaseline = $AllBaselines | Where-Object {$_.Uid.SPlit('@').Split(':')[1] -eq $Cluster.Uid.SPlit('@').Split(':')[1]}
        $ClusterVMHosts          = $Cluster | Get-VMHost

        Foreach ($VMhost in $ClusterVMHosts)
        {
            $VMhostVersion       = $VMhost.Version + ', ' + $VMhost.Build
            $AvaliableBaseLine = $vCenterSpecificBaseline.UpgradeRelease.Version + ', ' + $vCenterSpecificBaseline.UpgradeRelease.Build

            #Validates if the hosts version matches the baselines
            if ($VMhostVersion -eq $AvaliableBaseLine)
            {
                Write-Host "$($VMhost.Name) is at the highest version avaliable"
            }

            #Validates if the hosts major version is lower than the avaliable major version
            if ($VMhost.Version -lt $vCenterSpecificBaseline.UpgradeRelease.Version)
            {
                Write-Host "$($VMhost.Name) has a lower version than the highest avaliable - marking host for upgrade"
                $VMhost.ExtensionData.setCustomValue('State', 'MarkedForAutoUpgrade')
                $Cluster.ExtensionData.setCustomValue('MarkedForAutoUpgrade', 'Yes')
            }

            #Validates if the hosts version is the same as the baseline, and if the build version is lower
            if ($VMhost.Version -eq $vCenterSpecificBaseline.UpgradeRelease.Version -and $VMhost.Build -lt $vCenterSpecificBaseline.UpgradeRelease.Build)
            {
                Write-Host "$($VMhost.Name) has the correct version but a lower build - marking host for upgrade"
                $VMhost.ExtensionData.setCustomValue('State', 'MarkedForAutoUpgrade')
                $Cluster.ExtensionData.setCustomValue('MarkedForAutoUpgrade', 'Yes')
            }
        }

        #Validates if cluster is completely up to patch baseline
        $ClusterUpgradeValidation = $Cluster | Get-VMHost | Select-Object Version, Build
        if ($ClusterUpgradeValidation.Version -notmatch $vCenterSpecificBaseline.UpgradeRelease.Version -or $ClusterUpgradeValidation.Build -notmatch $vCenterSpecificBaseline.UpgradeRelease.Build)
        {
            Write-Host "$($Cluster.Name) marked for upgrade"
            $Cluster.ExtensionData.setCustomValue('MarkedForAutoUpgrade', 'Yes')
        }
        else
        {
            Write-Host "$($Cluster.Name) not marked for upgrade"
            $Cluster.ExtensionData.setCustomValue('MarkedForAutoUpgrade', 'No')
        }
    }
}

<#
     _____ _             _   _    _           _       _        ____   ____      ____  __ _    _           _
   / ____| |           | | | |  | |         | |     | |      / __ \ / _\ \    / /  \/  | |  | |         | |
  | (___ | |_ __ _ _ __| |_| |  | |_ __   __| | __ _| |_ ___| |  | | |_ \ \  / /| \  / | |__| | ___  ___| |_
   \___ \| __/ _` | '__| __| |  | | '_ \ / _` |/ _` | __/ _ \ |  | |  _| \ \/ / | |\/| |  __  |/ _ \/ __| __|
   ____) | || (_| | |  | |_| |__| | |_) | (_| | (_| | ||  __/ |__| | |    \  /  | |  | | |  | | (_) \__ \ |_
  |_____/ \__\__,_|_|   \__|\____/| .__/ \__,_|\__,_|\__\___|\____/|_|     \/   |_|  |_|_|  |_|\___/|___/\__|
                                  | |
                                  |_|
#>
function Start-UpdateOfVMHost
{
    $SelectedCluster    = Get-VMHost $VMhostObject.Name | Get-Cluster
    $SelectedBufferHost = $SelectedCluster         | Get-Datastore | Select-Object -First 1 | Get-VMHost | Where-Object {$_.Name -match 'buf' -and $_.state -eq 'Maintenance'} | Select-Object -First 1
    $AllHosts           = $SelectedCluster         | Get-VMHost
    $TotalMhz           = ($AllHosts.CpuTotalMhz   | Measure-Object -Sum).Sum + ($SelectedBufferHost.CpuTotalMhz | Measure-Object -Sum).Sum
    $TotalMzUsage       = ($AllHosts.CpuUsageMhz   | Measure-Object -Sum).Sum + ($SelectedBufferHost.CpuUsageMhz | Measure-Object -Sum).Sum
    $TotalMemory        = ($AllHosts.MemoryTotalGB | Measure-Object -Sum).Sum + ($SelectedBufferHost.MemoryTotalGB | Measure-Object -Sum).Sum
    $TotalMemoryUsage   = ($AllHosts.MemoryUsageGB | Measure-Object -Sum).Sum + ($SelectedBufferHost.MemoryUsageGB | Measure-Object -Sum).Sum

    $CalculatedTotal   = $TotalMhz + ($TotalMhz - ($AllHosts | Where-Object {$_.Name -eq $VMhostObject.Name}).CpuTotalMhz) * 0.35
    $CalculatedTotalGb = $TotalMemory + ($TotalMemory - ($AllHosts | Where-Object {$_.Name -eq $VMhostObject.Name}).MemoryTotalGB) * 0.35

    if ($TotalMzUsage -lt $CalculatedTotal)
    {
        Write-Host 'Cluster has enough CPU Usage for host upgrade, Validating memory'
        if ($TotalMemoryUsage -lt $CalculatedTotalGb)
        {
            Write-Host 'Cluster has enough Memory for host upgrade, upgrading host'

            $NumberOfHostsInProgress = (Get-VMHost $VMhostObject.Name | Get-Cluster | Get-VMHost | Get-Annotation | Where-Object {$_.Name -eq 'State' -and $_.Value -eq 'SelectedForUpgrade'}).AnnotatedEntity
            if ($NumberOfHostsInProgress.Count -lt 1)
            {
                Write-Host 'Currently no hosts are being upgraded, upgrading first host in cluster'

                Write-Host 'Adding bufferhost if possible'
                Move-VMHost $SelectedBufferHost -Destination $SelectedCluster -Confirm:$false

                Set-VMHost $SelectedBufferHost -State "Maintenance" -Confirm:$false -Evacuate
                $SelectedBufferHost.ExtensionData.setCustomValue('UpgradeBuffer', 'yes')
                $SelectedBufferHost.ExtensionData.setCustomValue('BufferSourceLocation', $SelectedBufferHost.Parent.Name)
                $NumberOfHostsMarked = (Get-VMHost $VMhostObject.Name | Get-Cluster | Get-VMHost | Get-Annotation | Where-Object {$_.Name -eq 'State' -and $_.Value -eq 'MarkedForAutoUpgrade'}).AnnotatedEntity | Select-Object -First 1
                $NumberOfHostsMarked.ExtensionData.setCustomValue('State', 'SelectedForUpgrade')

                Write-Host "$($VMhostObject.Name) scheduled for maintenance mode due to upgrade"
                $NumberOfHostsMarked.ExtensionData.setCustomValue('UpgradeState', 'ReadyForMaintenanceMode')
                $NumberOfHostsInProgress = (Get-VMHost $VMhostObject.Name | Get-Cluster | Get-VMHost | Get-Annotation | Where-Object {$_.Name -eq 'State' -and $_.Value -eq 'SelectedForUpgrade'}).AnnotatedEntity
            }
            elseif ($NumberOfHostsInProgress.Count -ge 1)
            {
                Write-Host 'An upgrade is already in progress, awaiting completion'
            }

            $NumberOfHostsInProgress = (Get-VMHost $VMhostObject.Name | Get-Cluster | Get-VMHost | Get-Annotation | Where-Object {$_.Name -eq 'State' -and $_.Value -eq 'SelectedForUpgrade'}).AnnotatedEntity
        }
    }
}

<#
  _    _           _       _        __      ____  __ _               _   
 | |  | |         | |     | |       \ \    / /  \/  | |             | |  
 | |  | |_ __   __| | __ _| |_ ___   \ \  / /| \  / | |__   ___  ___| |_ 
 | |  | | '_ \ / _` |/ _` | __/ _ \   \ \/ / | |\/| | '_ \ / _ \/ __| __|
 | |__| | |_) | (_| | (_| | ||  __/    \  /  | |  | | | | | (_) \__ \ |_ 
  \____/| .__/ \__,_|\__,_|\__\___|     \/   |_|  |_|_| |_|\___/|___/\__|
        | |                                                              
        |_|                                                              
#>
Function Update-VMHost
{
    Write-Host "Entering Update-VMHost Function"

    # Gathers all update baselines
    $AllBaselines     = Get-Baseline -TargetType Host -BaselineType Upgrade -BaselineContentType Static
    $SelectedHost     = Get-VMHost $VMhostObject.Name
    $SelectedBaseline = $AllBaselines | WHere-Object {$_.Uid.SPlit('@').Split(':')[1] -eq $SelectedHost.Uid.SPlit('@').Split(':')[1]}

    $SelectedHost.ExtensionData.setCustomValue('UpgradeState','Upgrading-host')
    Attach-Baseline     -Entity $SelectedHost -Baseline $SelectedBaseline

    $RemediationParams = @{
      Entity                  = $SelectedHost
      Baseline                = $SelectedBaseline
      HostFailureAction       = "Retry"
      HostNumberOfRetries     = "10"
      HostDisableMediaDevices = $true 
      Confirm                 = $false
    }
    
    Remediate-Inventory @RemediationParams
    #Remediate-Inventory -Entity $SelectedHost -Baseline $SelectedBaseline -HostFailureAction Retry -HostNumberOfRetries 10 -HostDisableMediaDevices $true -Confirm:$false
}

#endregion
