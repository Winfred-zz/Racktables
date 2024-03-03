# This is a powershell module that connects to the Racktables MySQL DB and performs various tasks.
# It is not intended for beginners. It is intended for people who have a good understanding of PowerShell.

# Don't just implement the functions, read the help functions provided first and ask questions if it's not entirely clear.
# This has been tested extensively in production, but only in a single environment. It might not work in your environment without modification.

# This module directly makes changes to the DB, this is inherently somewhat risky, since on updates of the DB, the tables might change.
# Always make sure you have a working backup of the DB before running any of these functions.

# set MySQLHost.txt in the script root (probably the module folder if used as intended) to the hostname/ip of your MySQL server.

# MySql.Data.dll is required for this module to work. It can be downloaded from https://dev.mysql.com/downloads/connector/net

# get-help Add-RTIPAddress -full
# get-help Add-RTLink -full
# get-help Add-RTObject -full
# get-help Add-RTPort -full
# get-help Compare-RTMACAddressesToWMI -full
# get-help Connect-RTTomysql -full
# get-help Convert-IPtoINT64 -full
# get-help Get-RTChassisdetails -full
# get-help Get-RTComputernamefromip -full
# get-help Get-RTComputerNameFromMAC -full
# get-help Get-RTDiskarraydetails -full
# get-help Get-RTIpaddresses -full
# get-help Get-RTServerDetails -full
# get-help Get-RTObjectid -full
# get-help Get-RTObjectnamesfromrack -full
# get-help Get-RTPDUDetails -full
# get-help Get-RTSwitchDetails -full
# get-help Move-RTLinks -full
# get-help Set-RTTag -full
# get-help Update-RTAssetTag -full
# get-help Remove-RTAttribute -full
# get-help Remove-RTIPAddress -full
# get-help Remove-RTLink -full
# get-help Update-RTAttribute -full
# get-help Update-RTComputername -full
# get-help Update-RTCputype -full
# get-help Update-RTharddisk -full
# get-help Update-RTHistory -full
# get-help Update-RTHwtype -full
# get-help Update-RTRacktables -full
# get-help Update-RTRacktablestwice -full
# get-help Update-RTRam -full
# get-help Update-RTSerialnumber -full

#=====================================================================
# Out-MyLogFile
#=====================================================================
Function Out-MyLogFile
{
<#
.SYNOPSIS
	Takes in from the pipeline and will write it to a file.
.DESCRIPTION
	Either appends or Overrides the data in the text file.
	This is mainly used for logging and fixes the issue where if you run multiple jobs at once that log to the same log file, the file is locked and you get an error.
.EXAMPLE
	"abcd" | Out-MyLogFile "D:\Temp\test.txt" -append
.EXAMPLE	
	"abcde" | Out-MyLogFile "D:\Temp\test.txt"
.NOTES
	Maybe it should be throwing an error? Not sure.
.LINK

#>
[cmdletbinding()]
param(
	[parameter(ValueFromPipeline)]$TextToWrite,
	[Parameter(Position=0, Mandatory=$True)]$FilePath,
	[Parameter(Position=1)][switch]$Append
)
	$i = 0 
	Do{
		$WriteError = $Null
		try
		{
			$TextToWrite | out-file -FilePath $FilePath -append:$Append
		}
		catch [System.IO.IOException] 
		{
			$i++
			$WriteError = $True
		}
	}until(!($WriteError) -or ($i -ge 100))
	if($WriteError)
	{
		write-host "Tried to write 100 times, to $FilePath but I failed each time." -fore red
	}
}

#=====================================================================
#Connect-RTToMysql
#=====================================================================
Function Connect-RTToMysql
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB
.DESCRIPTION

.EXAMPLE
	Connect-RTToMysql -Query "SELECT * FROM Object;"
.NOTES
	
.LINK

#>
Param(
  [Parameter(
  Mandatory = $true,
  ParameterSetName = '',
  ValueFromPipeline = $true)]
  [string]$Query
  )
	$CredsFile = join-path $PsScriptRoot "MySQL-Creds-$($env:Username).xml"
	$creds = Get-MyCredential $CredsFile
	$MySQLHost = get-content (join-path $PsScriptRoot "MySQLHost.txt")
	
	$MySQLDatabase = 'racktables'

	$ConnectionString = "server=" + $MySQLHost + ";port=3306;uid=" + $creds.UserName + ";pwd=" + $creds.GetNetworkCredential().password + ";database="+$MySQLDatabase + ";convert zero datetime=True"

	Try
	{
		$mysqldatadll = join-path $PsScriptRoot "MySql.Data.dll"
		[void][system.reflection.Assembly]::LoadFrom($mysqldatadll)
		$Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
		$Connection.ConnectionString = $ConnectionString
		$Connection.Open()
	
		$Command = New-Object MySql.Data.MySqlClient.MySqlCommand($Query, $Connection)
		$DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
		$DataSet = New-Object System.Data.DataSet
		#$RecordCount = $dataAdapter.Fill($dataSet, "data")
		$NULL = $dataAdapter.Fill($dataSet, "data")
		$DataSet.Tables[0]
	}Catch{
		Write-Host "ERROR : Unable to run query : $query `n $Error"
		$Error
	}Finally{
		$Connection.Close()
	}
}

#=====================================================================
# Update-RTAssetTag
#=====================================================================
Function Update-RTAssetTag
{
<#
.SYNOPSIS
	Inserts or updates the asset tag in racktables.
.DESCRIPTION
	
.EXAMPLE
	Update-RTAssetTag -ObjectName "VGER2" -AssetTag "01110"
.NOTES
	
.LINK

#>
Param(
[Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
[Parameter(Position=1, Mandatory=$True)][string]$AssetTag
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTAssetTag.log"
	$ObjectQuery = Connect-RTToMysql -Query "Select asset_no FROM Object
											WHERE name = '$ObjectName';"
	if(!$ObjectQuery)
	{
		$Message = "$(get-date) Unable to find $ObjectName in Racktables - return"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	if($ObjectQuery.asset_no -is [System.DBNull])
	{
		$Comment = "Updating AssetTag from nothing to $AssetTag"
	}else{
		$Comment = "Updating AssetTag from $($ObjectQuery.asset_no) to $AssetTag"
	}
	
	Update-RTHistory -ObjectName $ObjectName -comment $Comment
	
	$Message = "$(get-date) $Comment"
	write-host $Message -fore cyan
	$Message | Out-MyLogFile $logfile -append
	
	Connect-RTToMysql -Query "UPDATE Object
							SET asset_no = '$AssetTag'
							WHERE name = '$ObjectName';"
}

#=====================================================================
#Move-RTLinks
#=====================================================================
Function Move-RTLinks
{
<#
.SYNOPSIS
	Moves links from one switch (or server) to another.
	Assumes that there is no change to the port numbering.
.DESCRIPTION
	
.EXAMPLE
	Move-RTLinks -OldSwitch "SWITCH01" -NewSwitch "SWITCH01-new" -OldSwitchPortNamePattern "Gi 1/" -NewSwitchPortNamePattern "GE 1/0/"
.EXAMPLE
	Move-RTLinks -OldSwitch "SWITCH01" -NewSwitch "SWITCH01-new" -OldSwitchPortNamePattern "Gii 1/" -NewSwitchPortNamePattern "GEE 1/0/"
.EXAMPLE
	Move-RTLinks -OldSwitch "SWITCH01-new" -NewSwitch "SWITCH01" -OldSwitchPortNamePattern "GEE 1/0/" -NewSwitchPortNamePattern "Gii 1/"
.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$OldSwitch,
  [Parameter(Position=1, Mandatory=$True)][string]$NewSwitch,
  [Parameter(Position=2, Mandatory=$True)][string]$OldSwitchPortNamePattern,
  [Parameter(Position=3, Mandatory=$True)][string]$NewSwitchPortNamePattern
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Move-RTLinks.log"

	$Message = "$(get-date) Move-RTLinks -OldSwitch `"$($OldSwitch)`" -NewSwitch `"$($NewSwitch)`" -OldSwitchPortNamePattern `"$($OldSwitchPortNamePattern)`" -NewSwitchPortNamePattern `"$($NewSwitchPortNamePattern)`""
	write-host $Message -fore cyan
	$Message | Out-MyLogFile $logfile -append

	$OldSwitchPorts = Connect-RTToMysql -Query "SELECT Port.id, Port.name as Portname, Object.name as ObjectName, Link.porta, Link.portb FROM Port
												LEFT JOIN Object ON Object.id = Port.object_id
												LEFT JOIN Link ON (Link.portb = Port.id OR Link.porta = Port.id)
												WHERE Object.name = '$($OldSwitch)';"

	$NewSwitchPorts = Connect-RTToMysql -Query "SELECT Port.id, Port.name as Portname, Object.name as ObjectName, Link.porta, Link.portb FROM Port
												LEFT JOIN Object ON Object.id = Port.object_id
												LEFT JOIN Link ON (Link.portb = Port.id OR Link.porta = Port.id)
												WHERE Object.name = '$($NewSwitch)';"

	if(!$OldSwitchPorts -or !$NewSwitchPorts)
	{
		$Message = "$(get-date) Couldn't find any ports on either the old or the new switch. break"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		break
	}

	$ConnectedPortsOldSwitch = $OldSwitchPorts | where-object {$_.PortName.StartsWith($OldSwitchPortNamePattern) -and !($_.porta -is [System.DBNull]) -and !($_.portb -is [System.DBNull])}

	$UnConnectedPortsNewSwitch = $NewSwitchPorts | where-object {$_.PortName.StartsWith($NewSwitchPortNamePattern) -and ($_.porta -is [System.DBNull]) -and ($_.portb -is [System.DBNull])}

	if(!$ConnectedPortsOldSwitch -or !$UnConnectedPortsNewSwitch)
	{
		$Message = "$(get-date) Couldn't find any connected ports on the old switch or unconnected ports on the new switch. break"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		break
	}

	# making sure that all the linked matching ports on the old switch are available as open ports on the new switch
	foreach($ReducedPortName in $ConnectedPortsOldSwitch.PortName.Replace($OldSwitchPortNamePattern,""))
	{
		if(!($UnConnectedPortsNewSwitch.PortName.Replace($NewSwitchPortNamePattern,"") -eq $ReducedPortName))
		{
			$Message = "$(get-date) Connected Port $($OldSwitchPortNamePattern)$($ReducedPortName) on OldSwitch $OldSwitch is not available as a $($NewSwitchPortNamePattern) port on NewSwitch $NewSwitch. break"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			break
		}
	}

	foreach($Port in $ConnectedPortsOldSwitch)
	{
		$PortAorPortB = $Null
		#$NewPortId = $Null
		$NewPortName = $Null
		$NewPort = $Null
		$ConnectingPort = $Null
		if($Port.id -eq $Port.porta)
		{
			$PortAorPortB = "porta"
			$ConnectingPort = $Port.portb
		}
		if($Port.id -eq $Port.portb)
		{
			$PortAorPortB = "portb"
			$ConnectingPort = $Port.porta
		}
		$NewPortName = $Port.PortName.Replace($OldSwitchPortNamePattern,$NewSwitchPortNamePattern)
		$NewPort = $NewSwitchPorts | where-object {$_.PortName -eq $NewPortName}
		if($NewPort)
		{
			if($PortAorPortB)
			{
				$Message = "$(get-date) Moving old port `"$($Port.Portname)`" id ($($Port.id)) on $OldSwitch to new port `"$($NewPort.Portname)`" id ($($NewPort.id)) on $NewSwitch - ConnectingPort id = $ConnectingPort"
				write-host $Message -fore cyan
				$Message | Out-MyLogFile $logfile -append
				Connect-RTToMysql -Query "DELETE FROM Link WHERE $PortAorPortB = '$($Port.id)'"
				Connect-RTToMysql -Query "INSERT INTO Link (porta, portb) VALUES ($($ConnectingPort), '$($NewPort.id)');"
			}else{
				$Port
				$Message = "$(get-date) No PortAorPortB for old port $($Port.Portname) on $OldSwitch (this shouldn't be possible I think)"
				write-host $Message -fore red
				$Port | Out-MyLogFile $logfile -append
				$Message | Out-MyLogFile $logfile -append
			}
		}else{
			$Port
			$Message = "$(get-date) No NewPort found on $NewSwitch for old port $($Port.Portname) on $OldSwitch"
			write-host $Message -fore red
			$Port | Out-MyLogFile $logfile -append
			$Message | Out-MyLogFile $logfile -append
		}
	}
}

#=====================================================================
#Get-RTObjectNamesFromRack
#=====================================================================
Function Get-RTObjectNamesFromRack
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns names of either computers, chassis or diskarrays in a rack
.DESCRIPTION
	Can either return just the names or an array of objects containing the name, rack, u and enclosure name (if applicable).
.EXAMPLE
	Get-RTObjectNamesFromRack -RackName RACK01 -ComputerNames
	Get-RTObjectNamesFromRack -RackName RACK01 -ChassisNames
	Get-RTObjectNamesFromRack -RackName RACK01 -DiskArrayNames

.EXAMPLE
	$ArrayOfComputerObjects = Get-RTObjectNamesFromRack -RackName RACK01 -ComputerNames -ReturnObjects
	
	$ArrayOfComputerObjects |ft

	Name        Rack    U Enclosure
	----        ----    - ---------
	SERVER01 RACK01 15 ENCLOSURE01
	SERVER02 RACK01 15 ENCLOSURE02
	
.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$RackName,
  [switch]$ComputerNames,
  [switch]$ChassisNames,
  [switch]$DiskArrayNames,
  [switch]$SwitchNames,
  [switch]$PDUNames,
  [switch]$ReturnObjects
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Get-RTObjectNamesFromRack.log"

	$ObjectNames = @()

	if($ComputerNames)
	{
		# These queries return duplicates for multi-u servers, so it needs to be deduplicated at the end.

		$SQLQueryServers = Connect-RTToMysql -Query "select * FROM Object
								LEFT JOIN RackSpace ON Object.id = RackSpace.object_id
								LEFT JOIN Rack ON RackSpace.rack_id = Rack.id
								Where Rack.name = '$RackName'
								AND Object.objtype_id = 4;"
		if($ReturnObjects)
		{
			$ProcessedServers = @()
			foreach($Row in $SQLQueryServers)
			{
				# this if statement needs to be there, because some servers take up two U, so they return multiple rows here
				if(!($ProcessedServers -eq $Row.name))
				{
					$Object = New-Object -TypeName System.Object
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $Row.name
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Rack" -Value $RackName
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "U" -Value $Row.unit_no
					#naughty naughty... loading objects into something called "name". I'm sorry.
					$ObjectNames += $Object
					$ProcessedServers += $Row.name
				}
			}
		}else{
			$ObjectNames += $SQLQueryServers.name
		}

		$SQLQueryChassis = Connect-RTToMysql -Query "select * FROM Object
													LEFT JOIN RackSpace ON Object.id = RackSpace.object_id
													LEFT JOIN Rack ON RackSpace.rack_id = Rack.id
													LEFT JOIN EntityLink ON Object.id = EntityLink.parent_entity_id
													Where Rack.name = '$RackName'
													AND Object.objtype_id = 1502;"
		if($ReturnObjects)
		{
			$ProcessedEnclosureIDs = @()
			foreach($row in $SQLQueryChassis)
			{
				# this if statement needs to be there, because most chassis take up two U, so they return multiple rows here
				if(!($ProcessedEnclosureIDs -eq $Row.id))
				{
					$ProcessedEnclosureIDs += $Row.id
					
					foreach($id in ($SQLQueryChassis | where-object {$_.id -eq $Row.id} | % {$_.child_entity_id} | sort-object | get-unique))
					{
						$SQLQueryBlade = Connect-RTToMysql -Query "select Object.name FROM Object
																	Where Object.id = '$id'"
						$Object = New-Object -TypeName System.Object
						Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $SQLQueryBlade.name
						Add-Member -inputObject $Object -MemberType NoteProperty -Name "Rack" -Value $RackName
						Add-Member -inputObject $Object -MemberType NoteProperty -Name "U" -Value $Row.unit_no
						Add-Member -inputObject $Object -MemberType NoteProperty -Name "Enclosure" -Value $Row.name
						$ObjectNames += $Object
					}
				}
			}
		}else{
			foreach($id in ($SQLQueryChassis.child_entity_id | sort-object | get-unique))
			{
				$SQLQueryBlades = Connect-RTToMysql -Query "select Object.name FROM Object
															Where Object.id = '$id'"
				$ObjectNames += $SQLQueryBlades.name
			}
		}
	}
	if($ChassisNames)
	{
		$ObjtypeId = 1502
		$SQLQueryObjects = Connect-RTToMysql -Query "select Object.name,unit_no FROM Object
								LEFT JOIN RackSpace ON Object.id = RackSpace.object_id
								LEFT JOIN Rack ON RackSpace.rack_id = Rack.id
								Where Rack.name = '$RackName'
								AND Object.objtype_id = $ObjtypeId;"
		if($ReturnObjects)
		{
			$ProcessedChassis = @()
			foreach($Row in $SQLQueryObjects)
			{
				# this if statement needs to be there, because some servers take up two U, so they return multiple rows here
				if(!($ProcessedChassis -eq $Row.name))
				{
					$Object = New-Object -TypeName System.Object
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $Row.name
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Rack" -Value $RackName
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "U" -Value $Row.unit_no
					#naughty naughty... loading objects into something called "name". I'm sorry.
					$ObjectNames += $Object
					$ProcessedChassis += $Row.name
				}
			}
		}else{
			$ObjectNames += $SQLQueryObjects.name
		}
	}
	if($DiskArrayNames)
	{
		$ObjtypeId = 5
		$SQLQueryObjects = Connect-RTToMysql -Query "select Object.name,unit_no FROM Object
								LEFT JOIN RackSpace ON Object.id = RackSpace.object_id
								LEFT JOIN Rack ON RackSpace.rack_id = Rack.id
								Where Rack.name = '$RackName'
								AND Object.objtype_id = $ObjtypeId;"
		if($ReturnObjects)
		{
			$ProcessedDiskArray = @()
			foreach($Row in $SQLQueryObjects)
			{
				# this if statement needs to be there, because some servers take up two U, so they return multiple rows here
				if(!($ProcessedDiskArray -eq $Row.name))
				{
					$Object = New-Object -TypeName System.Object
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $Row.name
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Rack" -Value $RackName
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "U" -Value $Row.unit_no
					#naughty naughty... loading objects into something called "name". I'm sorry.
					$ObjectNames += $Object
					$ProcessedDiskArray += $Row.name
				}
			}
		}else{
			$ObjectNames += $SQLQueryObjects.name
		}
	}
	if($SwitchNames)
	{
		$ObjtypeId = 8
		$SQLQueryObjects = Connect-RTToMysql -Query "select Object.name FROM Object
								LEFT JOIN RackSpace ON Object.id = RackSpace.object_id
								LEFT JOIN Rack ON RackSpace.rack_id = Rack.id
								Where Rack.name = '$RackName'
								AND Object.objtype_id = $ObjtypeId;"
		if($ReturnObjects)
		{
			$ProcessedSwitches = @()
			foreach($Row in $SQLQueryObjects)
			{
				# this if statement needs to be there, because some servers take up two U, so they return multiple rows here
				if(!($ProcessedSwitches -eq $Row.name))
				{
					$Object = New-Object -TypeName System.Object
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $Row.name
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Rack" -Value $RackName
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "U" -Value $Row.unit_no
					#naughty naughty... loading objects into something called "name". I'm sorry.
					$ObjectNames += $Object
					$ProcessedSwitches += $Row.name
				}
			}
		}else{
			$ObjectNames += $SQLQueryObjects.name
		}
	}
	if($PDUNames)
	{
		$ObjtypeId = 2
		$SQLQueryObjects = Connect-RTToMysql -Query "select Object.name FROM Object
								LEFT JOIN RackSpace ON Object.id = RackSpace.object_id
								LEFT JOIN Rack ON RackSpace.rack_id = Rack.id
								Where Rack.name = '$RackName'
								AND Object.objtype_id = $ObjtypeId;"
		if($ReturnObjects)
		{
			$ProcessedSwitches = @()
			foreach($Row in $SQLQueryObjects)
			{
				# this if statement needs to be there, because some servers take up two U, so they return multiple rows here
				if(!($ProcessedSwitches -eq $Row.name))
				{
					$Object = New-Object -TypeName System.Object
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $Row.name
					Add-Member -inputObject $Object -MemberType NoteProperty -Name "Rack" -Value $RackName
					#naughty naughty... loading objects into something called "name". I'm sorry.
					$ObjectNames += $Object
					$ProcessedSwitches += $Row.name
				}
			}
		}else{
			$ObjectNames += $SQLQueryObjects.name
		}
	}
	if(!$ReturnObjects)
	{
		$ObjectNames = $ObjectNames | sort-object | get-unique
		$ObjectNames | Out-MyLogFile $logfile -append
	}
	return $ObjectNames
}

#=====================================================================
#Get-RTIPaddresses
#=====================================================================
Function Get-RTIPaddresses
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns all IP addresses associated to a computername
.DESCRIPTION

.EXAMPLE
	Get-RTIPaddresses -ComputerName SERVER01

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$ComputerName 
)
	$SQLQuery = Connect-RTToMysql -Query "SELECT
											IPv4Allocation.ip as IPinINT,
											CAST(INET_NTOA(IPv4Allocation.ip) AS CHAR) as IP,
											IPv4Allocation.name as NICName,
											Object.name as ComputerName
											FROM IPv4Allocation
											LEFT JOIN Object ON Object.id = object_id
											WHERE Object.name = '$ComputerName'"

return $SQLQuery
}

#=====================================================================
#Get-RTComputerNameFromIP
#=====================================================================
Function Get-RTComputerNameFromIP
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns the ComputerName associated with an IP.
.DESCRIPTION

.EXAMPLE
	Get-RTComputerNameFromIP -IP "10.0.0.1"

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$IP 
)
	$SQLQuery = Connect-RTToMysql -Query "SELECT
											IPv4Allocation.ip as IPinINT,
											CAST(INET_NTOA(IPv4Allocation.ip) AS CHAR) as IP,
											IPv4Allocation.name as NICName,
											Object.name as ComputerName
											FROM IPv4Allocation
											LEFT JOIN Object ON Object.id = object_id"

	$ComputerNameRow = $SQLQuery | where-object {$_.IP -eq  $IP}
	return $ComputerNameRow.ComputerName
}

#=====================================================================
#Get-RTComputerNameFromMAC
#=====================================================================
Function Get-RTComputerNameFromMAC
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns the ComputerName associated with an MacAddress.
.DESCRIPTION

.EXAMPLE
	Get-RTComputerNameFromMAC -MAC "AAAAAAAAAAAA"

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$MAC 
)

	$MAC = $MAC.replace(":","").replace("-","")
	if($MAC.Length -eq 12)
	{
		$SQLQuery = Connect-RTToMysql -Query "select Object.name from Object
												LEFT JOIN Port on Port.object_id = Object.id
												WHERE Port.l2address = '$MAC'"

		if($SQLQuery)
		{
			return $SQLQuery.name
		}else{
			write-host "MacAddress $MAC was not found in Racktables" -fore red
		}
	}else{
		write-host "MacAddress $MAC wasn't 12 characters long" -fore red
	}
}

#=====================================================================
#Get-RTObjectID
#=====================================================================
Function Get-RTObjectID
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns the object id associated with a computername
.DESCRIPTION

.EXAMPLE
	Get-RTObjectID -ComputerName SERVER01

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$ComputerName
)
	$SQLQuery = Connect-RTToMysql -Query "select id from Object where name = '$ComputerName'"

	if($SQLQuery)
	{
		return $SQLQuery.id
	}
}

#=====================================================================
#Get-RTServerDetails
#=====================================================================
Function Get-RTServerDetails
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns all:
	IPs,
	Ports,
	Details
	associated to a computername
.DESCRIPTION

.EXAMPLE
	$TopLevelObject = Get-RTServerDetails -ComputerName SERVER01
	
	$TopLevelObject.ComputerName
	
	SERVER01
	
	$TopLevelObject.IPs
	
	   IPinINT IP             NICName
	   ------- --             -------
	1231231231 10.0.0.1   NIC1
	1231231231 10.0.0.2 iDrac
	1231231231 15.15.15.1 NIC2
	
	$TopLevelObject.Ports
	
	PortName PortLabel InterfaceType MacAddress
	-------- --------- ------------- ----------
	idrac              1000Base-T    AA:AA:AA:AA:AA:AA
	NIC1               1000Base-T    AA:AA:AA:AA:AA:AA
	NIC2               1000Base-T    AA:AA:AA:AA:AA:AA
	NIC3               1000Base-T    AA:AA:AA:AA:AA:AA
	NIC4               1000Base-T    AA:AA:AA:AA:AA:AA

	$TopLevelObject.Details
	
	VisibleLabel : SERVER01
	SerialNumber : ASDASD
	HWType       : Dell PowerEdge%GPASS%R620
	OS           : Microsoft%GSKIP%Windows 7
	RAM          : 64GB
.EXAMPLE
	$ArrayOfRackTablesComputerObjects = Get-RTServerDetails -RackName RACK01
.NOTES
	This funciton might require manual changes to the SQL queries to suit your environment.
.LINK

#>
Param(
  [string]$RackName,
  [string]$ComputerName,
  [int]$ObjectId
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Get-RTServerDetails.log"
		
	if($ObjectId)
	{
		# 50333 = C-Series blades
		# 4 = server
		# 1504 = VM
		# this query is in here twice, it's a little lower as well under IdQuery.
		$NameQuery = Connect-RTToMysql -Query "Select name,asset_no from Object where id = $ObjectId AND (objtype_id = 50333 OR objtype_id = 4 OR objtype_id = 1504)"
		if($NameQuery)
		{
			$ComputerName = $NameQuery.name
		}
		if(!$ComputerName)
		{
			$Message = "$(get-date) ComputerName not found for ObjectId $ObjectId in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}

	$AllComputerNamesInRack = @()
	if($RackName)
	{
		$AllComputerNamesInRack += Get-RTObjectNamesFromRack -RackName $RackName -ComputerNames -ReturnObjects
		if($AllComputerNamesInRack.count -eq 0)
		{
			$Message = "$(get-date) RackName $RackName has no Computer Objects in Racktables"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
	}else{
		$Object = New-Object -TypeName System.Object
		Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $ComputerName
		$AllComputerNamesInRack += $Object
	}

	$ArrayOfRackTablesComputerObjects = @()

	foreach($ComputerNameObject in $AllComputerNamesInRack)
	{
		$ComputerName = $Null
		$ComputerName = $ComputerNameObject.name
		$IdQuery = Connect-RTToMysql -Query "Select id,asset_no from Object where name = '$ComputerName' AND (objtype_id = 50333 OR objtype_id = 4 OR objtype_id = 1504)"
		if($IdQuery)
		{
			$ObjectId = $IdQuery.id
			if($IdQuery.asset_no.ToString() -ne "")
			{
				$AssetTag = $IdQuery.asset_no
			}else{
				$AssetTag = $Null
			}
		}
		if(!$ObjectId)
		{
			$Message = "$(get-date) ObjectId for ComputerName $ComputerName not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
	
		$TopLevelObject = new-object object
		$TopLevelObject | add-member -membertype NoteProperty -Name "ComputerName" -value $ComputerName
		$TopLevelObject | add-member -membertype NoteProperty -Name "ObjectId" -value $ObjectId
		$TopLevelObject | add-member -membertype NoteProperty -Name "AssetTag" -value $AssetTag
		if($ComputerNameObject.Rack)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Rack" -value $ComputerNameObject.Rack
		}
		if($ComputerNameObject.U)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "U" -value $ComputerNameObject.U
		}
		if($ComputerNameObject.Enclosure)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Enclosure" -value $ComputerNameObject.Enclosure
		}
		
		#----------------------------------------------------------------------
		# IPs
		#----------------------------------------------------------------------
		$IPSQLQuery = Connect-RTToMysql -Query "SELECT
												IPv4Allocation.ip as IPinINT,
												CAST(INET_NTOA(IPv4Allocation.ip) AS CHAR) as IP,
												IPv4Allocation.name as NICName,
												Object.name as ComputerName
												FROM IPv4Allocation
												LEFT JOIN Object ON Object.id = object_id
												WHERE Object.name = '$ComputerName'"
								
		$IPObjects = @()
		foreach($Row in $IPSQLQuery)
		{
			$Object = new-object object
			$Object | add-member -membertype NoteProperty -Name "IPinINT" -value $Row.IPinINT
			$Object | add-member -membertype NoteProperty -Name "IP" -value $Row.IP
			$Object | add-member -membertype NoteProperty -Name "NICName" -value $Row.NICName
			
			$IPObjects += $Object
		}
		if($IPObjects)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "IPs" -value $IPObjects
		}else{
				$Message = "$(get-date) ComputerName $ComputerName has no IPs in Racktables."
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		# Ports
		#----------------------------------------------------------------------
		#----------------------------------------------------------------------
		# Ports
		#----------------------------------------------------------------------
		# have to do it like this, I had a combined porta or portb query before, but it takes over 25 seconds to complete on the prod DB (but it's fast on the RWS DB? whatever).
		$PortQuery = @()

		$PortQuery += Connect-RTToMysql -Query "SELECT
														Port.name,
														Port.label,
														Port.id,
														PortOuterInterface.oif_name,
														l2address,
														Link.portb,
														Link.porta
														FROM Port
														LEFT JOIN Link ON Link.portb = Port.id
														LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
														LEFT JOIN Object ON Object.id = object_id
														WHERE Object.name = '$ComputerName'"
		$PortQuery += Connect-RTToMysql -Query "SELECT
														Port.name,
														Port.label,
														Port.id,
														PortOuterInterface.oif_name,
														l2address,
														Link.portb,
														Link.porta
														FROM Port
														LEFT JOIN Link ON Link.porta = Port.id
														LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
														LEFT JOIN Object ON Object.id = object_id
														WHERE Object.name = '$ComputerName'"
		$ConnectedPorts = @()
		$ConnectedPorts += $PortQuery | where-object {!($_.porta -is [System.DBNull])}
		# now we have to filter out the duplicate ports returned in the query above.
		$FilteredPortQuery = @()
		foreach($Port in $PortQuery)
		{
			if($ConnectedPorts.Name -eq $Port.name)
			{
				# it's a connected port
				if($Port.porta -is [System.DBNull])
				{
					# but this is the query that doesn't have the porta/portb in it, so drop it
				}else{
					$FilteredPortQuery += $Port
				}
			}else{
				# oh, it's not connected, so add it just once
				$ConnectedPorts += $Port
				$FilteredPortQuery += $Port
			}
		}
		$PortObjects = @()
		foreach($Row in $FilteredPortQuery)
		{
			$Object = new-object object
			$Object | add-member -membertype NoteProperty -Name "PortName" -value $Row.name
			$Object | add-member -membertype NoteProperty -Name "PortLabel" -value $Row.label
			$Object | add-member -membertype NoteProperty -Name "InterfaceType" -value $Row.oif_name
			if($Row.l2address)
			{
				if(!($Row.l2address -is [System.DBNull]))
				{
					$Object | add-member -membertype NoteProperty -Name "MacAddress" -value ($Row.l2address.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")).Trim()
				}
			}
			if(!($Row.porta -is [System.DBNull]))
			{
				if($Row.porta -eq $Row.id)
				{
					$RemotePortID = $Row.portb
				}else{
					$RemotePortID = $Row.porta
				}
				$RemotePortQuery = Connect-RTToMysql -Query "SELECT
																Port.name as PortName,
																Port.label,
																PortOuterInterface.oif_name,
																l2address,
																Link.portb,
																Link.porta,
																Object.name as ObjectName
																FROM Port
																LEFT JOIN Link ON (Link.portb = Port.id OR Link.porta = Port.id)
																LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
																LEFT JOIN Object ON Object.id = object_id
																WHERE Port.id = $RemotePortID"
				if(!$RemotePortQuery)
				{
					$Message = "$(get-date) something went very wrong. Could not find RemotePortQuery for $RemotePortID, $ComputerName, $($Row.name)"
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					return
				}
				if($RemotePortQuery -is [array])
				{
					$Message = "$(get-date) something went very wrong. Received an array for RemotePortQuery for $RemotePortID, $SwitchName, $($Row.name)"
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					return
				}
				$Object | add-member -membertype NoteProperty -Name "RemotePortName" -value $RemotePortQuery.PortName
				$Object | add-member -membertype NoteProperty -Name "RemoteObjectName" -value $RemotePortQuery.ObjectName
				$Object | add-member -membertype NoteProperty -Name "Connected" -value $True
				if(!($RemotePortQuery.l2address -is [System.DBNull]))
				{
					if($RemotePortQuery.l2address.length -gt 0)
					{
						$Object | add-member -membertype NoteProperty -Name "RemoteMacAddress" -value ($RemotePortQuery.l2address.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")).Trim()
					}
				}
			}else{
				$Object | add-member -membertype NoteProperty -Name "Connected" -value $False
			}
			$PortObjects += $Object
		}
		if($PortObjects)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Ports" -value $PortObjects
		}else{
				$Message = "$(get-date) ComputerName $ComputerName has no Ports in Racktables."
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		# Details
		#----------------------------------------------------------------------
		$DetailQuery = Connect-RTToMysql -Query "Select
												Attribute.name as AttributeName,
												Dictionary.dict_value as Type,
												AttributeValue.string_value as string_value
												FROM Object
												LEFT JOIN AttributeValue ON Object.id = AttributeValue.object_id
												LEFT JOIN Attribute ON AttributeValue.attr_id = Attribute.id
												LEFT JOIN Dictionary ON Dictionary.dict_key = AttributeValue.uint_value
												WHERE Object.name = '$ComputerName'"
		if($DetailQuery)
		{
			$DetailsObject = new-object object
			$LabelQuery = Connect-RTToMysql -Query "SELECT Object.label FROM Object	WHERE Object.name = '$ComputerName'"
			
			$DetailsObject | add-member -membertype NoteProperty -Name "VisibleLabel" -value $LabelQuery.label
			$DetailsObject | add-member -membertype NoteProperty -Name "SerialNumber" -value ($DetailQuery | where-object {$_.AttributeName -eq "Service Tag/Serial Number"}).string_value
			$DetailsObject | add-member -membertype NoteProperty -Name "HWType" -value ($DetailQuery | where-object {$_.AttributeName -eq "HW type"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "SWType" -value ($DetailQuery | where-object {$_.AttributeName -eq "SW type"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "RAM" -value ($DetailQuery | where-object {$_.AttributeName -eq "RAM"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "CPUCount" -value ($DetailQuery | where-object {$_.AttributeName -eq "CPU Count"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "CPUType" -value ($DetailQuery | where-object {$_.AttributeName -eq "CPU Type"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "Controller" -value ($DetailQuery | where-object {$_.AttributeName -eq "Controller"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "Controller2" -value ($DetailQuery | where-object {$_.AttributeName -eq "Controller2"}).Type
			
			$DetailsObject | add-member -membertype NoteProperty -Name "HD00" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD00"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD01" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD01"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD02" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD02"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD03" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD03"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD04" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD04"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD05" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD05"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD06" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD06"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD07" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD07"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD08" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD08"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "HD09" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD09"}).Type
			
			$TopLevelObject | add-member -membertype NoteProperty -Name "Details" -value $DetailsObject
		}else{
			$Message = "$(get-date) ComputerName $ComputerName has no Details in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		if(!$IPObjects -and !$PortObjects -and !$DetailsObject)
		{
			$Message = "$(get-date) $ComputerName not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}else{
			$ArrayOfRackTablesComputerObjects += $TopLevelObject
		}
	}
	Return $ArrayOfRackTablesComputerObjects
}

#=====================================================================
#Get-RTPDUDetails
#=====================================================================
Function Get-RTPDUDetails
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns all:
	IPs,
	Ports,
	Details
	associated to a PDUName
.DESCRIPTION

.EXAMPLE
	$TopLevelObject = Get-RTPDUDetails -PDUName RACK01PDU-OUT
	
	$TopLevelObject.PDUName
	
	SERVER01
	
	$TopLevelObject.IPs
	
	   IPinINT IP             NICName
	   ------- --             -------
	1231231231 10.0.0.10   		NIC1
	1231231231 10.0.0.1		 	iDrac
	1231231231 15.15.15.1		 NIC2
	
	$TopLevelObject.Ports
	
	PortName PortLabel InterfaceType MacAddress
	-------- --------- ------------- ----------
	idrac              1000Base-T    AA:AA:AA:AA:AA:AA
	NIC1               1000Base-T    AA:AA:AA:AA:AA:AA
	NIC2               1000Base-T    AA:AA:AA:AA:AA:AA
	NIC3               1000Base-T    AA:AA:AA:AA:AA:AA
	NIC4               1000Base-T    AA:AA:AA:AA:AA:AA

	$TopLevelObject.Details
	
	VisibleLabel : SERVER01
	SerialNumber : ASDASD
	HWType       : Dell PowerEdge%GPASS%R620
	OS           : Microsoft%GSKIP%Windows 7
	RAM          : 64GB
.EXAMPLE
	$ArrayOfRackTablesComputerObjects = Get-RTServerDetails -RackName RACK01
.NOTES
	
.LINK

#>
Param(
  [string]$RackName,
  [string]$PDUName,
  [int]$ObjectId
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Get-RTServerDetails.log"
		
	if($ObjectId)
	{
		# 2 = PDUs
		$NameQuery = Connect-RTToMysql -Query "Select name,asset_no from Object where id = $ObjectId AND objtype_id = 2"
		if($NameQuery)
		{
			$PDUName = $NameQuery.name
		}
		if(!$PDUName)
		{
			$Message = "$(get-date) PDUName not found for ObjectId $ObjectId in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}

	$AllPDUNamesInRack = @()
	if($RackName)
	{
		$AllPDUNamesInRack += Get-RTObjectNamesFromRack -RackName $RackName -PDUNames -ReturnObjects
		if($AllPDUNamesInRack.count -eq 0)
		{
			$Message = "$(get-date) RackName $RackName has no PDU Objects in Racktables"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
	}else{
		$Object = New-Object -TypeName System.Object
		Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $PDUName
		$AllPDUNamesInRack += $Object
	}

	$ArrayOfRackTablesComputerObjects = @()

	foreach($PDUNameObject in $AllPDUNamesInRack)
	{
		$PDUName = $Null
		$PDUName = $PDUNameObject.name
		$IdQuery = Connect-RTToMysql -Query "Select id,asset_no from Object where name = '$PDUName' AND objtype_id = 2"
		if($IdQuery)
		{
			$ObjectId = $IdQuery.id
			if($IdQuery.asset_no.ToString() -ne "")
			{
				$AssetTag = $IdQuery.asset_no
			}else{
				$AssetTag = $Null
			}
		}
		if(!$ObjectId)
		{
			$Message = "$(get-date) ObjectId for PDUName $PDUName not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
	
		$TopLevelObject = new-object object
		$TopLevelObject | add-member -membertype NoteProperty -Name "PDUName" -value $PDUName
		$TopLevelObject | add-member -membertype NoteProperty -Name "ObjectId" -value $ObjectId
		$TopLevelObject | add-member -membertype NoteProperty -Name "AssetTag" -value $AssetTag
		if($PDUNameObject.Rack)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Rack" -value $PDUNameObject.Rack
		}
		
		#----------------------------------------------------------------------
		# IPs
		#----------------------------------------------------------------------
		$IPSQLQuery = Connect-RTToMysql -Query "SELECT
												IPv4Allocation.ip as IPinINT,
												CAST(INET_NTOA(IPv4Allocation.ip) AS CHAR) as IP,
												IPv4Allocation.name as NICName,
												Object.name as PDUName
												FROM IPv4Allocation
												LEFT JOIN Object ON Object.id = object_id
												WHERE Object.name = '$PDUName'"
								
		$IPObjects = @()
		foreach($Row in $IPSQLQuery)
		{
			$Object = new-object object
			$Object | add-member -membertype NoteProperty -Name "IPinINT" -value $Row.IPinINT
			$Object | add-member -membertype NoteProperty -Name "IP" -value $Row.IP
			$Object | add-member -membertype NoteProperty -Name "NICName" -value $Row.NICName
			
			$IPObjects += $Object
		}
		if($IPObjects)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "IPs" -value $IPObjects
		}else{
				$Message = "$(get-date) PDUName $PDUName has no IPs in Racktables."
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		# Ports
		#----------------------------------------------------------------------
		#----------------------------------------------------------------------
		# Ports
		#----------------------------------------------------------------------
		# have to do it like this, I had a combined porta or portb query before, but it takes over 25 seconds to complete on the prod DB.
		$PortQuery = @()

		$PortQuery += Connect-RTToMysql -Query "SELECT
														Port.name,
														Port.label,
														Port.id,
														PortOuterInterface.oif_name,
														l2address,
														Link.portb,
														Link.porta
														FROM Port
														LEFT JOIN Link ON Link.portb = Port.id
														LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
														LEFT JOIN Object ON Object.id = object_id
														WHERE Object.name = '$PDUName'"
		$PortQuery += Connect-RTToMysql -Query "SELECT
														Port.name,
														Port.label,
														Port.id,
														PortOuterInterface.oif_name,
														l2address,
														Link.portb,
														Link.porta
														FROM Port
														LEFT JOIN Link ON Link.porta = Port.id
														LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
														LEFT JOIN Object ON Object.id = object_id
														WHERE Object.name = '$PDUName'"
		$ConnectedPorts = @()
		$ConnectedPorts += $PortQuery | where-object {!($_.porta -is [System.DBNull])}
		# now we have to filter out the duplicate ports returned in the query above.
		$FilteredPortQuery = @()
		foreach($Port in $PortQuery)
		{
			if($ConnectedPorts.Name -eq $Port.name)
			{
				# it's a connected port
				if($Port.porta -is [System.DBNull])
				{
					# but this is the query that doesn't have the porta/portb in it, so drop it
				}else{
					$FilteredPortQuery += $Port
				}
			}else{
				# oh, it's not connected, so add it just once
				$ConnectedPorts += $Port
				$FilteredPortQuery += $Port
			}
		}
		$PortObjects = @()
		foreach($Row in $FilteredPortQuery)
		{
			$Object = new-object object
			$Object | add-member -membertype NoteProperty -Name "PortName" -value $Row.name
			$Object | add-member -membertype NoteProperty -Name "PortLabel" -value $Row.label
			$Object | add-member -membertype NoteProperty -Name "InterfaceType" -value $Row.oif_name
			if($Row.l2address)
			{
				if(!($Row.l2address -is [System.DBNull]))
				{
					$Object | add-member -membertype NoteProperty -Name "MacAddress" -value ($Row.l2address.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")).Trim()
				}
			}
			if(!($Row.porta -is [System.DBNull]))
			{
				if($Row.porta -eq $Row.id)
				{
					$RemotePortID = $Row.portb
				}else{
					$RemotePortID = $Row.porta
				}
				$RemotePortQuery = Connect-RTToMysql -Query "SELECT
																Port.name as PortName,
																Port.label,
																PortOuterInterface.oif_name,
																l2address,
																Link.portb,
																Link.porta,
																Object.name as ObjectName
																FROM Port
																LEFT JOIN Link ON (Link.portb = Port.id OR Link.porta = Port.id)
																LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
																LEFT JOIN Object ON Object.id = object_id
																WHERE Port.id = $RemotePortID"
				if(!$RemotePortQuery)
				{
					$Message = "$(get-date) something went very wrong. Could not find RemotePortQuery for $RemotePortID, $PDUName, $($Row.name)"
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					return
				}
				if($RemotePortQuery -is [array])
				{
					$Message = "$(get-date) something went very wrong. Received an array for RemotePortQuery for $RemotePortID, $SwitchName, $($Row.name)"
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					return
				}
				$Object | add-member -membertype NoteProperty -Name "RemotePortName" -value $RemotePortQuery.PortName
				$Object | add-member -membertype NoteProperty -Name "RemoteObjectName" -value $RemotePortQuery.ObjectName
				$Object | add-member -membertype NoteProperty -Name "Connected" -value $True
				if(!($RemotePortQuery.l2address -is [System.DBNull]))
				{
					if($RemotePortQuery.l2address.length -gt 0)
					{
						$Object | add-member -membertype NoteProperty -Name "RemoteMacAddress" -value ($RemotePortQuery.l2address.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")).Trim()
					}
				}
			}else{
				$Object | add-member -membertype NoteProperty -Name "Connected" -value $False
			}
			$PortObjects += $Object
		}
		if($PortObjects)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Ports" -value $PortObjects
		}else{
				$Message = "$(get-date) PDUName $PDUName has no Ports in Racktables."
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		# Details
		#----------------------------------------------------------------------
		$DetailQuery = Connect-RTToMysql -Query "Select
												Attribute.name as AttributeName,
												Dictionary.dict_value as Type,
												AttributeValue.string_value as string_value
												FROM Object
												LEFT JOIN AttributeValue ON Object.id = AttributeValue.object_id
												LEFT JOIN Attribute ON AttributeValue.attr_id = Attribute.id
												LEFT JOIN Dictionary ON Dictionary.dict_key = AttributeValue.uint_value
												WHERE Object.name = '$PDUName'"
		if($DetailQuery)
		{
			$DetailsObject = new-object object
			$LabelQuery = Connect-RTToMysql -Query "SELECT Object.label FROM Object	WHERE Object.name = '$PDUName'"
			
			$DetailsObject | add-member -membertype NoteProperty -Name "VisibleLabel" -value $LabelQuery.label
			$DetailsObject | add-member -membertype NoteProperty -Name "SerialNumber" -value ($DetailQuery | where-object {$_.AttributeName -eq "Service Tag/Serial Number"}).string_value
			$DetailsObject | add-member -membertype NoteProperty -Name "HWType" -value ($DetailQuery | where-object {$_.AttributeName -eq "HW type"}).Type
			$TopLevelObject | add-member -membertype NoteProperty -Name "Details" -value $DetailsObject
		}else{
			$Message = "$(get-date) PDUName $PDUName has no Details in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		if(!$IPObjects -and !$PortObjects -and !$DetailsObject)
		{
			$Message = "$(get-date) $PDUName not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}else{
			$ArrayOfRackTablesComputerObjects += $TopLevelObject
		}
	}
	Return $ArrayOfRackTablesComputerObjects
}

#=====================================================================
#Get-RTSwitchDetails
#=====================================================================
Function Get-RTSwitchDetails
{
<#
.SYNOPSIS
	
.DESCRIPTION

.EXAMPLE
	$SwitchObject = Get-RTSwitchDetails -SwitchName "RACK01S1"
.NOTES
	
.LINK

#>
Param(
  [string]$RackName,
  [string]$SwitchName,
  [int]$ObjectId
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Get-RTSwitchDetails.log"
		
	if($ObjectId)
	{
		$NameQuery = Connect-RTToMysql -Query "Select name,asset_no from Object where id = $ObjectId AND objtype_id = 8"
		if($NameQuery)
		{
			$SwitchName = $NameQuery.name
		}
		if(!$SwitchName)
		{
			$Message = "$(get-date) SwitchName not found for ObjectId $ObjectId in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}

	$AllSwitchNamesInRack = @()
	if($RackName)
	{
		$AllSwitchNamesInRack += Get-RTObjectNamesFromRack -RackName $RackName -SwitchNames -ReturnObjects
	}else{
		$Object = New-Object -TypeName System.Object
		Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $SwitchName
		$AllSwitchNamesInRack += $Object
	}

	$ArrayOfRackTablesSwitchObjects = @()

	foreach($SwitchNameObject in $AllSwitchNamesInRack)
	{
		$SwitchName = $Null
		$SwitchName = $SwitchNameObject.name
		$IdQuery = Connect-RTToMysql -Query "Select id,asset_no from Object where name = '$SwitchName' AND objtype_id = 8"
		if($IdQuery)
		{
			$ObjectId = $IdQuery.id
			if($IdQuery.asset_no.ToString() -ne "")
			{
				$AssetTag = $IdQuery.asset_no
			}else{
				$AssetTag = $Null
			}
		}
		if(!$ObjectId)
		{
			$Message = "$(get-date) ObjectId for SwitchName $SwitchName not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
	
		$TopLevelObject = new-object object
		$TopLevelObject | add-member -membertype NoteProperty -Name "SwitchName" -value $SwitchName
		$TopLevelObject | add-member -membertype NoteProperty -Name "ObjectId" -value $ObjectId
		$TopLevelObject | add-member -membertype NoteProperty -Name "AssetTag" -value $AssetTag
		
		#----------------------------------------------------------------------
		# Ports
		#----------------------------------------------------------------------
		# have to do it like this, I had a combined porta or portb query before, but it takes over 25 seconds to complete on the prod DB.
		$PortQuery = @()

		$PortQuery += Connect-RTToMysql -Query "SELECT
														Port.name,
														Port.label,
														Port.id,
														PortOuterInterface.oif_name,
														l2address,
														Link.portb,
														Link.porta
														FROM Port
														LEFT JOIN Link ON Link.portb = Port.id
														LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
														LEFT JOIN Object ON Object.id = object_id
														WHERE Object.name = '$SwitchName'"
		$PortQuery += Connect-RTToMysql -Query "SELECT
														Port.name,
														Port.label,
														Port.id,
														PortOuterInterface.oif_name,
														l2address,
														Link.portb,
														Link.porta
														FROM Port
														LEFT JOIN Link ON Link.porta = Port.id
														LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
														LEFT JOIN Object ON Object.id = object_id
														WHERE Object.name = '$SwitchName'"
		$ConnectedPorts = @()
		$ConnectedPorts += $PortQuery | where-object {!($_.porta -is [System.DBNull])}
		# now we have to filter out the duplicate ports returned in the query above.
		$FilteredPortQuery = @()
		foreach($Port in $PortQuery)
		{
			if($ConnectedPorts.Name -eq $Port.name)
			{
				# it's a connected port
				if($Port.porta -is [System.DBNull])
				{
					# but this is the query that doesn't have the porta/portb in it, so drop it
				}else{
					$FilteredPortQuery += $Port
				}
			}else{
				# oh, it's not connected, so add it just once
				$ConnectedPorts += $Port
				$FilteredPortQuery += $Port
			}
		}
		$PortObjects = @()
		foreach($Row in $FilteredPortQuery)
		{
			$Object = new-object object
			$Object | add-member -membertype NoteProperty -Name "PortName" -value $Row.name
			$Object | add-member -membertype NoteProperty -Name "PortLabel" -value $Row.label
			$Object | add-member -membertype NoteProperty -Name "InterfaceType" -value $Row.oif_name
			if($Row.l2address)
			{
				if(!($Row.l2address -is [System.DBNull]))
				{
					$Object | add-member -membertype NoteProperty -Name "MacAddress" -value ($Row.l2address.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")).Trim()
				}
			}
			if(!($Row.porta -is [System.DBNull]))
			{
				if($Row.porta -eq $Row.id)
				{
					$RemotePortID = $Row.portb
				}else{
					$RemotePortID = $Row.porta
				}
				# migth be able to improve performance by removing that OR statement from the link join.
				$RemotePortQuery = Connect-RTToMysql -Query "SELECT
																Port.name as PortName,
																Port.label,
																PortOuterInterface.oif_name,
																l2address,
																Link.portb,
																Link.porta,
																Object.name as ObjectName
																FROM Port
																LEFT JOIN Link ON (Link.portb = Port.id OR Link.porta = Port.id)
																LEFT JOIN PortOuterInterface ON PortOuterInterface.id = type
																LEFT JOIN Object ON Object.id = object_id
																WHERE Port.id = $RemotePortID"
				if(!$RemotePortQuery)
				{
					$Message = "$(get-date) something went very wrong. Could not find RemotePortQuery for $RemotePortID, $SwitchName, $($Row.name)"
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					return
				}
				if($RemotePortQuery -is [array])
				{
					$Message = "$(get-date) something went very wrong. Received an array for RemotePortQuery for $RemotePortID, $SwitchName, $($Row.name)"
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					return
				}
				$Object | add-member -membertype NoteProperty -Name "RemotePortName" -value $RemotePortQuery.PortName
				$Object | add-member -membertype NoteProperty -Name "RemoteObjectName" -value $RemotePortQuery.ObjectName
				$Object | add-member -membertype NoteProperty -Name "Connected" -value $True
				if(!($RemotePortQuery.l2address -is [System.DBNull]))
				{
					$Object | add-member -membertype NoteProperty -Name "RemoteMacAddress" -value ($RemotePortQuery.l2address.insert(2,":").insert(5,":").insert(8,":").insert(11,":").insert(14,":")).Trim()
				}
			}else{
				$Object | add-member -membertype NoteProperty -Name "Connected" -value $False
			}
			$PortObjects += $Object
		}
		if($PortObjects)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Ports" -value $PortObjects
		}else{
				$Message = "$(get-date) SwitchName $SwitchName has no Ports in Racktables."
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		# Details
		#----------------------------------------------------------------------
		$DetailQuery = Connect-RTToMysql -Query "Select
												Attribute.name as AttributeName,
												Dictionary.dict_value as Type,
												AttributeValue.string_value as string_value
												FROM Object
												LEFT JOIN AttributeValue ON Object.id = AttributeValue.object_id
												LEFT JOIN Attribute ON AttributeValue.attr_id = Attribute.id
												LEFT JOIN Dictionary ON Dictionary.dict_key = AttributeValue.uint_value
												WHERE Object.name = '$SwitchName'"
		if($DetailQuery)
		{
			$DetailsObject = new-object object
			$LabelQuery = Connect-RTToMysql -Query "SELECT Object.label FROM Object WHERE Object.name = '$SwitchName'"
			
			$DetailsObject | add-member -membertype NoteProperty -Name "VisibleLabel" -value $LabelQuery.label
			$DetailsObject | add-member -membertype NoteProperty -Name "Service Tag" -value ($DetailQuery | where-object {$_.AttributeName -eq "Service Tag/Serial Number"}).string_value
			$DetailsObject | add-member -membertype NoteProperty -Name "SerialNumber" -value ($DetailQuery | where-object {$_.AttributeName -eq "Serial Number"}).string_value
			$DetailsObject | add-member -membertype NoteProperty -Name "HWType" -value ($DetailQuery | where-object {$_.AttributeName -eq "HW type"}).Type
			$TopLevelObject | add-member -membertype NoteProperty -Name "Details" -value $DetailsObject
		}else{
			$Message = "$(get-date) SwitchName $SwitchName has no Details in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		if(!$IPObjects -and !$PortObjects -and !$DetailsObject)
		{
			$Message = "$(get-date) $SwitchName not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}else{
			$ArrayOfRackTablesSwitchObjects += $TopLevelObject
		}
		if($SwitchNameObject.Rack)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Rack" -value $SwitchNameObject.Rack
		}
	}
	Return $ArrayOfRackTablesSwitchObjects
}

#=====================================================================
#Get-RTChassisDetails
#=====================================================================
Function Get-RTChassisDetails
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns all Details associated to a Chassis
.DESCRIPTION

.EXAMPLE
	$TopLevelObject = Get-RTChassisDetails -ChassisName "CHASSIS01"
	$TopLevelObject.ChassisName
	$TopLevelObject.Details
	
	VisibleLabel              : CHASSIS01
	AssetTag                  :
	Service Tag/Serial Number :
	Serial Number             :
	HWType                    :

.EXAMPLE

Get-RTChassisDetails -RackName RACK01

.NOTES
	
.LINK

#>
Param(
	[string]$RackName,
	[string]$ChassisName,
	[int]$ObjectId
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Get-RTChassisDetails.log"
	$Message = "$(get-date) Get-RTChassisDetails -ChassisName $ChassisName -ObjectId $ObjectId"
	$Message | Out-MyLogFile $logfile -append

	if($ObjectId)
	{
		$NameQuery = Connect-RTToMysql -Query "Select name,asset_no from Object where id = $ObjectId AND objtype_id = 1502"
		if($NameQuery)
		{
			$ChassisName = $NameQuery.name
		}
		if(!$ChassisName)
		{
			$Message = "$(get-date) ObjectId $ObjectId not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}
	
	$AllChassisNamesInRack = @()
	if($RackName)
	{
		$AllChassisNamesInRack +=  Get-RTObjectNamesFromRack -RackName $RackName -ChassisNames -ReturnObjects
	}else{
		$Object = New-Object -TypeName System.Object
		Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $ChassisName
		$AllChassisNamesInRack += $Object
	}

	$ArrayOfRackTablesChassisObjects = @()

	foreach($ChassisNameObject in $AllChassisNamesInRack)
	{
		$ChassisName = $Null
		$ChassisName = $ChassisNameObject.name
		$IdQuery = Connect-RTToMysql -Query "Select id,asset_no from Object where name = '$ChassisName' AND objtype_id = 1502"
		if($IdQuery)
		{
			$ObjectId = $IdQuery.id
			if($IdQuery.asset_no.ToString() -ne "")
			{
				$AssetTag = $IdQuery.asset_no
			}else{
				$AssetTag = $Null
			}
		}
		
		if(!$ObjectId)
		{
			$Message = "$(get-date) ObjectId not found in Racktables for ChassisName $ChassisName"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
		
		$TopLevelObject = new-object object
		$TopLevelObject | add-member -membertype NoteProperty -Name "ChassisName" -value $ChassisName
		$TopLevelObject | add-member -membertype NoteProperty -Name "ObjectId" -value $ObjectId
		$TopLevelObject | add-member -membertype NoteProperty -Name "AssetTag" -value $AssetTag
		if($ChassisNameObject.Rack)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Rack" -value $ChassisNameObject.Rack
		}
		if($ChassisNameObject.U)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "U" -value $ChassisNameObject.U
		}
		if($ChassisNameObject.Enclosure)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Enclosure" -value $ChassisNameObject.Enclosure
		}
		#----------------------------------------------------------------------
		# Details
		#----------------------------------------------------------------------
		$DetailQuery = Connect-RTToMysql -Query "Select
												Attribute.name as AttributeName,
												Dictionary.dict_value as Type,
												AttributeValue.string_value as string_value
												FROM Object
												LEFT JOIN AttributeValue ON Object.id = AttributeValue.object_id
												LEFT JOIN Attribute ON AttributeValue.attr_id = Attribute.id
												LEFT JOIN Dictionary ON Dictionary.dict_key = AttributeValue.uint_value
												WHERE Object.name = '$ChassisName' AND Object.objtype_id = 1502"
		if($DetailQuery)
		{
			$DetailsObject = new-object object
			$LabelQuery = Connect-RTToMysql -Query "SELECT Object.label, Object.asset_no FROM Object WHERE Object.name = '$ChassisName'"
			# why are there two serial numbers? And some disk arrays have the serial where the asset tag goes..
			$DetailsObject | add-member -membertype NoteProperty -Name "VisibleLabel" -value $LabelQuery.label
			$DetailsObject | add-member -membertype NoteProperty -Name "AssetTag" -value $LabelQuery.asset_no
			$DetailsObject | add-member -membertype NoteProperty -Name "SerialNumber" -value ($DetailQuery | where-object {$_.AttributeName -eq "Service Tag/Serial Number"}).string_value
			$DetailsObject | add-member -membertype NoteProperty -Name "HWType" -value ($DetailQuery | where-object {$_.AttributeName -eq "HW type"}).Type
			
			$TopLevelObject | add-member -membertype NoteProperty -Name "Details" -value $DetailsObject
		}else{
			$Message = "$(get-date) ChassisName $ChassisName has no Details in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		# BladeNames
		#----------------------------------------------------------------------
		$BladeQuery = Connect-RTToMysql -Query "select child_entity_id from Object
												LEFT JOIN EntityLink ON Object.id = EntityLink.parent_entity_id
												 where objtype_id = 1502 and name = '$ChassisName'"
		if($BladeQuery)
		{
			$ComputerNames = @()
			foreach($id in $BladeQuery.child_entity_id)
			{
				$SQLQueryBlades = Connect-RTToMysql -Query "select Object.name FROM Object
														Where Object.id = '$id'"
				$ComputerNames += $SQLQueryBlades.name
			}
			$TopLevelObject | add-member -membertype NoteProperty -Name "ComputerNames" -value $ComputerNames
		}else{
			$Message = "$(get-date) ChassisName $ChassisName has no blades in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
		$ArrayOfRackTablesChassisObjects += $TopLevelObject
	}
	Return $ArrayOfRackTablesChassisObjects
}

#=====================================================================
#Get-RTDiskArrayDetails
#=====================================================================
Function Get-RTDiskArrayDetails
{
<#
.SYNOPSIS
	Connects to the Racktables MySQL DB and returns all Details associated to a DiskArray
.DESCRIPTION

.EXAMPLE
	$TopLevelObject = Get-RTDiskArrayDetails -ObjectId 4054
	$TopLevelObject.DiskArrayName
	$TopLevelObject.Details
	
	VisibleLabel              :
	AssetTag                  : ASSETTAG01
	Service Tag/Serial Number :
	Serial Number             : ASDASD
	HWType                    :
	Controller                :
.EXAMPLE
	$TopLevelObject = Get-RTDiskArrayDetails -RackName RACK01
.NOTES
	
.LINK

#>
Param(
	[string]$RackName,
	[string]$DiskArrayName,
	[int]$ObjectId
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Get-RTDiskArrayDetails.log"
	$Message = "$(get-date) Get-RTDiskArrayDetails -DiskArrayName $DiskArrayName -ObjectId $ObjectId"
	$Message | Out-MyLogFile $logfile -append

	if($ObjectId)
	{
		$NameQuery = Connect-RTToMysql -Query "Select name,asset_no from Object where id = $ObjectId AND objtype_id = 5"
		if($NameQuery)
		{
			$DiskArrayName = $NameQuery.name
		}
		if(!$DiskArrayName)
		{
			$Message = "$(get-date) ObjectId $ObjectId not found in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}

	$AllDiskArrayNamesInRack = @()
	if($RackName)
	{
		$AllDiskArrayNamesInRack += Get-RTObjectNamesFromRack -RackName $RackName -DiskArrayNames -ReturnObjects
	}else{
		$Object = New-Object -TypeName System.Object
		Add-Member -inputObject $Object -MemberType NoteProperty -Name "Name" -Value $DiskArrayName
		$AllDiskArrayNamesInRack += $Object
	}

	$ArrayOfRackTablesDiskArrayObjects = @()

	foreach($DiskArrayObject in $AllDiskArrayNamesInRack)
	{
		$DiskArrayName = $Null
		$DiskArrayName = $DiskArrayObject.Name
		$IdQuery = Connect-RTToMysql -Query "Select id,asset_no from Object where name = '$DiskArrayName' AND objtype_id = 5"
		if($IdQuery)
		{
			$ObjectId = $IdQuery.id
			if($IdQuery.asset_no.ToString() -ne "")
			{
				$AssetTag = $IdQuery.asset_no
			}else{
				$AssetTag = $Null
			}
		}
		if(!$ObjectId)
		{
			$Message = "$(get-date) ObjectId not found in Racktables for DiskArrayName $DiskArrayName"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
		
		$TopLevelObject = new-object object
		$TopLevelObject | add-member -membertype NoteProperty -Name "DiskArrayName" -value $DiskArrayName
		$TopLevelObject | add-member -membertype NoteProperty -Name "ObjectId" -value $ObjectId
		$TopLevelObject | add-member -membertype NoteProperty -Name "AssetTag" -value $AssetTag

		$AttachedToQuery = Connect-RTToMysql -Query "select Object.name from Object
													LEFT JOIN EntityLink ON EntityLink.parent_entity_id = Object.id
													where EntityLink.child_entity_id = $ObjectId AND objtype_id = 4"
		if($AttachedToQuery)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "AttachedToServer" -value $AttachedToQuery.name
		}
		#----------------------------------------------------------------------
		# Details
		#----------------------------------------------------------------------
		$DetailQuery = Connect-RTToMysql -Query "Select
												Attribute.name as AttributeName,
												Dictionary.dict_value as Type,
												AttributeValue.string_value as string_value
												FROM Object
												LEFT JOIN AttributeValue ON Object.id = AttributeValue.object_id
												LEFT JOIN Attribute ON AttributeValue.attr_id = Attribute.id
												LEFT JOIN Dictionary ON Dictionary.dict_key = AttributeValue.uint_value
												WHERE Object.name = '$DiskArrayName' AND Object.objtype_id = 5"
		if($DetailQuery)
		{
			$DetailsObject = new-object object
			$LabelQuery = Connect-RTToMysql -Query "SELECT Object.label, Object.asset_no FROM Object WHERE Object.name = '$DiskArrayName'"
			# why are there two serial numbers? And some disk arrays have the serial where the asset tag goes..
			$DetailsObject | add-member -membertype NoteProperty -Name "VisibleLabel" -value $LabelQuery.label
			$DetailsObject | add-member -membertype NoteProperty -Name "AssetTag" -value $LabelQuery.asset_no
			$DetailsObject | add-member -membertype NoteProperty -Name "SerialNumber" -value ($DetailQuery | where-object {$_.AttributeName -eq "Service Tag/Serial Number"}).string_value
			$DetailsObject | add-member -membertype NoteProperty -Name "HWType" -value ($DetailQuery | where-object {$_.AttributeName -eq "HW type"}).Type
			$DetailsObject | add-member -membertype NoteProperty -Name "Controller" -value ($DetailQuery | where-object {$_.AttributeName -eq "Controller"}).Type
			# hard disks.
			$i = 0
			do{
				$LeadingZero = $i.tostring("00")
				$DetailsObject | add-member -membertype NoteProperty -Name "HD$LeadingZero" -value ($DetailQuery | where-object {$_.AttributeName -eq "HD$LeadingZero"}).Type
				$i++
			}until($i -gt 24) # there is no 24 though. 24 = 25th disk. There are 24 max, not 25, whoever added that to racktables can't count.
			
			$TopLevelObject | add-member -membertype NoteProperty -Name "Details" -value $DetailsObject
		}else{
			$Message = "$(get-date) DiskArrayName $DiskArrayName has no Details in Racktables."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
		#----------------------------------------------------------------------
		
		if($DiskArrayObject.Rack)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "Rack" -value $DiskArrayObject.Rack
		}
		if($DiskArrayObject.U)
		{
			$TopLevelObject | add-member -membertype NoteProperty -Name "U" -value $DiskArrayObject.U
		}
		 $ArrayOfRackTablesDiskArrayObjects += $TopLevelObject
	}
	Return $ArrayOfRackTablesDiskArrayObjects
}

#=====================================================================
#Update-RTComputerName
#=====================================================================
Function Update-RTComputerName
{
<#
.SYNOPSIS
	Changes the Common name in Racktables (and adds the old name to Visible label if visible label is empty).
.DESCRIPTION

.EXAMPLE
	Update-RTComputerName -OldComputerName SERVER01 -NewComputerName SERVER02

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$OldComputerName,
  [Parameter(Position=1, Mandatory=$True)][string]$NewComputerName
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTComputerName.log"


	$NewComputerName = $NewComputerName.ToUpper()
	$SQLQueryNewName = Connect-RTToMysql -Query "SELECT * FROM Object
										WHERE Object.name = '$NewComputerName'"
	if($SQLQueryNewName)
	{
		$Message = "$(get-date) $OldComputerName - ERROR! NewComputerName $NewComputerName already exists! - id $($SQLQueryNewName.id)"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
	}
	$SQLQueryOldName = Connect-RTToMysql -Query "SELECT * FROM Object
										WHERE Object.name = '$OldComputerName'"
	$LabelExists = $False
	if(!($SQLQueryOldName))
	{
		$Message = "$(get-date) $OldComputerName - ERROR! OldComputerName $OldComputerName Not Found in Racktables!"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
	}else{
		if(!($SQLQueryOldName.label -is [System.DBNull]))
		{
			$LabelExists = $True
		}
	}
	if(!($LabelExists))
	{
		$Message = "$(get-date) $OldComputerName - Object.label is empty, updating it to $OldComputerName"
		write-host $Message -fore cyan
		$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "UPDATE Object
									SET label = '$OldComputerName'
									WHERE name = '$OldComputerName';"
	}else{
		$Message = "$(get-date) $OldComputerName - Object.label is already set, not updating - id $($SQLQueryOldName.id)"
		write-host $Message -fore cyan
		$Message | Out-MyLogFile $logfile -append
	}
	
	
	if(!($SQLQueryNewName) -and ($SQLQueryOldName))
	{
		$Message = "$(get-date) $OldComputerName - Updating Object.name to $NewComputerName - id $($SQLQueryOldName.id)"
		write-host $Message -fore cyan
		$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "UPDATE Object
									SET name = '$NewComputerName'
									WHERE name = '$OldComputerName';"
		Update-RTHistory -ObjectName $NewComputerName -comment "Updating Object.name to $NewComputerName"
	}
}

#=====================================================================
#Update-RTRAM
#=====================================================================
Function Update-RTRAM
{
<#
.SYNOPSIS
	Either updates or inserts the RAM value in racktables.
	Does require that the RAM value is already entered in the dictionary table. 
	
	select * from Dictionary
	WHERE chapter_id = 10001
	
	Which is edited here:
	
	https://racktablesURL/index.php?page=chapter&chapter_no=10001

.DESCRIPTION

.EXAMPLE
	Update-RTRAM -ComputerName SERVER01 -RAMAmountInGB 64

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ComputerName,
  [Parameter(Position=1, Mandatory=$True)][int]$RAMAmountInGB
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTRAM.log"

	$SQLQueryRAM = Connect-RTToMysql -Query "select
											Dictionary.chapter_id,
											Dictionary.dict_key,
											Dictionary.dict_value
											from Dictionary
											LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
											WHERE Chapter.name = 'RAM'"
	if(!$SQLQueryRAM)						
	{
		$Message = "$(get-date) $ComputerName - SQLQueryRAM is empty? - return."
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$SQLRow = $SQLQueryRAM | where-object {$_.dict_value -eq "$($RAMAmountInGB)GB"}
	if(!$SQLRow)
	{
		$Message = "$(get-date) $ComputerName - RAMAmountInGB does not match any dict_value"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ComputerName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ComputerName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttributeMap = Connect-RTToMysql -Query "select * from AttributeMap
														Where objtype_id = $($SQLQueryObjectID.objtype_id)
														AND chapter_id = $($SQLQueryRAM[0].chapter_id)"
	if(!$SQLQueryAttributeMap)
	{
		$Message = "$(get-date) $ComputerName - objtype_id $($SQLQueryObjectID.objtype_id) and chapter_id $($SQLQueryRAM[0].chapter_id) not found in AttributeMap table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																WHERE object_id = $($SQLQueryObjectID.id)
																AND attr_id = $($SQLQueryAttributeMap.attr_id)"
	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck.uint_value -eq $SQLRow.dict_key))
		{
			$Message = "$(get-date) $ComputerName - Updating RAM in Racktables - uint_value $($SQLRow.dict_key) - dict_value $($SQLRow.dict_value)"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
			Connect-RTToMysql -Query "UPDATE AttributeValue
									SET uint_value = $($SQLRow.dict_key)
									WHERE attr_id = $($SQLQueryAttributeMap.attr_id)
									AND object_id = $($SQLQueryObjectID.id);"
			Update-RTHistory -ObjectName $ComputerName -comment "Updating RAM from uint_value $($SQLQueryAttributeValueCheck.uint_value) to $RAMAmountInGB"
		}else{
			$Message = "$(get-date) $ComputerName - RAMAmountInGB is identical to current RAM in racktables"
			write-host $Message -fore yellow
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}else{
		$Message = "$(get-date) $ComputerName - Inserting RAM in Racktables VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttributeMap.attr_id), $($SQLRow.dict_key))"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "INSERT INTO AttributeValue (object_id, object_tid, attr_id, uint_value)
									VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttributeMap.attr_id), $($SQLRow.dict_key));"
		Update-RTHistory -ObjectName $ComputerName -comment "Inserting RAM $RAMAmountInGB"
	}
}


#=====================================================================
#Update-RTHWType
#=====================================================================
Function Update-RTHWType
{
<#
.SYNOPSIS
	Either updates or inserts the HWType value in racktables.
	Does require that the HWType is already entered in the dictionary table. 
	
	select * from Dictionary
	WHERE chapter_id = 11
	
	Which is edited here:
	
	https://racktablesURL/index.php?page=chapter&chapter_no=11

.DESCRIPTION

.EXAMPLE
	Update-RTHWType -ObjectName SERVER01 -Model "Dell PowerEdge C6320"
.EXAMPLE
	Update-RTHWType -ObjectName CHASSIS01 -Model "Dell PowerEdge C6300 Enclosure"
.EXAMPLE
	Update-RTHWType -ObjectName DISKARRAY01 -Model "Dell PowerVault MD1420"
.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$True)][string]$Model
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTHWType.log"

	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ObjectName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ObjectName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}

	if(($SQLQueryObjectID.objtype_id -eq 50333) -or ($SQLQueryObjectID.objtype_id -eq 4))
	{
		$ChapterName = "server models"
	}elseif($SQLQueryObjectID.objtype_id -eq 5)
	{
		$ChapterName = "disk array models"
	}elseif($SQLQueryObjectID.objtype_id -eq 1502)
	{
		$ChapterName = "server chassis models"
	}

	$SQLQueryModels = Connect-RTToMysql -Query "select
											Dictionary.chapter_id,
											Dictionary.dict_key,
											Dictionary.dict_value
											from Dictionary
											LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
											WHERE Chapter.name = '$ChapterName'"
	if(!$SQLQueryModels)						
	{
		$Message = "$(get-date) $ObjectName - SQLQueryModels is empty? - return."
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$SQLRow = $SQLQueryModels | where-object {$_.dict_value.replace("%GPASS%"," ") -eq $Model}
	if(!$SQLRow)
	{
		$Message = "$(get-date) $ObjectName - SQLQueryModels $Model does not match any dict_value"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}else{
		# some smartass added duplicate server models...
		if($SQLRow -is [array])
		{
			$SQLRow = $SQLRow[0]
		}
	}
	
	$SQLQueryAttributeMap = Connect-RTToMysql -Query "select * from AttributeMap
														Where objtype_id = $($SQLQueryObjectID.objtype_id)
														AND chapter_id = $($SQLQueryModels[0].chapter_id)"
	if(!$SQLQueryAttributeMap)
	{
		$Message = "$(get-date) $ObjectName - objtype_id $($SQLQueryObjectID.objtype_id) and chapter_id $($SQLQueryModels[0].chapter_id) not found in AttributeMap table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																WHERE object_id = $($SQLQueryObjectID.id)
																AND attr_id = $($SQLQueryAttributeMap.attr_id)"
	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck.uint_value -eq $SQLRow.dict_key))
		{
			$Message = "$(get-date) $ObjectName - Updating HWType in Racktables - uint_value $($SQLRow.dict_key) - dict_value $($SQLRow.dict_value)"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
			Connect-RTToMysql -Query "UPDATE AttributeValue
									SET uint_value = $($SQLRow.dict_key)
									WHERE attr_id = $($SQLQueryAttributeMap.attr_id)
									AND object_id = $($SQLQueryObjectID.id);"
									
									
			$OldSQLQueryModel = ($SQLQueryModels | where-object {$_.dict_key -eq $SQLQueryAttributeValueCheck.uint_value}).dict_value.replace("%GPASS%"," ")
			Update-RTHistory -ObjectName $ObjectName -Comment "changed HWType from $OldSQLQueryModel ($($SQLQueryAttributeValueCheck.uint_value)) to $Model ($($SQLRow.dict_key))"
		}else{
			$Message = "$(get-date) $ObjectName - HWType $Model is identical to current HWType in racktables"
			write-host $Message -fore yellow
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}else{
		$Message = "$(get-date) $ObjectName - Inserting HWType $Model in Racktables VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttributeMap.attr_id), $($SQLRow.dict_key))"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "INSERT INTO AttributeValue (object_id, object_tid, attr_id, uint_value)
									VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttributeMap.attr_id), $($SQLRow.dict_key));"
		Update-RTHistory -ObjectName $ObjectName -Comment "Set HWType from blank to $Model"
	}
}

#=====================================================================
#Update-RTHistory
#=====================================================================
Function Update-RTHistory
{
<#
.SYNOPSIS
	Inserts a row in the ObjectHistory table for the Object (name) specified.
.DESCRIPTION
	Always inserts with the name "racktables-sync" (though this can be changed if needed).
.EXAMPLE
	Update-RTHistory -ObjectName SERVER01 -Comment "Test Insert 2"
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$True)][string]$Comment
)

	$SQLQueryObject = Connect-RTToMysql -Query "select * FROM Object WHERE  Object.name = '$ObjectName'"
	if(!$SQLQueryObject)
	{
		$Message = "$(get-date) Update-RTHistory - $ObjectName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	Connect-RTToMysql -Query "INSERT INTO ObjectHistory (id, name, label, objtype_id, asset_no, has_problems, comment, ctime, user_name)
									VALUES ($($SQLQueryObject.id), 
									'$ObjectName',
									'$($SQLQueryObject.label)',
									$($SQLQueryObject.objtype_id),
									'$($SQLQueryObject.asset_no)',
									'$($SQLQueryObject.has_problems)',
									'$Comment',
									NOW(),
									'racktables-sync');"
}

#=====================================================================
#Update-RTHardDisk
#=====================================================================
Function Update-RTHardDisk
{
<#
.SYNOPSIS
	Either updates or inserts the HWType value in racktables.
	Does require that the HWType is already entered in the dictionary table. 
	
	select * from Dictionary
	WHERE chapter_id = 10002
	
	Which is edited here:
	
	https://racktablesURL/index.php?page=chapter&tab=edit&chapter_no=10002

.DESCRIPTION

.EXAMPLE
	Update-RTHardDisk -ComputerName "SERVER01" -HardDisk "HD00" -NewHardDiskName "600GB 2.5FF SAS 10K" -OldHardDiskName "10TB SATA"

.EXAMPLE
	Update-RTHardDisk -ComputerName "SERVER01" -HardDisk "HD01" -NewHardDiskName "600GB 2.5FF SAS 10K"

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ComputerName,
  [Parameter(Position=1, Mandatory=$True)][string]$HardDisk,
  [Parameter(Position=2, Mandatory=$True)][string]$NewHardDiskName,
  [Parameter(Position=3)][string]$OldHardDiskName
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTHardDisk.log"

	$HardDriveTypes = Connect-RTToMysql -Query "select
											Dictionary.chapter_id,
											Dictionary.dict_key,
											Dictionary.dict_value
											from Dictionary
											LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
											WHERE Chapter.name = 'Hard Drive Types'"
	if(!$HardDriveTypes)						
	{
		$Message = "$(get-date) $ComputerName - HardDriveTypes Query is empty? - return."
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$SQLRow = $HardDriveTypes | where-object {$_.dict_value -eq $NewHardDiskName}
	if(!$SQLRow)
	{
		$Message = "$(get-date) $ComputerName - HardDriveType $NewHardDiskName does not match any dict_value"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}else{
		if($SQLRow -is [array])
		{
			$SQLRow = $SQLRow[0]
		}
	}
	
	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ComputerName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ComputerName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttribute = Connect-RTToMysql -Query "select * FROM Attribute WHERE name = '$HardDisk'"
	if(!$SQLQueryAttribute)
	{
		$Message = "$(get-date) $ComputerName - $HardDisk not found in Attribute table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																WHERE object_id = $($SQLQueryObjectID.id)
																AND attr_id = $($SQLQueryAttribute.id)"
	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck.uint_value -eq $SQLRow.dict_key))
		{
			$Message = "$(get-date) $ComputerName - Updating HardDisk $HardDisk in Racktables - uint_value $($SQLRow.dict_key) - dict_value $($SQLRow.dict_value)"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
			Connect-RTToMysql -Query "UPDATE AttributeValue
									SET uint_value = $($SQLRow.dict_key)
									WHERE attr_id = $($SQLQueryAttribute.id)
									AND object_id = $($SQLQueryObjectID.id);"
			Update-RTHistory -ObjectName $ComputerName -comment "Updated HardDisk from uint_value $($SQLQueryAttributeValueCheck.uint_value) to $NewHardDiskName"
		}else{
			$Message = "$(get-date) $ComputerName - HardDisk $NewHardDiskName is identical to current HardDisk in racktables"
			write-host $Message -fore yellow
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}else{
		$Message = "$(get-date) $ComputerName - Inserting HardDisk $NewHardDiskName in Racktables VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), $($SQLRow.dict_key))"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "INSERT INTO AttributeValue (object_id, object_tid, attr_id, uint_value)
									VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), $($SQLRow.dict_key));"
		Update-RTHistory -ObjectName $ComputerName -comment "Inserting HardDisk $NewHardDiskName"
	}
}

#=====================================================================
#Update-RTCPUType
#=====================================================================
Function Update-RTCPUType
{
<#
.SYNOPSIS
	Either updates or inserts a CPU in racktables.
	Does require that the CPU is already entered in the dictionary table. 
.DESCRIPTION

.EXAMPLE
	Update-RTCPUType -ComputerName "SERVER01" -NewCPUType "Intel(R) Xeon(R) CPU E5-2630 v3 @ 2.40GHz"
.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ComputerName,
  [Parameter(Position=1, Mandatory=$True)][string]$NewCPUType
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTCPUType.log"
	$Attribute = "CPU Type"

	$AttributeList = Connect-RTToMysql -Query "select
											Dictionary.chapter_id,
											Dictionary.dict_key,
											Dictionary.dict_value
											from Dictionary
											LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
											WHERE Chapter.name = 'CPU Type'"
	if(!$AttributeList)						
	{
		$Message = "$(get-date) $ComputerName - AttributeList Query is empty? - return."
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$SQLRow = $AttributeList | where-object {$_.dict_value.split(";")[0].Trim() -eq $NewCPUType}
	if(!$SQLRow)
	{
		$Message = "$(get-date) $ComputerName - NewCPUType $NewCPUType does not match any dict_value"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}else{
		if($SQLRow -is [array])
		{
			$SQLRow = $SQLRow[0]
		}
	}
	
	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ComputerName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ComputerName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttribute = Connect-RTToMysql -Query "select * FROM Attribute WHERE name = '$Attribute'"
	if(!$SQLQueryAttribute)
	{
		$Message = "$(get-date) $ComputerName - $Attribute not found in Attribute table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																WHERE object_id = $($SQLQueryObjectID.id)
																AND attr_id = $($SQLQueryAttribute.id)"
	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck.uint_value -eq $SQLRow.dict_key))
		{
			$Message = "$(get-date) $ComputerName - Updating Attribute $Attribute in Racktables - uint_value $($SQLRow.dict_key) - dict_value $($SQLRow.dict_value)"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
			Connect-RTToMysql -Query "UPDATE AttributeValue
									SET uint_value = $($SQLRow.dict_key)
									WHERE attr_id = $($SQLQueryAttribute.id)
									AND object_id = $($SQLQueryObjectID.id);"
			Update-RTHistory -ObjectName $ComputerName -comment "Updated CPUType from uint_value $($SQLQueryAttributeValueCheck.uint_value) to $NewCPUType"
		}else{
			$Message = "$(get-date) $ComputerName - Attribute $NewCPUType is identical to current Attribute in racktables"
			write-host $Message -fore yellow
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}else{
		$Message = "$(get-date) $ComputerName - Inserting Attribute $NewCPUType in Racktables VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), $($SQLRow.dict_key))"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "INSERT INTO AttributeValue (object_id, object_tid, attr_id, uint_value)
									VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), $($SQLRow.dict_key));"
		Update-RTHistory -ObjectName $ComputerName -comment "Inserted CPUType $NewCPUType"
	}
}

#=====================================================================
#Update-RTSerialNumber
#=====================================================================
Function Update-RTSerialNumber
{
<#
.SYNOPSIS
	Either updates or inserts a SerialNumber into racktables.

.DESCRIPTION

.EXAMPLE
	Update-RTSerialNumber -ComputerName $($Object.ComputerName) -NewSerialNumber "TEST123"

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ComputerName,
  [Parameter(Position=1, Mandatory=$True)][string]$NewSerialNumber
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTSerialNumber.log"

	$Attribute = "Service Tag/Serial Number"
	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ComputerName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ComputerName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}

	$SQLQueryAttribute = Connect-RTToMysql -Query "select * FROM Attribute WHERE name = '$Attribute'"
	if(!$SQLQueryAttribute)
	{
		$Message = "$(get-date) $ComputerName - $Attribute not found in Attribute table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}

	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																WHERE object_id = $($SQLQueryObjectID.id)
																AND attr_id = $($SQLQueryAttribute.id)"

	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck.string_value -eq $NewSerialNumber))
		{
			$Message = "$(get-date) $ComputerName - Updating SerialNumber in Racktables from $($SQLQueryAttributeValueCheck.string_value) to $NewSerialNumber"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
			Connect-RTToMysql -Query "UPDATE AttributeValue
									SET string_value = '$NewSerialNumber'
									WHERE attr_id = $($SQLQueryAttribute.id)
									AND object_id = $($SQLQueryObjectID.id);"
			Update-RTHistory -ObjectName $ComputerName -comment "Updated SerialNumber in Racktables from $($SQLQueryAttributeValueCheck.string_value) to $NewSerialNumber"
		}else{
			$Message = "$(get-date) $ComputerName - NewSerialNumber $NewSerialNumber is identical to current SerialNumber in racktables"
			write-host $Message -fore yellow
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}else{
		$Message = "$(get-date) $ComputerName - Inserting NewSerialNumber $NewSerialNumber in Racktables"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "INSERT INTO AttributeValue (object_id, object_tid, attr_id, string_value)
									VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), '$NewSerialNumber');"
		Update-RTHistory -ObjectName $ComputerName -comment "Inserted NewSerialNumber $NewSerialNumber for ComputerName $ComputerName"
	}
}

#=====================================================================
#Update-RTAttribute
#=====================================================================
Function Update-RTAttribute
{
<#
.SYNOPSIS
	Either updates or inserts attribute values in racktables.
	Does require that the attribute is already entered in the dictionary table. 

.DESCRIPTION

.EXAMPLE
	Update-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "CPU Count" -ChapterName "CPU Count" -NewAttribute $NewCPUCount

.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$True)][string]$Attribute,
  [Parameter(Position=2, Mandatory=$True)][string]$ChapterName,
  [Parameter(Position=3, Mandatory=$True)][string]$NewAttribute
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTAttribute.log"

	$AttributeList = Connect-RTToMysql -Query "select
											Dictionary.chapter_id,
											Dictionary.dict_key,
											Dictionary.dict_value
											from Dictionary
											LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
											WHERE Chapter.name = '$ChapterName'"
	if(!$AttributeList)						
	{
		$Message = "$(get-date) $ObjectName - AttributeList Query is empty? - return."
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$SQLRow = $AttributeList | where-object {$_.dict_value -eq $NewAttribute}
	if(!$SQLRow)
	{
		$Message = "$(get-date) $ObjectName - NewAttribute $NewAttribute does not match any dict_value"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}else{
		if($SQLRow -is [array])
		{
			$SQLRow = $SQLRow[0]
		}
	}
	
	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ObjectName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ObjectName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttribute = Connect-RTToMysql -Query "select * FROM Attribute WHERE name = '$Attribute'"
	if(!$SQLQueryAttribute)
	{
		$Message = "$(get-date) $ObjectName - $Attribute not found in Attribute table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																WHERE object_id = $($SQLQueryObjectID.id)
																AND attr_id = $($SQLQueryAttribute.id)"
	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck.uint_value -eq $SQLRow.dict_key))
		{
			$Message = "$(get-date) $ObjectName - Updating Attribute $Attribute in Racktables to $NewAttribute - uint_value $($SQLRow.dict_key) - dict_value $($SQLRow.dict_value)"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
			Connect-RTToMysql -Query "UPDATE AttributeValue
									SET uint_value = $($SQLRow.dict_key)
									WHERE attr_id = $($SQLQueryAttribute.id)
									AND object_id = $($SQLQueryObjectID.id);"
			Update-RTHistory -ObjectName $ObjectName -comment "Updated Attribute $Attribute in Racktables from uint_value $($SQLQueryAttributeValueCheck.uint_value) to $NewAttribute"
		}else{
			$Message = "$(get-date) $ObjectName - Attribute $NewAttribute is identical to current Attribute in racktables"
			write-host $Message -fore yellow
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}else{
		$Message = "$(get-date) $ObjectName - Inserting Attribute $Attribute with value $NewAttribute in Racktables VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), $($SQLRow.dict_key))"
			write-host $Message -fore cyan
			$Message | Out-MyLogFile $logfile -append
		Connect-RTToMysql -Query "INSERT INTO AttributeValue (object_id, object_tid, attr_id, uint_value)
									VALUES ($($SQLQueryObjectID.id), $($SQLQueryObjectID.objtype_id), $($SQLQueryAttribute.id), $($SQLRow.dict_key));"
									
		Update-RTHistory -ObjectName $ObjectName -comment "Inserting Attribute $Attribute with value $NewAttribute"
	}
}

#=====================================================================
#Remove-RTAttribute
#=====================================================================
Function Remove-RTAttribute
{
<#
.SYNOPSIS
	Removes a single attribute from a single computer object. (changes it into NOT SET in racktables).
.DESCRIPTION

.EXAMPLE
	Remove-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "CPU Count"
.EXAMPLE
	Remove-RTAttribute -ObjectName "SERVER01" -Attribute "Controller"
.NOTES
	
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$True)][string]$Attribute
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Remove-RTAttribute.log"

	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ObjectName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ObjectName - Object.id not found in Object table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}

	$SQLQueryAttribute = Connect-RTToMysql -Query "select * FROM Attribute WHERE name = '$Attribute'"
	if(!$SQLQueryAttribute)
	{
		$Message = "$(get-date) $ObjectName - $Attribute not found in Attribute table"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}

	$SQLQueryAttributeValueCheck = Connect-RTToMysql -Query "select * FROM AttributeValue
																	WHERE object_id = $($SQLQueryObjectID.id)
																	AND attr_id = $($SQLQueryAttribute.id)"
	if($SQLQueryAttributeValueCheck)
	{
		if(!($SQLQueryAttributeValueCheck -is [array]))
		{
			
				$Message = "$(get-date) $ObjectName - Found $Attribute (attr_id: $($SQLQueryAttribute.id)) set on Object.id $($SQLQueryObjectID.id), removing as requested."
				write-host $Message -fore magenta
				$Message | Out-MyLogFile $logfile -append
			
			$Null = Connect-RTToMysql -Query "DELETE FROM AttributeValue
																	WHERE object_id = $($SQLQueryObjectID.id)
																	AND attr_id = $($SQLQueryAttribute.id)"
																	
			Update-RTHistory -ObjectName $ObjectName -comment "Wiped Attribute $Attribute from ObjectName $ObjectName"
		}else{
			$Message = "$(get-date) $ObjectName - Found $Attribute (attr_id: $($SQLQueryAttribute.id)) set on Object.id $($SQLQueryObjectID.id) - But it is an array? I don't understand. - exit."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
		}
	}else{
		$Message = "$(get-date) $ObjectName - $Attribute (attr_id: $($SQLQueryAttribute.id)) not found on Object.id $($SQLQueryObjectID.id) - Can't remove what is not there."
		write-host $Message -fore yellow
		$Message | Out-MyLogFile $logfile -append
	}
}

#=====================================================================
#Add-RTPort
#=====================================================================
Function Add-RTPort
{
<#
.SYNOPSIS
	Adds a single port to Racktables
.DESCRIPTION
	Adds a single port to Racktables
	Will also update the Mac Address if the port already exists, but the Mac Address is missing.
.EXAMPLE
	Add-RTPort -ObjectName $($Object.ComputerName) -MACAddress $Alert.MACAddressFromiDRAC -PortName "iDRAC" -InterfaceType "1000Base-T"
.EXAMPLE
	Add-RTPort -ObjectName RACK01TEST -MACAddress "AA:AA:AA:AA:AA:AA" -PortName "iDRAC" -InterfaceType "1000Base-T"
.EXAMPLE
	Add-RTPort -ObjectName SWITCH01 -PortName "GE1/0/1" -InterfaceType "1000Base-T"
.NOTES
	Checks to see if the MAC Address isn't already in use elsewhere.
	Checks to make sure the port name is unique for that object.
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$False)][string]$MACAddress,
  [Parameter(Position=2, Mandatory=$True)][string]$PortName,
  [Parameter(Position=3, Mandatory=$True)][string]$InterfaceType
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Add-RTPort.log"

	$SQLQueryObjectID = Connect-RTToMysql -Query "select Object.id, Object.objtype_id FROM Object WHERE  Object.name = '$ObjectName'"
	if(!$SQLQueryObjectID)
	{
		$Message = "$(get-date) $ObjectName - Object.id not found in Object table for $ObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}else{
		if(!$MACAddress)
		{
			if(!($SQLQueryObjectID.objtype_id -eq 8))
			{
				$Message = "$(get-date) $ObjectName is not a switch, so it needs a MacAddress for it to be entered into the DB (well, it doesn't REALLY have to.. but I want it like that)"
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
				return
			}
		}
	}
	$TypeID = $Null
	$SQLQueryInterfaceType = Connect-RTToMysql -Query "select * from PortOuterInterface where oif_name = '$InterfaceType'"
	if(!$SQLQueryInterfaceType)
	{
		$Message = "$(get-date) $ObjectName - no InterfaceType named $InterfaceType in table PortOuterInterface?"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}else{
		$TypeID = $SQLQueryInterfaceType.id
	}
	
	if($MACAddress)
	{
		$MACAddress = ($MACAddress.replace(":","")).ToUpper().Trim()
		if($MACAddress.length -ne 12)
		{
			$Message = "$(get-date) $ObjectName - MACAddress $MACAddress is not 12 characters?"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
		$SQLQueryPortMacCheck = Connect-RTToMysql -Query "select * FROM Port WHERE l2address = '$($MACAddress)'"
	}
	if($SQLQueryPortMacCheck)
	{
		$Message = "$(get-date) $ObjectName - MACAddress $MACAddress - already in Port table?"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		Return
	}
	$SQLQueryPortNameCheck = Connect-RTToMysql -Query "select * FROM Port WHERE object_id = '$($SQLQueryObjectID.id)' and name = '$PortName'"
	if($SQLQueryPortNameCheck)
	{
		if($MACAddress)
		{
			if(!$SQLQueryPortMacCheck)
			{
				if($SQLQueryPortNameCheck.l2address -is [System.DBNull])
				{
					$Message = "$(get-date) $ObjectName already has a port named $PortName, but has no MAC address. Adding Mac Address: $MACAddress"
					write-host $Message -fore green
					$Message | Out-MyLogFile $logfile -append
					
					Connect-RTToMysql -Query "UPDATE Port
											SET l2address = '$($MACAddress)'
											WHERE id = $($SQLQueryPortNameCheck.id)
											AND object_id = $($SQLQueryObjectID.id);"
					Update-RTHistory -ObjectName $ObjectName -comment "Updated MACAddress from nothing to $MACAddress for portname $PortName with InterfaceType $InterfaceType"
					Return
				}
			}else{
				$Message = "$(get-date) $ObjectName already has a port named $PortName and it already has a Mac Address set."
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
				Return	
			}
		}
	}
	
	$Message = "$(get-date) $ObjectName - Inserting Port VALUES ($($SQLQueryObjectID.id), '$($PortName)', 1, $TypeID, '$($MACAddress)')"
	write-host $Message -fore cyan
	$Message | Out-MyLogFile $logfile -append
	Connect-RTToMysql -Query "INSERT INTO Port (object_id, name, iif_id, type, l2address)
								VALUES ($($SQLQueryObjectID.id), '$($PortName)', 1, $TypeID, '$($MACAddress)');"
	if($MACAddress)
	{
		Update-RTHistory -ObjectName $ObjectName -comment "added PortName $PortName with MACAddress $MACAddress and InterfaceType $InterfaceType"
	}else{
		# Too spammy for switches with all those ports they have.
		if(!($SQLQueryObjectID.objtype_id -eq 8))
		{
			Update-RTHistory -ObjectName $ObjectName -comment "added PortName $PortName with InterfaceType $InterfaceType"
		}
	}
}

#=====================================================================
#Update-RTRacktables
#=====================================================================
Function Update-RTRacktables
{
<#
.SYNOPSIS
	Takes the output of Compare-RacktablesiDRACWMI and uses the comments to update various attributes.
.DESCRIPTION

	Currently fixes the following errors:
	
	RacktablesToOSHostnameMismatch,
	RAMMissingRacktables,
	RackTablesHWTypeMissing,
	(outdated, it fixes a whole lot more)

.EXAMPLE
	Update-RTRacktables -ArrayOfObjects $ArrayOfObjects

.NOTES

.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$ArrayOfObjects
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Update-RTRacktables.log"
	
	$PauseForErrors = $Null
	# Making sure all memory values are in the DB already:
	$RAMMissingRacktables = $ArrayOfObjects.Alerts | where-object {($_.Name -eq "RAMMissingRacktables") -or ($_.Name -eq "RAMIncorrectRacktablesOS")}
	if($RAMMissingRacktables)
	{
		$OSRAMArray = $RAMMissingRacktables.OSRAM | sort-object | get-unique
		
		$SQLQueryRAM = Connect-RTToMysql -Query "select
											Dictionary.chapter_id,
											Dictionary.dict_key,
											Dictionary.dict_value
											from Dictionary
											LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
											WHERE Chapter.name = 'RAM'"
		if(!$SQLQueryRAM)						
		{
			$Message = "$(get-date) SQLQueryRAM is empty?"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			$PauseForErrors = $True
		}
		foreach($OSRAM in $OSRAMArray)
		{
			$SQLRow = $Null
			$SQLRow = $SQLQueryRAM | where-object {$_.dict_value -eq "$($OSRAM)GB"}
			if(!$SQLRow)
			{
				$newOSRam = $OSRAM + 1
				$SQLRow = $SQLQueryRAM | where-object {$_.dict_value -eq "$($newOSRam)GB"}
				if(!$SQLRow)
				{
					$Message = "$(get-date) OSRAM `"$OSRAM`" does not match any dict_value?"
					write-host "$(($RAMMissingRacktables | where-object {$_.OSRAM -eq $OSRAM})[0] | format-list | out-string)" -fore cyan
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					$PauseForErrors = $True
				}
			}
		}
	}
	# Making sure all server models are in the DB already:
	$RackTablesHWTypeMissing = $ArrayOfObjects.Alerts | where-object {(($_.Name -eq "RackTablesHWTypeMissing") -or ($_.Name -eq "HWTypeMismatchRacktablesOS"))}
	if($RackTablesHWTypeMissing)
	{
		$SQLQueryServerModels = Connect-RTToMysql -Query "select
												Dictionary.chapter_id,
												Dictionary.dict_key,
												Dictionary.dict_value
												from Dictionary
												LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
												WHERE Chapter.name = 'server models'"
		if(!$SQLQueryServerModels)						
		{
			$Message = "$(get-date) SQLQueryServerModels is empty?"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			$PauseForErrors = $True
		}
		$HWTypeOSArray = $RackTablesHWTypeMissing.HWTypeOS | sort-object | get-unique
		foreach($HWTypeOS in $HWTypeOSArray)
		{
			$SQLRow = $Null
			$SQLRow = $SQLQueryServerModels | where-object {$_.dict_value.replace("%GPASS%"," ") -eq $HWTypeOS}
			if(!$SQLRow)
			{
				$Message = "$(get-date) HWTypeOS `"$HWTypeOS`" does not match any dict_value"
				write-host "$(($RackTablesHWTypeMissing | where-object {$_.HWTypeOS -eq $HWTypeOS})[0] | format-list | out-string)" -fore cyan
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
				$PauseForErrors = $True
			}
		}
	}
	# Making sure all hard disks are in the DB already:
	$RacktablesHardDiskMismatching = $ArrayOfObjects.Alerts | where-object {(($_.Name -eq "RacktablesHardDiskMismatching") -or ($_.Name -eq "RacktablesHardDiskisMissing"))}
	if($RacktablesHardDiskMismatching)
	{
		$HardDriveTypes = Connect-RTToMysql -Query "select
												Dictionary.chapter_id,
												Dictionary.dict_key,
												Dictionary.dict_value
												from Dictionary
												LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
												WHERE Chapter.name = 'Hard Drive Types'"
		if(!$HardDriveTypes)						
		{
			$Message = "$(get-date) $ComputerName - HardDriveTypes Query is empty?"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			$PauseForErrors = $True
		}
		$iDRACHardDiskNameArray = $RacktablesHardDiskMismatching.iDRACHardDiskName | sort-object | get-unique
		foreach($iDRACHardDiskName in $iDRACHardDiskNameArray)
		{
			$SQLRow = $Null
			$SQLRow = $HardDriveTypes | where-object {$_.dict_value -eq $iDRACHardDiskName}
			if(!$SQLRow)
			{
				$Message = "$(get-date) HardDriveType `"$iDRACHardDiskName`" does not match any dict_value"
				write-host "$(($RacktablesHardDiskMismatching | where-object {$_.iDRACHardDiskName -eq $iDRACHardDiskName})[0] | format-list | out-string)" -fore cyan
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
				$PauseForErrors = $True
			}
		}
	}
	# Making sure all CPUs are in the DB already:
	$MismatchingCPUTypeInRacktables = $ArrayOfObjects.Alerts | where-object {(($_.Name -eq "MismatchingCPUTypeInRacktables") -or ($_.Name -eq "NoCPUTypeInRacktables"))}
	if($MismatchingCPUTypeInRacktables)
	{
		$CPUTypesRT = Connect-RTToMysql -Query "select
												Dictionary.chapter_id,
												Dictionary.dict_key,
												Dictionary.dict_value
												from Dictionary
												LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
												WHERE Chapter.name = 'CPU Type'"
		if(!$CPUTypesRT)						
		{
			$Message = "$(get-date) $ComputerName - CPUTypesRT Query is empty?"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			$PauseForErrors = $True
		}
		$CPUTypeArray = $MismatchingCPUTypeInRacktables.CPUType | sort-object | get-unique
		foreach($CPUType in $CPUTypeArray)
		{
			$SQLRow = $Null
			$SQLRow = $CPUTypesRT | where-object {$_.dict_value.split(";")[0].Trim() -eq $CPUType}
			if(!$SQLRow)
			{
				$Message = "$(get-date) CPUType `"$CPUType`" does not match any dict_value"
				write-host "$(($MismatchingCPUTypeInRacktables | where-object {$_.CPUType -eq $CPUType})[0] | format-list | out-string)" -fore cyan
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
				$PauseForErrors = $True
			}
		}
	}

	# Checking all standard attribute values.
	$Attributes = 	"Controller",
					"Controller2",
					"SW type",
					"CPU Count"

	$ChapterNames = "Expansion RAID Controllers",
					"Expansion RAID Controllers",
					"server OS type",
					"CPU Count"
	
	$AlertNames = 	("RacktablesControllerMismatch","RacktablesControllerMissing"),
					("RacktablesController2Mismatch", "RacktablesController2Missing"),
					("MismatchingSWTypeInRacktables","NoSWTypeInRacktables"),
					("NoCPUCountInRacktables","MismatchingCPUCountInRacktables")
					
	$PropertyNames = "ControllerName",
					"Controller2Name",
					"OSName",
					"CPUCount"
					
	$i = 0
	Foreach($Attribute in $Attributes)
	{	
		$InRacktablesAlerts = $ArrayOfObjects.Alerts | where-object {$AlertNames[$i] -eq $_.Name}
		
		if($InRacktablesAlerts)
		{
			$DictionaryEntries = Connect-RTToMysql -Query "select
													Dictionary.chapter_id,
													Dictionary.dict_key,
													Dictionary.dict_value
													from Dictionary
													LEFT JOIN Chapter ON Chapter.id = Dictionary.chapter_id
													WHERE Chapter.name = '$($ChapterNames[$i])'"
			if(!$DictionaryEntries)						
			{
				$Message = "$(get-date) $ComputerName - DictionaryEntries Query for chapter name $($ChapterNames[$i]) is empty?"
				write-host $Message -fore red
				$Message | Out-MyLogFile $logfile -append
				$PauseForErrors = $True
			}
			$PropertyValuesArray = $InRacktablesAlerts.($PropertyNames[$i]) | sort-object | get-unique
			foreach($PropertyValue in $PropertyValuesArray)
			{
				$SQLRow = $DictionaryEntries | where-object {$_.dict_value.Trim() -eq $PropertyValue}
				if(!$SQLRow)
				{
					$Message = "$(get-date) PropertyValue `"$PropertyValue`" does not match any dict_value in chapter name $($ChapterNames[$i])"
					write-host "$(($InRacktablesAlerts | where-object {$_.($PropertyNames[$i]) -eq $PropertyValue})[0] | format-list | out-string)" -fore cyan
					write-host $Message -fore red
					$Message | Out-MyLogFile $logfile -append
					$PauseForErrors = $True
				}
			}
		}
		$i++
	}
	
	if($PauseForErrors)
	{
		# insterting a pause (so we can exit out of the code with ctrl-c), because we have dictionary values that are missing from the chapters.
		Start-Pause
		start-sleep 5
	}
	
	Foreach($Object in $ArrayOfObjects)
	{
		if(!($Object.ComputerName))
		{
			$Message = "$(get-date) $Object.ComputerName is empty?"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
		Foreach($Alert in $Object.Alerts)
		{
			# Hard disks.
			if($Alert.Name -eq "RacktablesHardDiskMismatching")
			{
				Update-RTHardDisk -ComputerName $($Object.ComputerName) -HardDisk "HD0$($Alert.HardDiskNumber)" -NewHardDiskName $Alert.iDRACHardDiskName -OldHardDiskName $Alert.RackTablesHardDiskName
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if($Alert.Name -eq "RacktablesHardDiskisMissing")
			{
				Update-RTHardDisk -ComputerName $($Object.ComputerName) -HardDisk "HD0$($Alert.HardDiskNumber)" -NewHardDiskName $Alert.iDRACHardDiskName
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if($Alert.Name -eq "RacktablesHardDiskNotIniDRAC")
			{
				Remove-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "HD0$($Alert.HardDiskNumber)"
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			
			# Controllers
			if(($Alert.Name -eq "RacktablesControllerMismatch") -or ($Alert.Name -eq "RacktablesControllerMissing"))
			{
				Update-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "Controller" -ChapterName "Expansion RAID Controllers" -NewAttribute $Alert.ControllerName
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if(($Alert.Name -eq "RacktablesController2Mismatch") -or ($Alert.Name -eq "RacktablesController2Missing"))
			{
				Update-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "Controller2" -ChapterName "Expansion RAID Controllers" -NewAttribute $Alert.Controller2Name
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if($Alert.Name -eq "RacktablesControllerNotFound")
			{
				Remove-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "Controller"
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if($Alert.Name -eq "RacktablesController2NotFound")
			{
				Remove-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "Controller2"
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			
			# SWTypes
			if(($Alert.Name -eq "MismatchingSWTypeInRacktables") -or ($Alert.Name -eq "NoSWTypeInRacktables"))
			{
				Update-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "SW type" -ChapterName "server OS type" -NewAttribute $Alert.OSName
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}

			# CPUs
			if(($Alert.Name -eq "MismatchingCPUTypeInRacktables") -or ($Alert.Name -eq "NoCPUTypeInRacktables"))
			{
				Update-RTCPUType -ComputerName $($Object.ComputerName) -NewCPUType $Alert.CPUType
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if(($Alert.Name -eq "NoCPUCountInRacktables") -or ($Alert.Name -eq "MismatchingCPUCountInRacktables"))
			{
				Update-RTAttribute -ObjectName $($Object.ComputerName) -Attribute "CPU Count" -ChapterName "CPU Count" -NewAttribute $Alert.CPUCount
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}

			# Ram updates
			if(($Alert.Name -eq "RAMMissingRacktables") -or ($Alert.Name -eq "RAMIncorrectRacktablesOS"))
			{
				Update-RTRAM -ComputerName $($Object.ComputerName) -RAMAmountInGB $Alert.OSRAM
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			# Hardware type updates
			if(($Alert.Name -eq "RackTablesHWTypeMissing") -or ($Alert.Name -eq "HWTypeMismatchRacktablesOS"))
			{
				Update-RTHWType -ObjectName $($Object.ComputerName) -Model $Alert.HWTypeOS
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if($Alert.Name -eq "NoSerialNumberFoundInRacktables")
			{
				Update-RTSerialNumber -ComputerName $($Object.ComputerName) -NewSerialNumber $Alert.Win32SystemEnclosureSerialNumber
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
			if(($Alert.Name -eq "RacktablesiDRACPORTisMissing") -and ($Alert.MACAddressFromiDRAC))
			{
				Add-RTPort -ObjectName $($Object.ComputerName) -MACAddress $Alert.MACAddressFromiDRAC -PortName "iDRAC" -InterfaceType "1000Base-T"
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
		}
		# Computername changes
		Foreach($Alert in $AllAlerts)
		{
			# This alert always needs to be done last, since it changes the ComputerName, which is used by all the other functions.
			if($Alert.Name -eq "RacktablesToOSHostnameMismatch")
			{
				Update-RTComputerName -OldComputerName $Alert.RackTablesComputerName -NewComputerName $Alert.Win32ComputersystemName
				$Alert | add-member -membertype NoteProperty -Name "Fixed" -value $True
			}
		}
	}
}


#=====================================================================
#Add-RTObject
#=====================================================================
Function Add-RTObject
{
<#
.SYNOPSIS
	Adds new objects to Racktables
.DESCRIPTION

.EXAMPLE
	Add-RTObject -ObjectName "RACK01-PDU" -RackName "RACK01" -ObjectType "PDU" -Front
.EXAMPLE
	Add-RTObject -ObjectName "RACK01-PDU" -RackName "RACK01" -ObjectType "PDU" -Front
.EXAMPLE
	Add-RTObject -ObjectName "RACK01-PDU" -RackName "RACK01" -ObjectType "PDU" -Front -TopHalf
.EXAMPLE
	Add-RTObject -ObjectName "SERVER01" -ParentName "CHASSIS01" -ObjectType "Dell C-Series"
.EXAMPLE
	$ComputerNames = "SERVER01","SERVER02","SERVER03","SERVER04"
	$IPs = "10.0.0.1", "10.0.0.2", "10.0.0.3", "10.0.0.4"
	$iDRACMacs = "AA:AA:AA:AA:AA:AA","BB:BB:BB:BB:BB:BB","CC:CC:CC:CC:CC:CC","DD:DD:DD:DD:DD:DD"
	$SerialNumbers = "ASDASD1","ASDASD2","ASDASD3","ASDASD4"

	$i = 0
	Foreach($ComputerName in $ComputerNames)
	{
		Add-RTObject -ObjectName $ComputerName -ParentName "CHASSIS01" -ObjectType "Dell C-Series"
		Add-RTPort -ObjectName $ComputerName -MACAddress $iDRACMacs[$i] -PortName "iDRAC" -InterfaceType "1000Base-T"
		Add-RTIPAddress -ObjectName $ComputerName -IPAddress $IPs[$i] -NICName iDRAC
		Update-RTSerialNumber -ComputerName $ComputerName -NewSerialNumber $SerialNumbers[$i]
		$i++
	}
.PARAMETER U
	This can either hold a single U or multiple Us
	You specify multiple Us like this: 1,2,3,4
	If it's a PDU, this parameter is not needed as it's assumed that the PDU spans the full length of the rack.
.PARAMETER FRONT
	These are reserved for PDUs. Don't specify for servers
.PARAMETER BACK
	These are reserved for PDUs. Don't specify for servers
.PARAMETER TOPHALF
	If there are 4 PDUs in the rack they need to be split in two or else they can't be entered in racktables
.PARAMETER BOTTOMHALF
	If there are 4 PDUs in the rack they need to be split in two or else they can't be entered in racktables
.NOTES
	
.LINK

#>
param(
[Parameter(Position=0, Mandatory=$True)]$ObjectName,
[Parameter(Position=1)]$RackName,
[Parameter(Position=2)]$ObjectType,
[array]$U,
[string]$ParentName,
[switch]$Front,
[switch]$Back,
[switch]$TopHalf,
[switch]$BottomHalf
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Add-RTObject.log"
	$ObjectTypeQuery = Connect-RTToMysql -Query "select * from Dictionary where chapter_id = 1"
	
	$Row = $ObjectTypeQuery | where-object {$_.dict_value -eq $ObjectType}
	if(!$Row)
	{
		$Message = "$(get-date) ObjectType $ObjectType not found, these were your options:"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		(Connect-RTToMysql -Query "select dict_value from Dictionary where chapter_id = 1").dict_value
		Return
	}
	$ObjectTypeID = $Row.dict_key
	if($RackName)
	{
		if($ParentName)
		{
			$Message = "$(get-date) You can't specify both a RackName and a ParentName, it's either/or."
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
		
		$RackIDQuery = Connect-RTToMysql -Query "select * from Object where name = '$RackName' and objtype_id = 1560"
		$RackID = $RackIDQuery.id
		if(!$RackID)
		{
			$Message = "$(get-date) RackName $RackName not found"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
		if($ObjectType -eq "PDU")
		{
			# all the Us from the Rack.
			$Height = $Null
			$Height = (Connect-RTToMysql -Query "select height from Rack where id = $RackID").height
			
			if($TopHalf)
			{
				$U = (($Height /2) + 1)..$Height
			}elseif($BottomHalf)
			{
				$U = 1..($Height /2)
			}else{
				$U = 1..$Height
			}
		}
		if(!$U)
		{
			$Message = "$(get-date) $(get-date) No U was specified"
			write-host $Message -fore red
			Return
		}
		
		$Atom = $Null
		if($Front){$Atom = "front"}
		if($Back){$Atom = "rear"}
		if(!$Atom){$Atom = "interior"}
		
		$ULow = $U[0]
		$UHigh = $U[$U.count -1]
		$UQuery = Connect-RTToMysql -Query "select * from RackSpace where rack_id = $RackID and unit_no BETWEEN $ULow and $UHigh and atom = '$Atom'"
		if($UQuery)
		{
			$Message = "$(get-date) Some of the Us specified in rack $RackName are already in use ($U)"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			Return
		}
	}else{
		$ParentObjectNameCheck = $Null
		$ParentObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ParentName'"
		if(!$ParentObjectNameCheck)
		{
			$Message = "$(get-date) ParentObjectNameCheck $ParentObjectNameCheck does not exist in Racktables"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
	}
	$ObjectNameCheck = $Null
	$ObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ObjectName'"
	if($ObjectNameCheck)
	{
		$Message = "$(get-date) ObjectName $ObjectName already exists in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$Message = "$(get-date) Creating Object $ObjectName with ObjectTypeID $ObjectTypeID"
	write-host $Message -fore cyan
	$Message | Out-MyLogFile $logfile -append
	
	Connect-RTToMysql -Query "INSERT INTO Object (name,objtype_id)
							VALUES ('$ObjectName','$ObjectTypeID');"
	start-sleep 1
	$ObjectQuery = $Null
	$ObjectQuery = Connect-RTToMysql -Query "select * from Object where name = '$ObjectName'"
	if(!$ObjectQuery)
	{
		$Message = "$(get-date) ObjectName $ObjectName Failed to add Object to Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Update-RTHistory -ObjectName $ObjectName -comment "Created Object $ObjectName with $ObjectType $$ObjectType"
	
	$ObjectID = $ObjectQuery.id
	if($RackName)
	{
		$Message = "$(get-date) Assigning Us to $ObjectName in Rack $RackName ($($U[0])..$($U[$U.count -1]))"
		write-host $Message -fore cyan
		$Message | Out-MyLogFile $logfile -append
		if($ObjectID)
		{
			foreach($SingleU in $U)
			{
				Connect-RTToMysql -Query "INSERT INTO RackSpace (rack_id,unit_no,atom,state,object_id)
											VALUES ('$RackID','$SingleU','$Atom','T','$ObjectID');"
			}
		}else{
			$Message = "$(get-date) Could not find newly created ObjectID for ObjectName $ObjectName"
			write-host $Message -fore red
			$Message | Out-MyLogFile $logfile -append
			return
		}
		Update-RTHistory -ObjectName $ObjectName -comment "Assigned Us to $ObjectName in Rack $RackName ($($U[0])..$($U[$U.count -1]))"
	}
	if($ParentName)
	{
		Connect-RTToMysql -Query "INSERT INTO EntityLink (parent_entity_type, parent_entity_id, child_entity_type, child_entity_id)
								VALUES ('object','$($ParentObjectNameCheck.id)','object','$ObjectID');"
											
		Update-RTHistory -ObjectName $ObjectName -comment "Assigned to Parent $ParentName"
		Update-RTHistory -ObjectName $ParentName -comment "Assigned child $ObjectName"
	}
}

#=====================================================================
#Add-RTIPAddress
#=====================================================================
Function Add-RTIPAddress
{
<#
.SYNOPSIS
	Adds an IP address to an object in racktables.
.DESCRIPTION

.EXAMPLE
	Add-RTIPAddress -ObjectName RACK01TEST -IPAddress 10.10.10.10 -NICName Nic1

.NOTES
	will verify that the IP falls in an existing ranges.
	Checks that the object name exists and that the IP isnt currently in use already.
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$True)]$IPAddress,
  [Parameter(Position=2, Mandatory=$True)][string]$NICName
)
	# map subnet masks to number of IPs in that range (yeah, it's lame... but whatever, it works - mostly. I think you can add broadcast addresses as valid IPs. Ah well, close enough).
	$HashTable = @{
	30=4
	29=8
	28=16
	27=32
	26=64
	25=128
	24=256
	23=512
	22=1024
	21=2048
	20=4096
	19=8192
	18=16384
	17=32768
	16=65536}

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Add-RTIPAddress.log" 
	$ObjectNameCheck = $Null
	$ObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ObjectName'"
	if(!$ObjectNameCheck)
	{
		$Message = "$(get-date) ObjectName $ObjectName not found in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$NICNameCheck = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where object_id = $($ObjectNameCheck.id) and name = '$NICName'"
	if($NICNameCheck)
	{
		$Message = "$(get-date) ObjectName $ObjectName already has a NIC named $NICName, can't have one NICName with two IPs (in Racktables..)"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$IPinINT64 = Convert-IPtoINT64 $IPAddress
	
	$IPInUseQuery = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where ip = $IPinINT64"
	
	if($IPInUseQuery)
	{
		$InUseObjectNameQuery = Connect-RTToMysql -Query "Select name FROM Object where id = '$($IPInUseQuery.object_id)'"
		$Message = "$(get-date) $IPAddress is already in use on $($InUseObjectNameQuery.name)"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$IPRangesQuery = Connect-RTToMysql -Query "select * from IPv4Network"
	if(!($IPRangesQuery | where-object {($_.ip -lt $IPinINT64) -and (($_.ip + $HashTable[[int]$_.mask]) -gt $IPinINT64)}))
	{
		$Message = "$(get-date) $IPAddress does not fall within any of the ranges in IPv4Network, so I can't assign"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Connect-RTToMysql -Query "INSERT INTO IPv4Allocation (object_id, ip, name, type) VALUES ($($ObjectNameCheck.id), $IPinINT64, '$NICName', 'regular');"
	
	$Message = "$(get-date) Added $IPAddress with $NICName to $ObjectName"
	$Message | Out-MyLogFile $logfile -append
	
	start-sleep 1
	$IPAddedQuery = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where ip = $IPinINT64"
	
	if(!($IPAddedQuery))
	{
		$Message = "$(get-date) Failed to add $IPAddress to $($ObjectNameCheck.name)"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Update-RTHistory -ObjectName $ObjectName -comment "Added $IPAddress with $NICName to $ObjectName"
}

#=====================================================================
# Remove-RTLink
#=====================================================================
Function Remove-RTLink
{
<#
.SYNOPSIS
	Removes a network connection between a switch and a device (server or other device with IP).
.DESCRIPTION
	Needs the object name (switch, server or other device type with a port) and the port name.
	Only removes the link, does not delete the port.
.EXAMPLE
	Remove-RTLink -ObjectName SERVER01 -PortName "Local Area Connection"
.EXAMPLE
	Remove-RTLink -ObjectName SWITCH01 -PortName "Gi 1/3"
.NOTES
	
.LINK

#>
Param(
[Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
[Parameter(Position=1, Mandatory=$True)][string]$PortName
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Remove-RTLink.log"
	
	$ObjectNameCheck = $Null
	$ObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ObjectName'"
	if(!$ObjectNameCheck)
	{
		$Message = "$(get-date) ObjectName $ObjectName not found in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$PortNameCheck = Connect-RTToMysql -Query "select * from Port where name = '$PortName' and object_id = $($ObjectNameCheck.id)"
	if(!$PortNameCheck)
	{
		$Message = "$(get-date) PortName $PortName not found in Racktables for ObjectName $ObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$LinkCheckQuery = Connect-RTToMysql -Query "select * from Link where portb = $($PortNameCheck.id) OR porta = $($PortNameCheck.id)"
	if(!$LinkCheckQuery)
	{
		$Message = "$(get-date) PortName $PortName on ObjectName $ObjectName is not linked to anything"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Connect-RTToMysql -Query "DELETE from Link where portb = $($PortNameCheck.id) OR porta = $($PortNameCheck.id)"
	
	$Message = "$(get-date) removed link from $PortName on ObjectName $ObjectName - porta: $($LinkCheckQuery.porta) portb: $($LinkCheckQuery.portb)"
	$Message | Out-MyLogFile $logfile -append

	start-sleep 1
	$SecondLinkCheckQuery = Connect-RTToMysql -Query "select * from Link where portb = $($PortNameCheck.id) OR porta = $($PortNameCheck.id)"
	if($SecondLinkCheckQuery)
	{
		$Message = "$(get-date) Failed to remove link from $PortName on ObjectName $ObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Update-RTHistory -ObjectName $ObjectName -comment "removed link from $PortName - porta: $($LinkCheckQuery.porta) portb: $($LinkCheckQuery.portb)"
}

#=====================================================================
# Convert-IPtoINT64
#=====================================================================
function Convert-IPtoINT64
{
param($ip) 

$octets = $ip.split(".")
return [int64]([int64]$octets[0]*16777216 +[int64]$octets[1]*65536 +[int64]$octets[2]*256 +[int64]$octets[3])
}

#=====================================================================
# Remove-RTIPAddress
#=====================================================================
Function Remove-RTIPAddress
{
<#
.SYNOPSIS
	Removes an IP Address that's assigned to an Object.
.DESCRIPTION

.EXAMPLE
	Remove-RTIPAddress -ObjectName SERVER01 -IPAddress 10.0.0.1
.NOTES

.LINK

#>
Param(
[Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
[Parameter(Position=1, Mandatory=$True)][string]$IPAddress
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Remove-RTIPAddress.log"
	
	$ObjectNameCheck = $Null
	$ObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ObjectName'"
	if(!$ObjectNameCheck)
	{
		$Message = "$(get-date) ObjectName $ObjectName not found in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	
	$IPinINT64 = Convert-IPtoINT64 $IPAddress
	
	$IPInUseQuery = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where ip = $IPinINT64 AND object_id = $($ObjectNameCheck.id)"
	if(!$IPInUseQuery)
	{
		$Message = "$(get-date) $IPAddress is not in use by $ObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Connect-RTToMysql -Query "DELETE from IPv4Allocation where ip = $IPinINT64 AND object_id = $($ObjectNameCheck.id)"
	start-sleep 1
	$IPInUseSecondCheckQuery = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where ip = $IPinINT64 AND object_id = $($ObjectNameCheck.id)"
	if($IPInUseSecondCheckQuery)
	{
		$Message = "$(get-date) Failed to remove IPAddress $IPAddress with NicName $($IPInUseQuery.name) from ObjectName $ObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Update-RTHistory -ObjectName $ObjectName -comment "Removed IPAddress $IPAddress with NicName $($IPInUseQuery.name)"
}

#=====================================================================
# Add-RTLink
#=====================================================================
Function Add-RTLink
{
<#
.SYNOPSIS
	
.DESCRIPTION
	
.EXAMPLE
	Add-RTLink -FromObjectName RACK01TESTx -FromPortName "test" -ToObjectName RACK01S2 -ToPortName "Gi 1/10"
.EXAMPLE
	Add-RTLink -FromObjectName RACK01TESTx -FromPortName "test" -ToObjectName RACK01S2 -ToPortName "Gi 1/10" -JustFixIt
.PARAMETER JustFixIt
	When this switch is specified, and the TO and FROM ports are mismatching in type, the FROM port will be changed to match the TO type.
.NOTES
	
.LINK

#>
Param(
[Parameter(Position=0, Mandatory=$True)][string]$FromObjectName,
[Parameter(Position=1, Mandatory=$True)][string]$FromPortName,
[Parameter(Position=2, Mandatory=$True)][string]$ToObjectName,
[Parameter(Position=3, Mandatory=$True)][string]$ToPortName,
[switch]$JustFixIt
)

	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Remove-RTLink.log"
	# FROM Checks-------------------------------------------------------
	#$ObjectNameCheck = $Null
	$FromObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$FromObjectName'"
	if(!$FromObjectNameCheck)
	{
		$Message = "$(get-date) FromObjectName $FromObjectName not found in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$FromPortNameCheck = Connect-RTToMysql -Query "select * from Port where name = '$FromPortName' and object_id = $($FromObjectNameCheck.id)"
	if(!$FromPortNameCheck)
	{
		$Message = "$(get-date) FromPortName $FromPortName not found in Racktables for FromObjectName $FromObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$FromLinkCheckQuery = Connect-RTToMysql -Query "select * from Link where portb = $($FromPortNameCheck.id) OR porta = $($FromPortNameCheck.id)"
	if($FromLinkCheckQuery)
	{
		$Message = "$(get-date) FromPortName $FromPortName on FromObjectName $FromObjectName is already connected to something"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	# TO Checks-------------------------------------------------------
	#$ObjectNameCheck = $Null
	$ToObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ToObjectName'"
	if(!$ToObjectNameCheck)
	{
		$Message = "$(get-date) ToObjectName $ToObjectName not found in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$ToPortNameCheck = Connect-RTToMysql -Query "select * from Port where name = '$ToPortName' and object_id = $($ToObjectNameCheck.id)"
	if(!$ToPortNameCheck)
	{
		$Message = "$(get-date) ToPortName $ToPortName not found in Racktables for ToObjectName $ToObjectName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$ToLinkCheckQuery = Connect-RTToMysql -Query "select * from Link where portb = $($ToPortNameCheck.id) OR porta = $($ToPortNameCheck.id)"
	if($ToLinkCheckQuery)
	{
		$Message = "$(get-date) ToPortName $ToPortName on ToObjectName $ToObjectName is already connected to something"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	# -------------------------------------------------------
	
	if($FromPortNameCheck.type -ne $ToPortNameCheck.type)
	{
		$Message = "$(get-date) Type Mismatch! FROM:$($FromPortNameCheck.type) TO:$($ToPortNameCheck.type)"
		write-host $Message -fore yellow
		$Message | Out-MyLogFile $logfile -append
		
		if($JustFixIt)
		{
			$Message = "$(get-date) Going to just fix it by setting the FROM port type identical to the TO port type"
			write-host $Message -fore green
			$Message | Out-MyLogFile $logfile -append
			
			Connect-RTToMysql -Query "UPDATE Port
										SET type = $($ToPortNameCheck.type)
										where id = $($FromPortNameCheck.id)"
		}else{
			Return
		}
	}
	
	Connect-RTToMysql -Query "INSERT INTO Link (porta, portb) VALUES ($($FromPortNameCheck.id),$($ToPortNameCheck.id));"
	
	$Message = "$(get-date) Link added between $FromObjectName - $FromPortName - $($FromPortNameCheck.id) : $ToObjectName - $ToPortName - $($ToPortNameCheck.id)"
	$Message | Out-MyLogFile $logfile -append

	start-sleep 1
	$SecondLinkCheckQuery = Connect-RTToMysql -Query "select * from Link where portb = $($FromPortNameCheck.id) OR porta = $($FromPortNameCheck.id)"
	if(!$SecondLinkCheckQuery)
	{
		$Message = "$(get-date) Failed to add link between $FromObjectName - $FromPortName - $($FromPortNameCheck.id) : $ToObjectName - $ToPortName - $($ToPortNameCheck.id)"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Update-RTHistory -ObjectName $FromObjectName -comment "Link added between $FromObjectName - $FromPortName and $ToObjectName - $ToPortName"
}

#=====================================================================
#Set-RTTag
#=====================================================================
Function Set-RTTag
{
<#
.SYNOPSIS
	Assigns a tag to an object
.DESCRIPTION

.EXAMPLE
	Set-RTTag -ObjectName SERVER01 -Tag SomethingThatExists
.NOTES
	Checks before assigning to make sure it's not already assigned.
	Can only assign existing tags. New tags still need to be manually created first.
.LINK

#>
Param(
	[Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
	[Parameter(Position=1, Mandatory=$True)][string]$Tag
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Set-RTTag.log"
	
	
	$TagTreeQuery = Connect-RTToMysql -Query "select * FROM TagTree"
	if(!$TagTreeQuery)
	{
		$Message = "$(get-date) $ObjectName - Unable to connect to Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$MatchingTagRow = $TagTreeQuery | where-object {$_.tag -eq $Tag}
	if(!$MatchingTagRow)
	{
		$Message = "$(get-date) $ObjectName - Unable find a tag named `"$Tag`" in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$ObjectIDQuery = Connect-RTToMysql -Query "select * FROM Object WHERE name = '$ObjectName'"
	if(!$ObjectIDQuery)
	{
		$Message = "$(get-date) $ObjectName - Unable find a ObjectName named `"$ObjectName`" in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$TagStorageCheckQuery = Connect-RTToMysql -Query "select * FROM TagStorage WHERE tag_id = $($MatchingTagRow.id) AND entity_id = $($ObjectIDQuery.id)"
		if($TagStorageCheckQuery)
	{
		$Message = "$(get-date) $ObjectName - $tag is already applied to $ObjectName"
		write-host $Message -fore yellow
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Connect-RTToMysql -Query "INSERT INTO TagStorage (entity_realm, entity_id, tag_id, tag_is_assignable, user, date)
								VALUES ('object',
								$($ObjectIDQuery.id),
								$($MatchingTagRow.id),
								'yes',
								'racktables-sync',
								NOW()
								)"
}

#=====================================================================
#Remove-RTIPAddress
#=====================================================================
Function Remove-RTIPAddress
{
<#
.SYNOPSIS
	Removes an IP address from an object in racktables.
.DESCRIPTION

.EXAMPLE
	Remove-RTIPAddress -ObjectName RACK01TEST -IPAddress 10.10.10.10
.NOTES
	Checks that the ObjectName exists and that the IP is in use by ObjectName.
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)][string]$ObjectName,
  [Parameter(Position=1, Mandatory=$True)]$IPAddress
)
	$LogFilePath = join-path $PsScriptRoot "logs"
	$logfile = join-path $LogFilePath "Remove-RTIPAddress.log" 
	$ObjectNameCheck = $Null
	$ObjectNameCheck = Connect-RTToMysql -Query "select * from Object where name = '$ObjectName'"
	if(!$ObjectNameCheck)
	{
		$Message = "$(get-date) ObjectName $ObjectName not found in Racktables"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$IPinINT64 = Convert-IPtoINT64 $IPAddress
	$IPInUseQuery = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where ip = $IPinINT64"
	if(!$IPInUseQuery)
	{
		$Message = "$(get-date) $IPAddress is not assigned to any computer in Racktables"
		write-host $Message -fore Yellow
		$Message | Out-MyLogFile $logfile -append
		return
	}
	if($IPInUseQuery -is [array])
	{
		$Message = "$(get-date) $IPAddress is assigned multiple times in Racktables, so I don't want to make changes to it. Sorry."
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	$NICName = $IPInUseQuery.name
	$InUseObjectNameQuery = Connect-RTToMysql -Query "Select name FROM Object where id = '$($IPInUseQuery.object_id)'"
	if(!($InUseObjectNameQuery.name -eq $ObjectName))
	{
		$Message = "$(get-date) $IPAddress is in use, but not on ObjectName $ObjectName, I found it instead on Objectname $($InUseObjectNameQuery.name) - NicName $NICName"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Connect-RTToMysql -Query "DELETE FROM IPv4Allocation WHERE ip = $IPinINT64"
	$Message = "$(get-date) Removed $IPAddress from $ObjectName"
	$Message | Out-MyLogFile $logfile -append
	start-sleep 1
	$IPAddedQuery = Connect-RTToMysql -Query "Select * FROM IPv4Allocation where ip = $IPinINT64"
	if($IPAddedQuery)
	{
		$Message = "$(get-date) Failed to remove $IPAddress From $($ObjectNameCheck.name)"
		write-host $Message -fore red
		$Message | Out-MyLogFile $logfile -append
		return
	}
	Update-RTHistory -ObjectName $ObjectName -comment "Removed $IPAddress with $NICName from $ObjectName"
}

#=====================================================================
#Compare-RTMACAddressesToWMI
#=====================================================================
Function Compare-RTMACAddressesToWMI
{
<#
.SYNOPSIS
	Compares Racktables Port names and MacAddresses against WMI.
.DESCRIPTION
	Will fix several instances:
	
	incorrectly named ports.
	
	in all other instances it will exit.
	Requires a Racktables object, see Get-RTServerDetails for details.
	
	Fixes:
	
	MAC matches, but nic name mismatches (will update racktables).
.EXAMPLE
	$i = 237
	foreach($RacktablesObject in $RacktablesComputers[$i..425])
	{
		write-host "$($RacktablesObject.ComputerName) - $i" -fore magenta
		Compare-RTMACAddressesToWMI -RacktablesObject $RacktablesObject
		$i++
	}
.NOTES
	filters out iDRACs
.LINK

#>
Param(
  [Parameter(Position=0, Mandatory=$True)]$RacktablesObject
)
	$FilteredServiceNames = "Rasl2tp","RasSstp", "RasAgileVpn","PptpMiniport","RasPppoe","NdisWan","NdisWan","kdnic","tunnel","RasGre","usbrndis6"
	$FilteredNics = $Null
	$FilteredNics = get-wmiobject win32_networkadapter -ComputerName $RacktablesObject.ComputerName | where-object {!($FilteredServiceNames -eq $_.ServiceName)} | where-object {$_.MacAddress}
	if(!$FilteredNics)
	{
		write-host "$($RacktablesObject.ComputerName) - Didn't respond to WMI."
		return
	}
	#$FilteredNicsUnmatched = $FilteredNics
	
	if(!(($RacktablesObject.Ports.MacAddress.count) -eq ($RacktablesObject.Ports.Macaddress | sort-object | get-unique).count))
	{
		write-host "$($RacktablesObject.ComputerName) - Duplicate mac in racktables? $($RacktablesObject.Ports.MacAddress.count) - $(($RacktablesObject.Ports.Macaddress | sort-object | get-unique).count)"
	}
	foreach($PortObject in ($RacktablesObject.Ports | where-object {!($_.Portname -eq "iDRAC") -and !($_.Portname -eq "iBMC")}))
	{
		if(!($PortObject.MacAddress))
		{
			write-host "$($RacktablesObject.ComputerName) $($PortObject.PortName) - missing MacAddress in Racktables" -fore Yellow
			
			$MatchingNic = $Null
			$MatchingNic = $FilteredNics | where-object {$_.NetConnectionID -eq $PortObject.PortName}
			if($MatchingNic)
			{
				$RAWMacAddress = $Null
				$RAWMacAddress = $MatchingNic.MACAddress
				if($RAWMacAddress)
				{
					$MacAddress = $Null
					$MacAddress = ($MatchingNic.MACAddress).Replace(":","")
					write-host "$($RacktablesObject.ComputerName) $($PortObject.PortName) - Found Nic in WMI that matches Portname in Racktables, Updating Racktables" -fore Green
					Connect-RTToMysql -Query "Update Port SET l2address = '$MacAddress' WHERE object_id = $($RacktablesObject.ObjectId) and name = '$($PortObject.PortName)'"
					Update-RTHistory -ObjectName $RacktablesObject.ComputerName -comment "Added missing MAC:$MacAddress to port: $($PortObject.PortName)"
					
					$RacktablesObject = Get-RTServerDetails -ComputerName $RacktablesObject.ComputerName
					
				}else{
					write-host "$($RacktablesObject.ComputerName) $($PortObject.PortName) - Can't find macaddress for matching object in WMI" -fore Red
				}
			}else{
				write-host "$($RacktablesObject.ComputerName) $($PortObject.PortName) - missing MacAddress in Racktables - couldn't find matching Nic by name in WMI" -fore Red
				foreach($FilteredNic in $FilteredNics)
				{
					write-host "$($FilteredNic.MacAddress) - $($FilteredNic.NetConnectionID)"
				}
				return
			}
		}
		$MatchingNic = $Null
		$MatchingNic = $FilteredNics | where-object {$_.MACAddress -eq $PortObject.MacAddress}
		if($MatchingNic)
		{
			if($MatchingNic.NetConnectionID -eq $PortObject.PortName)
			{
				write-host "$($RacktablesObject.ComputerName) $($PortObject.PortName) - all good." -fore cyan
			}else{
				write-host "$($RacktablesObject.ComputerName) - $($PortObject.PortName) does not match on name (but does match on MAC)"
				write-host "$($RacktablesObject.ComputerName) - $($PortObject.MacAddress) - $($MatchingNic.NetConnectionID)"
				$MacAddressCleaned = $Null
				$MacAddressCleaned = $PortObject.MacAddress.replace(":","")
				write-host "Update Port SET name = '$($MatchingNic.NetConnectionID)' WHERE Port.l2address = '$MacAddressCleaned'"
				Connect-RTToMysql -Query "Update Port SET name = '$($MatchingNic.NetConnectionID)' WHERE Port.l2address = '$MacAddressCleaned'"
				Update-RTHistory -ObjectName $RacktablesObject.ComputerName -comment "changed port name to $($MatchingNic.NetConnectionID) for port with MAC $MacAddressCleaned"
			}
			$FilteredNics = $FilteredNics | where-object {!($_.MACAddress -eq $PortObject.MacAddress)}
		}else{
			$MatchingNic = $Null
			$MatchingNic = $FilteredNics | where-object {$_.NetConnectionID -eq $PortObject.PortName}
			if($MatchingNic)
			{
				if(($MatchingNic.NetConnectionStatus -eq 0) -or ($MatchingNic.NetConnectionStatus -eq 4))
				{
					write-host "$($RacktablesObject.ComputerName) - $($PortObject.PortName) is disabled"
				}
			}else{
				write-host "$($RacktablesObject.ComputerName) - $($PortObject.PortName) does not match on MAC or name"
				write-host "$($RacktablesObject.ComputerName) - Racktables: $($PortObject.MacAddress)"
				return
			}
		}
	}
}


#=====================================================================
# Get-MyCredential
#=====================================================================
function Get-MyCredential
{
<#
.SYNOPSIS
	Get-MyCredential
.DESCRIPTION
	If a credential is stored in $CredPath, it will be used.
	If no credential is found, Export-Credential will start and offer to
	Store a credential at the location specified.
.EXAMPLE
	Get-MyCredential -CredPath `$CredPath
.NOTES

.LINK


#>
param(
[Parameter(Position=0, Mandatory=$true)]$CredPath,
$UserName
)
	if (!(Test-Path -Path $CredPath -PathType Leaf)) {
		Export-Credential (Get-Credential -Credential $UserName) $CredPath
	}
	$cred = Import-Clixml $CredPath
	$cred.Password = $cred.Password | ConvertTo-SecureString
	$Credential = New-Object System.Management.Automation.PsCredential($cred.UserName, $cred.Password)
	Return $Credential
}
#=====================================================================
# Export-Credential
#=====================================================================
function Export-Credential
<#
.SYNOPSIS
	
.DESCRIPTION
	This saves a credential to an XML File, for use with Get-MyCredential
.EXAMPLE
	Export-Credential $CredentialObject $FileToSaveTo
.NOTES

.LINK

#>
{
param(
$cred,
$path
)
      $cred = $cred | Select-Object *
      $cred.password = $cred.Password | ConvertFrom-SecureString
      $cred | Export-Clixml $path
}