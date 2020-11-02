Function ParseDescription {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$description
	)	
	Process {
		if ( $description -eq "Project Team" ) { $roles = "Project Team" }
		else {
			$tower = $description.Split("-")[0].Trim()
			$roledesc = $description.Split("-")[1].Trim()
			if ( $roledesc -notlike "*Admins*" ) { $roles = $roledesc.Replace(" Admin","").Split("/").Trim() }
			else { $roles = $roledesc.Split("/").Trim()	}
		}
		if ( $roles -eq "Client Side Support" ) { $region = $description.Split("-")[2].Trim() } else { $region = $null }
		if ( $description -like "* - IT*" ) { $level = $description.Split("-")[2].Trim() } else { $level = "STD" }
		if ( $description -like "*Client Side Support* - IT*" ) { $level = $description.Split("-")[3].Trim() }
		[pscustomobject]@{ "Tower"=$tower; "Roles"=$roles; "Level"=$level; "Region"=$region }
	}
}

Function Get-AtosGroups {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$level,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string[]]$roles,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$region
	)
	Begin { $groups = @() }
	Process {
		if ( $level -eq "STD" ) { 
			$groups = "GRP-SEC-GLB-IO-VPNAccess","GRP-SEC-GLB-SPIO-R"
			if ( $roles -eq "Client Side Support" ) { $groups += "GRP-FWL-GLB-CSC" }
		}
		if ( $level -eq "IT4" ) {
			switch ( $roles ) {
				'placeholder'		{ $groups += "" }
			}
		}
		if ( $level -eq "IT5" ) { 
			switch ( $roles ) { # Admins
				'AD'					{ $groups += "GRP-RBC-R-CORP-IO-ADDS-Access" }
				'Application Packaging'	{ $groups += "GRP-RBC-R-GLB-IO-AppMgmt-Packaging" }
				'Malware'				{ $groups += "GRP-RBC-R-GLB-IO-AVAM-Admins" } # AV / Malware Admin
				'Backup'				{ $groups += "GRP-RBC-R-GLB-IO-Backup-Admins" }
				'BeatBox'				{ $groups += "GRP-RBC-R-GLB-IO-Beatbox-Admins" }
				'CES'					{ $groups += "GRP-RBC-F-GLB-CiscoEmailSecurity-ESA-Administrators", "GRP-RBC-F-GLB-CiscoEmailSecurity-SMA-Administrators" }
				'Citrix'				{ $groups += "GRP-RBC-R-GLB-IO-Citrix-Admins" }
				'Client Side Support'	{ $groups += "GRP-RBC-R-$region-IO-ClientSideSupport" }
				'DSE'					{ $groups += "GRP-RBC-R-GLB-IO-DSE" }
				'DSE Supervisor'		{ $groups += "GRP-RBC-R-GLB-IO-DSE-Supervisors" }
				'Exchange'				{ $groups += "GRP-RBC-R-GLB-IO-Exchange-Admins","GRP-RBC-F-GLB-CiscoEmailSecurity-SMA-MessageTracking" }
				'Flexera Admins'		{ $groups += "GRP-RBC-R-GLB-IO-Flexera-Admins" }
				'Flexera Operators'		{ $groups += "GRP-RBC-R-GLB-IO-Flexera-Operators" }
				'IAM'					{ $groups += "GRP-RBC-R-GLB-IO-IAM-Admins" }
				'IPAM'					{ $groups += "GRP-RBC-R-GLB-IO-IPAM-Admins" }
				'Linux'					{ $groups += "GRP-RBC-R-GLB-IO-Linux-Admins" }
				'Network'				{ $groups += "GRP-RBC-R-GLB-IO-Network-Admins" }
				'Network Tools'			{ $groups += "GRP-RBC-R-GLB-IO-NetworkTools-Admins" }
				'Netwrix'				{ $groups += "GRP-RBC-R-GLB-IO-Netwrix-Admins" }
				'PKI'					{ $groups += "GRP-RBC-R-GLB-IO-PKI-Admins" }
				'SCCM'					{ $groups += "GRP-RBC-R-GLB-IO-SCCM-Admins" }
				'SCOM'					{ $groups += "GRP-RBC-R-GLB-IO-SCOM-Admins" }
				'Secret Server Admins'	{ $groups += "GRP-RBC-R-GLB-IO-SecretServer-Admins" }
				'Service Desk Level 1'	{ $groups += "GRP-RBC-R-GLB-IO-ServiceDesk-Lvl1" }
				'Service Desk Level 2'	{ $groups += "GRP-RBC-R-GLB-IO-ServiceDesk-Lvl2" }
				'Service Desk Supervisor'	{ $groups += "GRP-RBC-R-GLB-IO-ServiceDesk-Supervisors" }
				'Service Now'			{ $groups += "GRP-RBC-R-GLB-IO-ServiceNow-Admins" }
				'SharePoint'			{ $groups += "GRP-RBC-R-GLB-IO-SharePoint-Admins" }
				'Skype'					{ $groups += "GRP-RBC-R-GLB-IO-S4B-Admins" }
				'SN eDiscovery Admins'	{ $groups += "GRP-RBC-R-GLB-IO-SneDiscovery-Admins" }
				'SQL'					{ $groups += "GRP-RBC-R-GLB-IO-Database-Admins" }
				'Storage'				{ $groups += "GRP-RBC-R-GLB-IO-Storage-Admins" }
				'Virtualisation'		{ $groups += "GRP-RBC-R-GLB-IO-Virtualisation-Admins" }
				'Voice'					{ $groups += "GRP-RBC-R-GLB-IO-Voice-Admins" }
			}
		}
		if ( $level -eq "IT9" ) {
			switch ( $roles ) {
				'AD' { $groups += "GRP-RBC-R-CORP-IO-ADDS-Admins" }
			}
		}
		if ( $level -eq "ITC" ) { 
			switch ( $roles ) {
				'Azure Subscription Contributor'	{ $groups += "GRP-RBC-R-GLB-IO-AzureSubscription-Contributor" }
				'Azure Subsciption Owner'			{ $groups += "GRP-RBC-R-GLB-IO-AzureSubscription-Owner" }
			}
		}
		return $groups
	}
}