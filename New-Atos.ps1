function New-Atos {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[ValidateLength(7,7)][ValidatePattern("[A|Y]\d{6}|[NL]\d{5}")]
		[Alias("Atos ID")]
		[string]$AtosID,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$givenName,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("surname")]
		[string]$sn,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$description,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("Email Address")]
		[string]$mail,
		[string]$dc = "dccan100dom20.corp.pbwan.net"
	)

	Begin {
		# define base OU
		$baseOU="OU=WSPObjects,DC=corp,DC=pbwan,DC=net"

		$ITC = @{
			'Enterprise Compute - Exchange / Skype Admin' = $True
			'Enterprise Compute - SharePoint Admin' = $True
			'BDS (Security Mgmt) - IAM Admin' = $True
		}

		$IT4 = @{
			'Enterprise Compute - Exchange / Skype Admin' = $True
		}

		$IT5 = @{
			'Enterprise Compute - AD Admin' = $True
			'Enterprise Compute - Application Packaging' = $True
			'Enterprise Compute - Backup Admin' = $True
			'Enterprise Compute - Citrix Admin' = $True
			'Enterprise Compute - Exchange / Skype Admin' = $True
			'Enterprise Compute - SCOM Admin' = $True
			'Enterprise Compute - SCCM Admin' = $True
			'Enterprise Compute - SharePoint Admin' = $True
			'Enterprise Compute - Storage Admin' = $True
			'Enterprise Compute - Virtualisation Admin' = $True
			'Network - Network Admin' = $True
			'BDS (Security Mgmt) - IAM Admin' = $True
			'Service Desk - Service Desk Level 1' = $True
			'Service Desk - Service Desk Level 2' = $True
			'Service Desk - Service Desk Supervisor' = $True
		}
		
		# define ITC OU
		$ITCtargetOU = @{
			'Enterprise Compute - Exchange / Skype Admin - ITC' = "OU=CLD Mgmt Users,OU=_Cloud Mgmt,$baseOU"
			'BDS (Security Mgmt) - IAM Admin - ITC' = "OU=CLD Mgmt Users,OU=_Cloud Mgmt,$baseOU"
			'Enterprise Compute - SharePoint Admin - ITC' = "OU=CLD Mgmt Users,OU=_Cloud Mgmt,$baseOU"
		}
		
		# define IT4 OU
		$IT4targetOU = @{
			'Enterprise Compute - Exchange / Skype Admin - IT4' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
		}
		
		# define IT5 OU - except for CSS
		$IT5targetOU = @{
			'Enterprise Compute - AD Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - Application Packaging - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - Backup Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - Citrix Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - Exchange / Skype Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - SCOM Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - SCCM Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - Storage Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - SharePoint Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Enterprise Compute - Virtualisation Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Network - Network Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'BDS (Security Mgmt) - IAM Admin - IT5' = "OU=GA Users,OU=_Resource Mgmt,$baseOU"
			'Service Desk - Service Desk Level 1 - IT5' = "OU=DA Users Lvl1,OU=_Resource Mgmt,$baseOU"
			'Service Desk - Service Desk Level 2 - IT5' = "OU=DA Users Lvl2,OU=_Resource Mgmt,$baseOU"
			'Service Desk - Service Desk Supervisor - IT5' = "OU=DA Users Lvl2,OU=_Resource Mgmt,$baseOU"
		}

		# define ITC groups
		$ITCgroups = @{
			'Enterprise Compute - Exchange / Skype Admin - ITC' = "GRP-RBC-R-GLB-IO-Exchange-Admins", "GRP-RBC-F-GLB-CiscoEmailSecurity-SMA-MessageTracking"
		}

		# define IT4 groups
		$IT4groups = @{
			'Enterprise Compute - Exchange / Skype Admin - IT4' = "GRP-RBC-R-GLB-IO-Exchange-Compliance-Admins"
		}

		# define IT5 groups - except for CSS
		$IT5groups = @{
			'Enterprise Compute - AD Admin - IT5' = "GRP-RBC-R-CORP-IO-ADDS-Access"
			'Enterprise Compute - Application Packaging - IT5' = "GRP-RBC-R-GLB-IO-AppMgmt-Packaging"
			'Enterprise Compute - Backup Admin - IT5' = "GRP-RBC-R-GLB-IO-Backup-Admins"
			'Enterprise Compute - Citrix Admin - IT5' = "GRP-RBC-R-GLB-IO-Citrix-Admins"
			'Enterprise Compute - Exchange / Skype Admin - IT5' = "GRP-RBC-R-GLB-IO-Exchange-Admins", "GRP-RBC-F-GLB-CiscoEmailSecurity-SMA-MessageTracking", "GRP-RBC-R-GLB-IO-S4B-Admins"
			'Enterprise Compute - SCOM Admin - IT5' = "GRP-RBC-R-GLB-IO-SCOM-Admins"
			'Enterprise Compute - SCCM Admin - IT5' = "GRP-RBC-R-GLB-IO-SCCM-Admins"
			'Enterprise Compute - Storage Admin - IT5' = "GRP-RBC-R-GLB-IO-Storage-Admins"
			'Enterprise Compute - SharePoint Admin - IT5' = "GRP-RBC-R-GLB-IO-SharePoint-Admins"
			'Enterprise Compute - Virtualisation Admin - IT5' = "GRP-RBC-R-GLB-IO-Virtualisation-Admins"
			'Network - Network Admin - IT5' = "GRP-RBC-R-GLB-IO-Network-Admins"
			'BDS (Security Mgmt) - IAM Admin - IT5' = "GRP-RBC-R-GLB-IO-IAM-Admins"
			'Service Desk - Service Desk Level 1 - IT5' = "GRP-RBC-R-GLB-IO-ServiceDesk-Lvl1"
			'Service Desk - Service Desk Level 2 - IT5' = "GRP-RBC-R-GLB-IO-ServiceDesk-Lvl2"
			'Service Desk - Service Desk Supervisor - IT5' = "GRP-RBC-R-GLB-IO-ServiceDesk-Supervisors"
		}
		
		# Fqdn                         Site
		# ----                         ----
		# enaarlyncpool.corp.pbwan.net Site:Enterprise-Culpeper
		# endenlyncpool.corp.pbwan.net Site:Enterprise-Denver
		# apbnelyncpool.corp.pbwan.net Site:AP-Brisbane
		# ealonlyncpool.corp.pbwan.net Site:EA-London
		# eastolyncpool.corp.pbwan.net Site:EA-Stockholm
		
		# define region's lync pools.
		# $lyncpool = $lyncpools[$region]
		$lyncpools = @{
			'AU' = 'apbnelyncpool.corp.pbwan.net'
			'GB' = 'ealonlyncpool.corp.pbwan.net'
			'KR' = 'apbnelyncpool.corp.pbwan.net'
			'CN' = 'apbnelyncpool.corp.pbwan.net'
			'SE' = 'eastolyncpool.corp.pbwan.net'
			'SG' = 'apbnelyncpool.corp.pbwan.net'
			'TW' = 'apbnelyncpool.corp.pbwan.net'
		}
		
		# define database per region
		# $exchdb = $exchdbs[$region]
		$exchdbs = @{
			'AU' = 'APDAG02'
			'CA' = 'CADAG01'
			'CN' = 'ASDAG02'
			'GB' = 'EADAG01'
			'KR' = 'ASDAG02'
			'SE' = 'SEDAG01'
			'SG' = 'ASDAG02'
			'TW' = 'ASDAG02'
			'US' = 'ENDAG01'
			'ZA' = 'ZADAG01'
		}
		
		Write-Host "Starting"
		
		# Remove Stale PSS
		$stalePSS = Get-PSSession | ?{($_.Availability -eq "None") -or ($_.State -eq "Closed")}
		if ( $stalePSS ) { 
			Write-Host -ForegroundColor Black -BackgroundColor Yellow "Removing Stale PS Sessions"
			Remove-PSSession $stalePSS
		}
		
		try {
			if (-Not $(Get-PSSession|where ConfigurationName -eq "Microsoft.Exchange")) {
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Exchange Session"
				$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://dccan100exc01/PowerShell/ -Authentication Kerberos -Credential (Get-Credential -Message "Enter your Exchange credentials")
				Import-PSSession $ExchSession | Out-Null
			}

			# TODO add exchange online session

			if (-Not $(Get-PSSession|where {$_.ComputerName -like "*lyncpool*" -or $_.ComputerName -like "*fepool*"}) ) {
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Lync Session"
				#$LyncSession = New-PSSession -ConnectionUri https://enaarlyncpool.corp.pbwan.net/OcsPowerShell -Credential (Get-Credential -Message "Enter your Lync credentials") #-Authentication NegotiateWithImplicitCredential 
				$LyncSession = New-PSSession -ConnectionUri https://glb-s4b-fepool-dccan100.corp.pbwan.net/OcsPowerShell -Credential $it5creds
				Import-PSSession $LyncSession | Out-Null
			}
		}	
		catch {
			Write-Host -ForegroundColor Black -BackgroundColor Red "Can't proceed without valid Exchange and Lync/S4B sessions"
			break
		}
	}
	
	Process {
		# normal account processing
	
		# define common attributes, trimming possible spaces from Excel
		$givenName = $givenname.Trim()	
		$sn = $sn.Trim()
		$displayName = "$sn, $givenName"
		$description = $_.description.Trim()
		$AtosID = $AtosID.Trim()
		$stdSAM = "XX$AtosID"
		$cn = "$displayName ($stdSAM)"
		$stdgroups = "GRP-SEC-GLB-IO-VPNAccess", "GRP-SEC-GLB-SPIO-R"
		
		# CSS staff gets client visibility identity, a mailbox and lync account
		if ( $description -like "End User Device - Client Side Support*" ) {
			$css = $true
			#$userPrincipalName = "$givenName.$sn.atos@wsp.com"
			$userPrincipalName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes("$givenName.$sn.atos@wsp.com"))
			$userPrincipalName = $userPrincipalName.Replace(" ",".")
			$userPrincipalName = $userPrincipalName.Replace("'","")
			$sip = $userPrincipalName
			$mail = $userPrincipalName
			$stdgroups += "GRP-FWL-GLB-CSC"
			$region = $description.Split("-").Trim()[2]
			# get random available database name from the region's Exchange DAG
			if ($exchdbs[$region] ) {
				$exchdb = $(Get-MailboxDatabase | where {($_.MasterServerOrAvailabilityGroup -eq $exchdbs[$region]) -and ($_.IsExcludedFromProvisioning -eq $False) -and ($_.IsSuspendedFromProvisioning -eq $False) }|Get-Random).Name
			} else {
				$exchdb = "CADAG01-DB01"
			}
			if ($lyncpools[$region]) { $lyncpool = $lyncpools[$region] } else { $lyncpool = "enaarlyncpool.corp.pbwan.net" }
		}
		else {
			$css = $false
			#$userPrincipalName = "$givenName.$sn@wsp.com"
			$userPrincipalName = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes("$givenName.$sn@wsp.com"))
			$userPrincipalName = $userPrincipalName.Replace(" ",".")
			$userPrincipalName = $userPrincipalName.Replace("'","")
			$mail = $mail.ToLower().Trim()
			$sip = $null
		}
		
		try {
			# create account
			New-ADUser -Server $dc -Name $cn -AccountPassword (ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force) `
				-DisplayName $displayName -Enabled $True -GivenName $givenName -Surname $sn `
				-Path "OU=IO Users,OU=_Resource Mgmt,$baseOU" -Samaccountname $stdSAM -UserPrincipalName $UserPrincipalName `
				-Description $description -Company Atos -Department "IT Infrastructure and Operations"
			
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - Account created"
			[pscustomobject]@{"Atos ID"=$AtosID;"Domain"="CORP";"givenname"=$givenName;"surname"=$sn;"description"=$description;`

				"SAMAccountName"=$stdSAM;"UserPrincipalName"=$UserPrincipalName;"CN"=$cn;"Email Address"=$mail;"SIP Address"=$sip}
			# enable mailbox + lync or mailuser depending if CSS
			if ($css) { 
				Enable-Mailbox -DomainController $dc $stdSAM -displayName $displayname -alias $stdSAM -primarySMTPAddress $mail -Database $exchdb | Out-Null
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - CSS Staff - Mailbox added"
				Enable-CsUser -Identity $stdSAM -RegistrarPool $lyncpool -SipAddress sip:$sip -DomainController $dc | Out-Null
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - CSS Staff - Lync-enabled"
			}
			else {
				Enable-MailUser -DomainController $dc $stdSAM -displayName $displayname -alias $stdSAM -externalemailaddress $mail -primarySMTPAddress $mail | Out-Null
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - Mail-enabled"
			}
				
			# add membership
			if ( $stdgroups ) { $stdgroups | %{ Add-ADGroupMember -Server $dc -Identity $_ -Members $stdSAM } }
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - Membership added"

			Write-Host -ForegroundColor Black -BackgroundColor Green "$stdSAM - Done"
		}
		
		catch {
			Write-Host -ForegroundColor Black -BackgroundColor Red "$stdSAM - $_"
			continue
		}

		# normal account processing done

		# IT4/IT5/ITC elevated account processing
		
		if ( $description -eq "Project Team" ) {
			# skip no-role specific accounts
			Write-Host -ForegroundColor Black -BackgroundColor Green "$stdSAM - Project Team - Elevated account not required"
		}
	
		if ( $ITC[$description] ) {
			# create ITC accounts
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - Creating ITC Account"
	
			$ITCdescription = $description + " - ITC"
			$sam = $stdSAM + "ITC"
			$cn = "$displayName ($sam)"
			$UserPrincipalName = $sam + "@wsp.com"
			$sip = $null
			if ( $ITCtargetOU[$ITCdescription] ) {
				$OU = $ITCtargetOU[$ITCdescription]
			} else {
				Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - Target OU for $description not defined"
				break
			}
			if ( $ITCgroups[$ITCdescription] ) {
				$groups = $ITCgroups[$ITCdescription]
			} else {
				Write-Host -ForegroundColor Black -BackgroundColor Yellow "$sam - group membership for $description not defined"
			}
			
			try {
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Creating ITC Account for $description"
				New-ADUser -Server $dc -Name $cn -AccountPassword (ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force) `
					-DisplayName $displayName -Enabled $True -GivenName $givenName -Surname $sn `
					-Path $OU -Samaccountname $sam -UserPrincipalName $UserPrincipalName -Description $ITCdescription
				
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Account created"
				[pscustomobject]@{"Atos ID"=$AtosID;"Domain"="CORP";"givenname"=$givenName;"surname"=$sn;"description"=$ITCdescription;`
					"SAMAccountName"=$sam;"UserPrincipalName"=$UserPrincipalName;"CN"=$cn;"Email Address"=$mail;"SIP Address"=$sip}
				
				if ( $groups ) { $groups | %{ Add-ADGroupMember -Server $dc -Identity $_ -Members $sam } }
				
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Membership added"

				if ( $description -like "Enterprise Compute - *Exchange*" ) {
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need WSPGROUP & MGWSP IT5 Accounts"
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need Exchange Admin role in O365"
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need local app account in DMarcian"
				}

				if ( $description -like "Enterprise Compute - *Skype*" ) {
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need Skype Admin role in O365"
				}

				if ( $description -like "Enterprise Compute - *SharePoint*" ) {
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need SharePoint Admin role in O365"
				}
				
				Write-Host -ForegroundColor Black -BackgroundColor Green "$sam - Done"
			}

			catch {
				Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - $_"
				continue
			}
		
		}

		
		if ( $IT4[$description] ) {
			# create IT4 accounts
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - Creating IT4 Account"

			$IT4description = $description + " - IT4"
			$sam = $stdSAM + "IT4"
			$cn = "$displayName ($sam)"
			$UserPrincipalName = $sam + "@wsp.com"
			$sip = $null
			if ( $IT4targetOU[$IT4description] ) {
				$OU = $IT4targetOU[$IT4description]
			} else {
				Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - Target OU for $description not defined"
				break
			}
			if ( $IT4groups[$IT4description] ) {
				$groups = $IT4groups[$IT4description]
			} else {
				Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - group membership for $description not defined"
				break
			}
			
			try {
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Creating IT4 Account for $description"
				New-ADUser  -Server $dc -Name $cn -AccountPassword (ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force) `
					-DisplayName $displayName -Enabled $True -GivenName $givenName -Surname $sn `
					-Path $OU -Samaccountname $sam -UserPrincipalName $UserPrincipalName -Description $IT4description
				
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Account created"
				[pscustomobject]@{"Atos ID"=$AtosID;"Domain"="CORP";"givenname"=$givenName;"surname"=$sn;"description"=$IT4description;`
					"SAMAccountName"=$sam;"UserPrincipalName"=$UserPrincipalName;"CN"=$cn;"Email Address"=$mail;"SIP Address"=$sip}
				
				if ( $groups ) { $groups | %{ Add-ADGroupMember -Server $dc -Identity $_ -Members $sam } }
				
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Membership added"

				if ( $description -like "Enterprise Compute - *Exchange*" ) { 
					# Enable the IT4 account as a mailbox, disable POP, MAPI and IMAP
					Enable-Mailbox -DomainController $dc $sam -displayName $displayname -alias $sam -primarySMTPAddress $sam@wsp.com | Out-Null
					Set-Mailbox -DomainController $dc $sam -HiddenFromAddressListsEnabled $True | Out-Null
					Set-CASMailbox -DomainController $dc $sam -OWAEnabled $False -ActiveSyncEnabled $False -PopEnabled $False -IMAPEnabled $False
					Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - IT4 Exchange Admin - Mailbox added"
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need Exchange Discovery Admin role in Exchange Online"
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need Admin role in Symantec EV.Cloud"
				}
				
				Write-Host -ForegroundColor Black -BackgroundColor Green "$sam - Done"
			}

			catch {
				Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - $_"
				continue
			}
		
		}

		if ( $IT5[$description] -OR $css ) {
			# create IT5 accounts
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$stdSAM - Creating IT5 Account"
		
			$IT5description = $description + " - IT5"
			$sam = $stdSAM + "IT5"
			$cn = "$displayName ($sam)"
			$UserPrincipalName = $sam + "@wsp.com"
			$sip = $null
			if ( $css ) {
				$OU = "OU=DA Users CSS,OU=_Resource Mgmt,$baseOU"
				$groups = "GRP-RBC-R-$region-IO-ClientSideSupport"
			}
			else {
				if ( $IT5targetOU[$IT5description] ) {
					$OU = $IT5targetOU[$IT5description]
				} else {
					Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - Target OU for $description not defined"
					break
				}
				if ( $IT5groups[$IT5description] ) {
					$groups = $IT5groups[$IT5description]
				} else {
					Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - group membership for $description not defined"
					break
				}
			}
			
			try {
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Creating IT5 Account for $description"
				New-ADUser -Server $dc -Name $cn -AccountPassword (ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force) `
					-DisplayName $displayName -Enabled $True -GivenName $givenName -Surname $sn `
					-Path $OU -Samaccountname $sam -UserPrincipalName $UserPrincipalName -Description $IT5description
				
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Account created"
				[pscustomobject]@{"Atos ID"=$AtosID;"Domain"="CORP";"givenname"=$givenName;"surname"=$sn;"description"=$IT5description;`
					"SAMAccountName"=$sam;"UserPrincipalName"=$UserPrincipalName;"CN"=$cn;"Email Address"=$mail;"SIP Address"=$sip}
				
				if ( $groups ) { $groups | %{ Add-ADGroupMember -Server $dc -Identity $_ -Members $sam } }
				
				Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "$sam - Membership added"

				if ( $description -like "Service Desk - Service Desk Level *" ) {
					Write-Host -ForegroundColor Black -BackgroundColor Magenta "$sam - Need Exchange Recipient Management role in Exchange Online"
				}

				Write-Host -ForegroundColor Black -BackgroundColor Green "$sam - Done"
			}

			catch {
				Write-Host -ForegroundColor Black -BackgroundColor Red "$sam - $_"
				continue
			}
		}
	
		# it9 processing
		# "OU=SA Users,OU=Domain Mgmt,$domainDN" 
		# $description = "$description - IT9"
		# $sam = $stdSAM + "IT9"
		# $cn = "$displayName ($sam)"
		# $UserPrincipalName = $sam + "@wsp.com"
		# membership GRP-RBC-R-CORP-IO-ADDS-Admins

	}
	
	End {
		if ( $ExchSession ) { Remove-PSSession $ExchSession }
		if ( $LyncSession ) { Remove-PSSession $LyncSession }
		Write-Host "Done"
	}
}






