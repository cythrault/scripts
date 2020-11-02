Function Atos-IT9 {
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("SAMAccountName")]
		[string]$sam,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$givenName,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("Surname")]
		[string]$sn
	)

	#$ErrorActionPreference = "SilentlyContinue"
	$AESKey = Get-Content \\corp\ca\data$\users\Quebec\martin.thomas\scripts\AESKey.key
	
	$forests = @{
		'CORP' 			= @("wsp.com",	 	"dc=corp,dc=pbwan,dc=net", 	"corp.pbwan.net",	"camt055545it9")
		'DMZ' 			= @("dmz.local", 	"dc=dmz,dc=local", 			"dmz.local",		"mt1")
		'FOCUSED' 		= @("focuscorp.ca", "dc=focuscorp,dc=ca", 		"focuscorp.ca",		"admin-mt")
		'GCG' 			= @("gcg.local", 	"dc=gcg,dc=local", 			"gcg.local",		"mt2")
		'MMM' 			= @("mmm.ca", 		"dc=mmm,dc=ca", 			"mmm.ca",			"mt1")
		'WSPNET' 		= @("wspgroup.net", "dc=wspgroup,dc=net", 		"wspgroup.net",		"mt1")
#		'WSPGEO' 		= @("wspgeo.local", "dc=wspgeo,dc=local", 		"10.16.212.4",		"mt1")
#		'1WSPTEST' 		= @("1wsptest.com", "dc=1wsptest,dc=com", 		"10.254.11.202",	"")
#		'EEWSPGROUP' 	= @("wspgroup.ee", 	"dc=wspgroup,dc=ee", 		"10.254.11.210",	"")
	}
	
	$creds=@{}
	$forests.keys | %{
		if ( $forests[$_][3] ) {
			$domain = $_
			$username = $forests[$_][3]
			$pwdTxt =  Get-Content \\corp\ca\data$\users\Quebec\martin.thomas\scripts\$username.aes
			$securePwd = $pwdTxt | ConvertTo-SecureString -Key $AESKey
			$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList "$domain\$username", $securePwd
			$creds.Add( $_, $credentials )
		} else { $creds.Add( $_, $(Get-Credential "$_\") ) }
	}
	
	$forests.keys | %{
		$netbios = $_
		$suffix = $forests[$netbios][0]
		$dname = $forests[$netbios][1]
		$dc = $forests[$netbios][2]
		
		Write-Host "Adding $sam ($sn, $givenName) in $netbios - $suffix - $dname"
		
		New-ADUser -Name "$sn, $givenName ($sam)" -GivenName $givenName -Surname $sn -DisplayName "$sn, $givenName" `
			-Samaccountname $sam -UserPrincipalName "$sam@$suffix" -Description "$netbios Active Directory Admin - IT9" `
			-AccountPassword (ConvertTo-SecureString "DeL0re@nDMZ12" -AsPlainText -Force) `
			-Path "OU=SA Users,OU=Domain Mgmt,$dname" -Server $dc -Credential $creds[$netbios] -Enabled $True -Verbose
	}
}