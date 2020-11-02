# C:\Scripts\Get-ExchangeDelegates.ps1|Export-Excel Exchange-Delegates-20201005.xlsx -BoldTopRow -AutoSize -FreezeTopRow -TableName "Delegates" -WorksheetName "Delegates"
$domains = ([System.Directoryservices.Activedirectory.Forest]::GetCurrentForest()).Domains
$pdc = ([System.Directoryservices.Activedirectory.Forest]::GetCurrentForest()).RootDomain.PDCRoleOwner.Name
ForEach ($domain in $domains) {
	$dc = $domain.PdcRoleOwner.Name
	Write-Host "Extracting users from $($domain.Name) / $dc"
	$mbxs += Get-ADUser -Filter {(homeMDB -like "*") -and (msExchDelegateListLink -like "*")} -Properties msExchDelegateListLink, displayName -Server $dc
}

$c=0
ForEach ($mbx in $mbxs) {
	Write-Progress -activity "Extracting msExchDelegateListLink for $mbx" -percentComplete ($c++ / $mbxs.count*100)
	$mbx.msExchDelegateListLink | ForEach-Object{
		$user = Get-ADUser -Server $pdc`:3268 -Identity $PSItem -Properties displayName
		[pscustomobject]@{
			'Mailbox CN' 			= $mbx.Name
			'Mailbox Name' 			= $mbx.displayName
			'Mailbox DN' 			= $mbx.distinguishedName
			'User CN'				= $user.Name
			'User Name'				= $user.displayName
			'User DN'				= $user.distinguishedName
		}
	}
}