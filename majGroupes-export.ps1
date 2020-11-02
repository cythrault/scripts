$props=@(
	'Name', 'distinguishedName', 'sAMAccountName', 'displayName', 'description',
	'mail', 'mailNickname', 'legacyExchangeDN', 'msExchHideFromAddressLists', 'msExchRemoteRecipientType',
	@{Name='proxyAddresses';Expression={ $_.proxyAddresses -join ';' }}, 'whenCreated', 'whenChanged',
	'managedBy',@{Name='managed';Expression={(Get-ADUSer $_.managedBy -Properties *).displayName}}
)
$workbook = 'P:\projets\Modifications aux groupes Outlook.xlsx'
$groupes = Import-Excel -Path $workbook -WorksheetName Modifications
$groupes | %{

	$displayName=$_."Groupes Actuel"

	if ((-Not $displayname) -and ($_."Nouveaux groupes")) {
		Write-Host -ForegroundColor Green "DL: $($_."Nouveaux groupes") :: nouvelle DL"
	} else {
		$groupe = Get-ADGroup -Server biblio.bnquebec.ca -filter {displayName -eq $displayName} -properties * | select $props
		if (-Not $groupe) { $groupe = Get-ADGroup -Server bnquebec.ca -Filter {displayName -eq $displayName} -properties * | select $props }
		if (-Not $groupe) { $groupe = Get-ADGroup -Server archives.bnquebec.ca -Filter {displayName -eq $displayName} -properties * | select $props }
	}
	
	if ( ( $_."Nouveaux groupes" -eq "à supprimer" ) -and ( $groupe ) ) {
		Write-Host -ForegroundColor Red "DL: $displayName :: $($groupe.sAMAccountName) :: $($groupe.managed) :: à détruire"
	} elseif (( $_."Nouveaux groupes" -eq "à supprimer" ) -and (-Not $groupe )) {
		Write-Host -BackgroundColor DarkRed "DL: $displayName :: Non-trouvé! :: Non-trouvé! :: à détruire"
	} elseif (($displayName) -and ($groupe)) {
		Write-Host "DL: $displayName :: $($groupe.sAMAccountName) :: $($groupe.managed)"
		$groupe | Export-Excel -Path $workbook -WorksheetName Actuel -Append -FreezeTopRow
	} elseif (($displayName) -and (-Not $groupe)) {
		Write-Host -BackgroundColor DarkRed "DL: $displayName :: Non-trouvé!"
	}
}