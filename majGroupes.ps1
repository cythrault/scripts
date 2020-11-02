#requires -version 3
$ErrorActionPreference = "SilentlyContinue"

$workbook = 'P:\projets\Modifications aux groupes Outlook.xlsx'
$groupes = Import-Excel -Path $workbook -WorksheetName Modifications

$groupes | ForEach-Object {
	[string]$displayname = $PSItem."Nouveaux groupes"
	[string]$shortname = $PSItem."Noms Courts"
	#Write-Host $displayname
	#Write-Host $shortname
	
	if (($PSItem."Nouveaux groupes" -ne "supprimer")) {
		[string]$proprietaire = $PSItem."Nouveaux proprietaires"
		if ( $proprietaire ) {
			[string]$managedby = ( Get-ADUser -Server biblio.bnquebec.ca -filter {displayName -eq $proprietaire} ).distinguishedName
		} else {
			Write-Host -BackgroundColor DarkRed "Nouveau propriétaire invalide ou non défini. $proprietaire"
			return
		}
		if (-Not $managedby ) { Write-Host -BackgroundColor DarkRed "Nouveau propriétaire invalide ou non défini. $proprietaire"; return }
	}
	
	[string]$server = switch -wildcard ( $PSItem.DN ) {
		"*DC=archives,DC=bnquebec,DC=ca" { "PVWPDC0C.archives.bnquebec.ca" }
		"*DC=biblio,DC=bnquebec,DC=ca" { "PVWPDC0G.biblio.bnquebec.ca" }
		default { "PVWPDC0D.bnquebec.ca" }
	}
	
	if ($PSItem."Nouveaux groupes" -eq "supprimer") { 
		Write-Host -BackgroundColor DarkGreen "Supression du groupe $displayName"
		Remove-DistributionGroup -Identity $PSItem.DN -DomainController $server -BypassSecurityGroupManagerCheck
		}
	
	if ($PSItem.DN -eq $null) {
		Write-Host -BackgroundColor DarkGreen "Ajout du groupe $displayName"
		New-DistributionGroup -Name $displayname -Alias $shortname -DisplayName $displayname -DomainController PVWPDC0D.bnquebec.ca `
			-ManagedBy $managedby -OrganizationalUnit "OU=Distributions,OU=Groupes,OU=1.Gestion,DC=bnquebec,DC=ca" `
			-PrimarySmtpAddress "$shortname@banq.qc.ca" -SamAccountName $shortname -Type Distribution		
	}
	
	if (($PSItem."Nouveaux groupes" -ne "supprimer") -and ($PSItem."Groupes Actuel" -ne $null)) {
		Write-Host -BackgroundColor DarkGreen "Modification du groupe $displayName"
		Set-DistributionGroup -Identity $PSItem.DN -Alias $shortname -BypassSecurityGroupManagerCheck -DisplayName $displayName `
			-DomainController $server -ManagedBy $managedby -Name $displayName -SamAccountName $shortname -SimpleDisplayName $shortname
	}
}
