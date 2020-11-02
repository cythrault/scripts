$source = Get-ADUser dchagnon
$dest = Get-ADUser mtrudeau
Get-ADUser -Identity $source -Properties MemberOf | Select-Object MemberOf -ExpandProperty MemberOf | %{
	[string]$server = switch -wildcard ( $PSItem ) {
		"*DC=archives,DC=bnquebec,DC=ca" { "archives.bnquebec.ca" }
		"*DC=biblio,DC=bnquebec,DC=ca" { "biblio.bnquebec.ca" }
		default { "bnquebec.ca" }
	}
	$user = Get-ADUser $dest | Select-Object -ExpandProperty distinguishedName
	$members = Get-ADGroupMember -Identity $PSItem -Server $server -Recursive | Select -ExpandProperty distinguishedName
	If ($members -contains $user) {
		Write-Host -ForegroundColor Black -BackgroundColor DarkYellow "$dest already exist in $PSItem"
	} Else {
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Adding $PSItem from $server"
		Add-AdGroupMember -Identity $PSItem -Server $server -Members $dest
	}
}