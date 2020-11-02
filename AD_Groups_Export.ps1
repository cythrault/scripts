$DateTime = Get-Date -f "yyyy-MM"
$outfile = "AD_Groups"+$DateTime+".xlsx"
$output = @()
$ADGroups = Get-ADGroup -Filter {name -like "LaserFiche*"} -Server bnquebec.ca
$ADGroups += Get-ADGroup -Filter {name -like "LaserFiche*"} -Server biblio.bnquebec.ca

$i=0
$tot = $ADGroups.count

foreach ($ADGroup in $ADGroups) {
	$i++
	$status = "{0:N0}" -f ($i / $tot * 100)
	Write-Progress -Activity "Exporting AD Groups" -status "Processing Group $i of $tot : $status% Completed" -PercentComplete ($i / $tot * 100)

	$members = ""
	$membersarr = Get-ADGroup -filter {Name -eq $ADGroup.Name} -Server bnquebec.ca | Get-ADGroupMember | select Name
	$membersarr += Get-ADGroup -filter {Name -eq $ADGroup.Name} -Server biblio.bnquebec.ca | Get-ADGroupMember | select Name
	if ($membersarr) {
		foreach ($member in $membersarr) {
			$members = $members + ";" + $member.Name
		}
		$members = $members.Substring(1,($members.Length) -1)
	}

	$hashtabl = $NULL
	$hashtabl = [ordered]@{
		"Name" = $ADGroup.Name
		"Category" = $ADGroup.GroupCategory
		"Scope" = $ADGroup.GroupScope
		"Members" = $Members
	}

	$output += New-Object PSObject -Property $hashtabl
}

$output | Sort-Object Name | Export-Excel -AutoSize -FreezeTopRow $outfile