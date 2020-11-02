param (
    [string]$target = "Berri"
	)

Function Get-FSMO {

	$forest = Get-ADForest (Get-ADForest).Name
	$domaines = (Get-ADForest).Domains

	"DomainNamingMaster", "SchemaMaster" | %{
		Write-Progress -Activity "Querying $PSItem for $forest"
		$DomainController = (Get-ADForest $forest).$PSItem
		#$VM = Get-VM | ?{$_.Guest.Hostname -eq $DomainController} |select @{Name="Datacenter";Expression={Get-Datacenter -VMHost $_.vmhost}}
		$VM = Get-VM $DomainController.Split(".")[0] | Select @{Name="Datacenter";Expression={Get-Datacenter -VMHost $_.vmhost}}
		[pscustomobject]@{FSMO=$PSItem; Forest=$forest; Domain=$null; DC=$DomainController; Datacenter=$VM.Datacenter.Name}
	}

	ForEach ($domaine In $domaines) {
		"InfrastructureMaster", "PDCEmulator", "RIDMaster" | %{
			Write-Progress -Activity "Querying $PSItem for $domaine"
			$DomainController = (Get-ADDomain $domaine).$PSItem
			#$VM = Get-VM | ?{$_.Guest.Hostname -eq $DomainController} |select @{Name="Datacenter";Expression={Get-Datacenter -VMHost $_.vmhost}}
			$VM = Get-VM $DomainController.Split(".")[0] | Select @{Name="Datacenter";Expression={Get-Datacenter -VMHost $_.vmhost}}
			[pscustomobject]@{FSMO=$PSItem; Forest=$null; Domain=$domaine; DC=$DomainController; Datacenter=$VM.Datacenter.Name}
		}
	}

}

if ( $global:DefaultVIServers.Count -ne 2 ) {
	Write-Host "Connecting to vCenters..."
	try { Connect-VIServer -Server pvavcs0a.bnquebec.ca,pvavcs0b.bnquebec.ca -SaveCredentials }
	catch { Write-Host -Foreground Red "Need connection to both vCenters."; exit }
}

if ( $global:DefaultVIServers.Count -ne 2 ) { Write-Host -Foreground Red "Need connection to both vCenters."; exit }

Import-Module ActiveDirectory

Write-Host "Checking FSMOs for currents forest/domains. Target:" $target

$fsmo = Get-FSMO
if ( $fsmo.Datacenter -NotMatch $target ) {
	Write-Host -ForegroundColor Black -BackgroundColor Green "One or more FSMO is not enabled on" $target
	$fsmo | Format-Table -AutoSize
	$fsmo | %{
		If ( $PSItem.Datacenter -ne $target ) {
			Write-Host -ForegroundColor Black -BackgroundColor DarkRed "$($PSItem.FSMO) on $($PSItem.DC) is not on target DC."
			if ( $PSItem.Forest ) { $AllDCs = ( Get-ADDomainController -Filter * -Server $PSItem.Forest).Name }
			if ( $PSItem.Domain ) { $AllDCs = ( Get-ADDomainController -Filter * -Server $PSItem.Domain).Name }
			[string]$OtherDC = $AllDCs -NotMatch $PSItem.DC.Split(".")[0]
			$VM = Get-VM $OtherDC | Select @{Name="Datacenter";Expression={Get-Datacenter -VMHost $_.vmhost}}
			if ( $VM.Datacenter.Name -eq $target ) {
				Write-Host -ForegroundColor Black -BackgroundColor Green "Moving $($PSItem.FSMO) from $($PSItem.DC) to $OtherDC ($target)"
				Move-ADDirectoryServerOperationMasterRole -Identity $OtherDC -OperationMasterRole $PSItem.FSMO -Verbose
			}
		}
	}
	$fsmo = Get-FSMO
	$fsmo | Format-Table -AutoSize
	if ( $fsmo.Datacenter -NotMatch $target ) { Write-Host -ForegroundColor Black -BackgroundColor DarkRed "One or more FSMO is not enabled on" $target }
} else {
	Write-Host -ForegroundColor Black -BackgroundColor Green "All FSMOs are enabled on" $target
}