param(
	[string] $searchbase = "OU=Canada,DC=corp,DC=pbwan,DC=net",
	[string] $domain = "corp.pbwan.net",
	[Object[]] $servers = @(Get-ADComputer -SearchBase $searchbase -Server $domain -Filter {name -notlike "*nas0*" -and name -notlike "ca*-*" -and (name -like "ca*" -or name -like "sv*" -or name -like "serv*" -or name -like "srv*" -or name -like "vm*" -or name -like "dccan*")} | select name)
	)

$finallist = @()
Write-Host "Found $($servers.count) servers."

foreach($server in $servers){
	Write-Progress -Activity "scanning $($servers.count) servers" -Status "in progress" -PercentComplete ($cpt++ / $servers.count * 100) -CurrentOperation ("scanning $($server.name)")
	try {
		$networks = gwmi win32_networkadapterconfiguration -ComputerName $server.name -ErrorAction stop| ?{$_.ipenabled}
		
		foreach($network in $networks){
			$servobj = new-object psobject | select name, ipaddress, dnsSearchList, version
			$servobj.name = $server.name
			$servobj.ipaddress = $network.ipaddress[0]
			$dnsSearchlist = ""
			foreach($dnsserv in $network.DNSServerSearchOrder){
				if($dnsSearchlist -ne ""){
					$dnsSearchlist += ","
					}
				$dnsSearchlist += $dnsserv
				}
			$servobj.dnssearchlist = $dnsSearchlist
			$servobj.version = (gwmi win32_operatingsystem -ComputerName $server.name).caption
			Write-Output $servobj
		}

	}catch [Exception]{
		$servobj = new-object psobject | select name, ipaddress, dnsSearchList, version
		$servobj.name = $server.name
		$servobj.ipaddress = "couldn't resolve WMI"
		$servobj.dnssearchlist = "couldn't resolve WMI"
		$servobj.version = "couldn't resolve WMI"
		Write-Output $servobj
		}
	
	Write-Progress -Activity "scanning $($servers.count) servers" -Status "in progress" -PercentComplete ($cpt / $servers.count * 100) -CurrentOperation ("scanning $($server.name)")
}

Write-Host "Done."