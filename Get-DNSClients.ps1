$c = 0
$currentdate=$(get-date -uformat %Y%m%d-%H%M)
New-PSDrive -Name dest -PSProvider FileSystem -Root "\\bnquebec\tous\Informatique et télécommunications\DIT\EXP\Infrastructure\112.Services domaine AD.DHCP.DNS\1.Soutien"

#((Get-ADForest).Domains |%{ (Get-ADDomain -Server $_).replicadirectoryservers} |%{Get-ADDomainController -Identity $_ -Server $_}) | Select Name,Site,IsGlobalCatalog,IPv4Address,OperatingSystem | Export-Csv -notype dest:\all-dcs.csv
#$alldcs=$(((Get-ADForest).Domains |%{ (Get-ADDomain -Server $_).replicadirectoryservers} |%{Get-ADDomainController -Identity $_}) | Select Name, IPv4Address | Where-Object {$_.name -NotLike "CACOL2*" -and $_.name -NotLike "CACOL4*" -and $_.name -NotLike "USAAR*" -and $_.name -NotLike "UKCOL*" -and $_.name -NotLike "*COL1SE"})

#$alldcs|foreach {Set-DnsServerDiagnostics -ComputerName $_.Name -All $false}
#$alldcs|foreach {Set-DnsServerDiagnostics -ComputerName $_.Name -Queries $true -QuestionTransactions $true -ReceivePackets $true -UdpPackets $true -EnableLoggingToFile $true -MaxMBFileSize 10485760 -EnableLogFileRollover $false -LogFilePath DNSQueries.log}

$alldns = Import-Csv dest:\all-dcs.csv -Delimiter ','

$dcips = @()
(((Get-ADForest).Domains |%{ (Get-ADDomain -Server $_).replicadirectoryservers} |%{Get-ADDomainController -Identity $_ -Server $_}) | select IPv4Address) | foreach {$dcips += $($_.IPv4Address)}
$dcips += "::1"
$dcips += "127.0.0.1"

$alldns | ForEach {
	Write-Host -NoNewline "Crunching log from $($_.Name) ($([math]::Round($(Get-Childitem -file \\$($_.HostName)\c$\Windows\system32\dns\DNSQueries.log).Length/1024/1024,1))MB). "
	(Get-Content \\$($_.HostName)\c$\Windows\system32\dns\DNSQueries.log) | Select-String -Pattern $dcips -NotMatch -SimpleMatch | Select-String -Pattern "PACKET" -CaseSensitive > dest:\DNSQueries-$($_.Name).log
	(Get-Content dest:\DNSQueries-$($_.Name).log) | ? {$_.trim() -ne "" } | Set-Content dest:\DNSQueries-$($_.Name).log
	$requests = Import-Csv dest:\DNSQueries-$($_.Name).log -Delimiter ' ' -Header 'date','time','period','wt1','packet','s1','wt2','proto','dir','ClientIP'
	Write-Host -NoNewLine "Found $($requests.count) requests from "
	$requests | Add-Member -MemberType NoteProperty -Name Hostname -Value $null
	$requests | Add-Member -MemberType NoteProperty -Name DC -Value $($_.Name)
	$clients = $($requests | Sort -Property ClientIP -Unique)
	$clients | ForEach {
		try {
			$clients[$c].Hostname = $([System.Net.Dns]::GetHostByAddress($clients[$c].ClientIP).hostname)
		}
		catch {
			$clients[$c].Hostname = "Unresolvable"
		}
		$c++
	}
	$clients | Sort -Property ClientIP -Unique | Select ClientIP, Hostname, DC | Export-Excel -Path dest:\DNSClients-$($_.Name).xlsx
	$clients | Sort -Property ClientIP -Unique | Select ClientIP, Hostname, DC | Export-Excel -Append -Path dest:\DNSClients-AllDCs-$currentdate.xlsx
	Write-Host "$c unique clients."
	$c = 0
}