param (
[Parameter(
         Mandatory=$true,
         Position=1,
         HelpMessage="Machine name/IP?"
      )]
      [string]$machine)

#$server = Get-ADComputer -SearchBase $searchbase -Filter {name -eq $machine} | select name
$networks = gwmi win32_networkadapterconfiguration -ComputerName $machine -ErrorAction stop| ?{$_.ipenabled}
    foreach($network in $networks){
    $servobj = new-object psobject | select name, ipaddress, dnsSearchList
	echo "Name: " $network.name
	echo "IP: " $network.ipaddress[0]
    $dnsSearchlist = ""
    echo "DNS Servers: " $network.DNSServerSearchOrder
    }