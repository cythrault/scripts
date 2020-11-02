##global variables############################
$dnsServerIps = @("10.254.8.72","10.252.8.72")
$logfile = "DcDns.csv"
$failsafe = "OU=WINRMTest,OU=Svr,OU=Cmp,OU=GNV,DC=gcg,DC=local" #comment line to remove failsafe
##############################################
$scriptLocation = Split-Path $script:MyInvocation.MyCommand.Path
$outcsv = $scriptLocation + "\" + $logfile
$logobj = @()

function Get-ComputerSite($ComputerName)
{
   $site = nltest /server:$ComputerName /dsgetsite 2>$null
   if($LASTEXITCODE -eq 0){ $site[0] }
}

####################get list of sites ######################
$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
$siteContainerDN = ("CN=Sites," + $configNCDN)
$liste = @()
$sites = Get-ADObject -SearchBase $siteContainerDN -filter { objectClass -eq "site" -and name -like "ca*" -and description -notlike "focus*" -and description -notlike "pb*"} -properties "siteObjectBL", "location", "description" | select Name, Location, Description
foreach($site in $sites){
    $serverContainerDN = "CN=Servers,CN=" + $($site.name) + "," + $siteContainerDN
    $ErrorActionPreference = "silentlycontinue"
    $obj = Get-ADObject -SearchBase $serverContainerDN -filter {objectClass -eq "Server"} -SearchScope OneLevel  -Properties "DNSHostName", "Description" | select name, dnshostname
    $ErrorActionPreference = "Continue"
    if($obj -eq $null){
        $nouvelobj = new-object psobject | select siteName, dnshostname, name
        $nouvelobj.sitename = $site.name
        $liste += $nouvelobj
    }
    else{
        foreach($ob in $obj){
            $nouvelobj = new-object psobject | select siteName, dnshostname, name
            $nouvelobj.sitename = $site.name
            $nouvelobj.dnshostname = $ob.dnshostname
            $nouvelobj.name = $ob.name
            $liste += $nouvelobj
        }
    }

}
$liste = $liste | ?{($_.dnshostname -ne $null -and $_.name -ne $null) -or ($_.dnshostname -eq $null -and $_.name -eq $null)}
##############################################################
function get-DomainControllerFromSite ($siteName){
$adlist = @()
    $adlist += $liste | ?{$_.sitename -eq $siteName}
    if($adlist.count -gt 1){
        $listeadcontrollers = @()
        foreach($adcontroller in $adlist){
        $ErrorActionPreference = "silentlyContinue"
            $ad = (Get-ADDomainController -Identity $adcontroller.dnshostname -ErrorAction SilentlyContinue)
            $ErrorActionPreference = "Continue"
            if($ad -ne $null){
                $listeadcontrollers += $ad
            }
        }
        
        $listServers = @()
        $listServers += ($listeadcontrollers |? {$_.operatingsystem -like "*2012*"})
        $listServers += ($listeadcontrollers |? {$_.operatingsystem -like "*2008*"})
        $listServers += ($listeadcontrollers |? {$_.operatingsystem -like "*2003*"})
        $testconnect = $false
        $i = 0
        while(($testconnect -ne $true) -and ($i -lt $listServers.count)){
        $testconnect = Test-Connection -ComputerName $listServers[$i].hostname -Quiet
        $i++
        }
        return $listServers[$i-1].hostname
    }
    elseif($adlist.count -eq 1){
        return $adlist[0].dnshostname
    }
    else{
        return $null
    }
}
# end function get-domaincontrollerfromsite####################

$searchbase = "DC=gcg, DC=local"
$servers = Get-ADComputer -SearchBase $searchbase -Filter {operatingsystem -like "*server*"} -Properties operatingsystem
if($failsafe -ne $null){
#$servers = Get-ADComputer -SearchBase $failsafe -filter * -Properties operatingsystem
$servers = get-adcomputer -filter {name -like "sv026dc01"} -properties operatingsystem
}
$totalcount = $servers.count
$list2012 = @()
$listOther = @()
$host.EnterNestedPrompt()
#we want to go throught the whole server list
foreach($server in $servers){
    #we can only do the 2012 servers
    if($server.operatingsystem -like "*server*2012*"){
        $list2012 += $server
    }else{
        $listOther += $server
    }
}
$cpt = 1
#go through the to-do list for the 2012 servers+
foreach($server in $list2012){
    write-host "doing $cpt out of $totalcount"
    $site = Get-ComputerSite($server.name)
    $dc = get-DomainControllerFromSite($site)

    #if site has a DC
    if($dc -ne $null){
      $dcIp =  ([System.Net.Dns]::GetHostAddresses($dc)).ipaddresstostring
    if($dcIp.count -gt 1){$dcip = $dcip[0]}
      $ErrorActionPreference = "silentlyContinue"
      $adobj = (Get-ADDomainController -Identity $server -ErrorAction SilentlyContinue)
      $erroractionpreference = "continue"
        if($adobj -ne $null){
            $tempIps = $dnsServerIps + "127.0.0.1"
        }
        else{
            $tempIps = $dcIp + " " + $dnsServerIps
            $tempIps = $tempIps.split(" ")
        } 
    }
    #site has no DC
    else{
        $tempIps = $dnsServerIps
    }
    
     
    #TODO log what will be changed#
    $objlogfile = new-object psobject | select servername, os, site, isDC, hasDcOnSite, oldIps, newIps
    $objlogfile.os = $server.operatingsystem
    $objlogfile.site = $site
    $objlogfile.servername = $server.name
    $ErrorActionPreference = "SilentlyContinue"
    $adobj = (Get-ADDomainController -Identity $($server.name))
    if($adobj -ne $null){$objlogfile.isDC = "yes"}else{$objlogfile.isDC = "no"}
    $errorActionPreference = "Continue"
    if($dc -ne $null){$objlogfile.hasDConSite = "yes"}else{$objlogfile.hasDConSite = "no"}
    $cimsession = new-cimsession -ComputerName $server.name -ErrorAction SilentlyContinue
    #if we can establish connection
    if($cimsession -ne $null){
    $objlogfile.oldIps = ((gwmi Win32_NetworkAdapterConfiguration -ComputerName $($server.name)).dnsserversearchorder | out-string).replace("`n",",")
    $objlogfile.newips = ("" + $tempips).replace(" ",",")
    #then apply changes
    Get-DnsClientServerAddress -CimSession $cimsession| ?{$_.serveraddresses -ne $null -and $_.addressfamily -eq 2 -and $_.InterfaceAlias -notlike "*isatap*"} | set-dnsclientserveraddress -serveraddresses $tempIps
    }
    else{
    $objlogfile.oldIps = "server offline or not responding"
    $objlogfile.newips = "server offline or not responding"
    }

    $logobj += $objlogfile
    $cpt++
}
#go through the to-do list for the other servers (2003, 2008)
foreach($server in $listOther){
write-host "doing $cpt out of $totalcount"
  $tempsIps = @()
    $site = Get-ComputerSite($server.name)
    $dc = get-DomainControllerFromSite($site)
    #validate if server is member of a site with DC
    if($dc -ne $null){
    $nicname = (invoke-command -ComputerName $($server.name) -ScriptBlock {gwmi win32_networkadapter | ?{$_.netconnectionstatus -eq 2}}).netconnectionID
    $dcIp =  ([System.Net.Dns]::GetHostAddresses($dc)).ipaddresstostring
    if($dcIp.count -gt 1){$dcip = $dcip[0]}

    #if server IS a dc
    $errorActionPreference = "SilentlyContinue"
    $isDC = Get-ADDomainController -Identity $server.name
    $ErrorActionPreference = "continue"
    if($isdc -ne $null){
        $tempsIps = $dnsServerIps + "127.0.0.1"
    }
    #else it is a member server
    else{
        $tempsIps = $dcIp + " "+ $dnsServerIps
        $tempsips = $tempsips.split(" ") 
    }
    
    #TODO log what will be changed######################################################################
    $objlogfile = new-object psobject | select servername, os, site, isDC, hasDcOnSite, oldIps, newIps
    $objlogfile.os = $server.operatingsystem
    $objlogfile.site = $site
    $objlogfile.servername = $server.name
    $ErrorActionPreference = "silentlycontinue"
    $dcTest = (Get-ADDomainController -Identity $server.name -ErrorAction SilentlyContinue)
    $ErrorActionPreference = "Continue"
    if($dcTest -ne $null){$objlogfile.isDC = "yes"}else{$objlogfile.isDC = "no"}
    if($dc -ne $null){$objlogfile.hasDConSite = "yes"}else{$objlogfile.hasDConSite = "no"}
    ## need to find a way to log the old IPs for 2003-2008
    $wmiconfigs = (gwmi Win32_NetworkAdapterConfiguration -ComputerName $server.name -ErrorAction SilentlyContinue)
    if($wmiconfigs -ne $null){
    $objlogfile.oldIps = (($wmiconfigs.dnsserversearchorder | out-string).replace("`n",",")).replace(" ","")
    $objlogfile.newIps = ("" + $tempsips).replace(" ",",")
    
    $nicname = '"' + $nicname + '"'
    if($server.operatingsystem -like "*2003*"){
    
    $command1 =  $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip set dns $nicname static $($tempsIps[0]) primary")
    $command2 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip add dns $nicname $($tempsIps[1])")
    $command3 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip add dns $nicname $($tempsIps[2])")
    }
    else{
    $command1 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip set dnsserver $nicname static $($tempsIps[0]) primary")
    $command2 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip add dnsserver $nicname $($tempsIps[1])")
    $command3 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip add dnsserver $nicname $($tempsIps[2])")
    }
    $r = invoke-command -ComputerName $server.name -ScriptBlock $command1
    $r = invoke-command -ComputerName $server.name -ScriptBlock $command2
    $r = invoke-command -ComputerName $server.name -ScriptBlock $command3
    }
    else{
    $objlogfile.oldIps = "server offline or not responding"
    $objlogfile.newIps = "server offline or not responding"
    }
    $logobj += $objlogfile
    ###end logs
    
    
}
    #server is in a site without DC
    else{

    $nicname = (invoke-command -ComputerName $($server.name) -ScriptBlock {gwmi win32_networkadapter | ?{$_.netconnectionstatus -eq 2}}).netconnectionID
    $tempsIps = $dnsServerIps

    #TODO log what will be changed######################################################################
    $objlogfile = new-object psobject | select servername, os, site, isDC, hasDcOnSite, oldIps, newIps
    $objlogfile.os = $server.operatingsystem
    $objlogfile.site = $site
    $objlogfile.servername = $server.name
    if((Get-ADDomainController -Identity $server -ErrorAction SilentlyContinue) -ne $null){$objlogfile.isDC = "yes"}else{$objlogfile.isDC = "no"}
    if($dc -ne $null){$objlogfile.hasDConSite = "yes"}else{$objlogfile.hasDConSite = "no"}
    $cimsession = new-cimsession -ComputerName $server.name
     ## need to find a way to log the old IPs for 2003-2008
    $wmiconfigs = (gwmi Win32_NetworkAdapterConfiguration -ComputerName $server.name -ErrorAction SilentlyContinue)
    if($wmiconfigs -ne $null){
    $objlogfile.oldIps = (($wmiconfigs.dnsserversearchorder | out-string).replace("`n",",")).replace(" ","")
    $objlogfile.newIps = ("" + $tempips).replace(" ",",")

    $nicname = '"' + $nicname + '"'

     if($server.operatingsystem -like "*2003*"){
    $command1 =  $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip set dns $nicname static $($tempsIps[0]) primary")
    $command2 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip add dns $nicname $($tempsIps[1])")
    }
    else{
    $command1 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip set dnsserver $nicname static $($tempsIps[0]) primary")
    $command2 = $ExecutionContext.InvokeCommand.NewScriptBlock("netsh interface ip add dnsserver $nicname $($tempsIps[1])")
    }
    
    invoke-command -ComputerName $server.name -ScriptBlock $command1
    invoke-command -ComputerName $server.name -ScriptBlock $command2
    
    }
    else{
    $objlogfile.oldIps = "server offline or not responding"
    $objlogfile.newIps = "server offline or not responding"
    }

    $logobj += $objlogfile
    }
    
    ###end logs

    
    $cpt++
}

$scriptLocation = Split-Path $script:MyInvocation.MyCommand.Path
$outcsv = $scriptLocation + "\" + $logfile
$logobj | export-csv -NoTypeInformation -Delimiter ';' -Path $outcsv
