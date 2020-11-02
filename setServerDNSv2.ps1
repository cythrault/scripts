##load servers to do from .csv
$csvfile = Read-Host "Please enter the csv path:"

$command = {
    param($serverobj)
    ## in job #### get os
    $os = $serverobj.os
    $logobj = new-object psobject | select servername, os, site, isDC, hasDcOnSite, oldIps, newIps
    $logobj.servername = $serverobj.servername
    $logobj.os = $serverobj.os
    $logobj.site = $serverobj.site
    $logobj.isDC = $serverobj.isDC
    $logobj.hasDcOnSite = $serverobj.hasDcOnSite
    $logobj.oldIps = $serverobj.oldIps
    $logobj.shouldbe = $serverobj.newIps  

    ##2012 server
    if($os -like "*2012*"){
        #set-config
        $cimsession = new-cimsession -ComputerName $logobj.servername -ErrorAction SilentlyContinue
        Get-DnsClientServerAddress -CimSession $cimsession| ?{$_.serveraddresses -ne $null -and $_.addressfamily -eq 2 -and $_.InterfaceAlias -notlike "*isatap*"} | set-dnsclientserveraddress -serveraddresses $($logobj.shouldbe)
        #validate Ip
        $logobj.newIps = ((gwmi Win32_NetworkAdapterConfiguration -ComputerName $($logobj.servername)).dnsserversearchorder | out-string).replace("`n",",")
    }
    ##not 2012 server
    else{
        $nicname = (invoke-command -ComputerName $($logobj.servername) -ScriptBlock {gwmi win32_networkadapter | ?{$_.netconnectionstatus -eq 2}}).netconnectionID
        $nicname = '"' + $nicname + '"'
        $tempsIps = $($logobj.shouldbe).split(",")
        #set-config
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
            $r = invoke-command -ComputerName $logobj.servername -ScriptBlock $command1
            $r = invoke-command -ComputerName $logobj.servername -ScriptBlock $command2
            $r = invoke-command -ComputerName $logobj.servername -ScriptBlock $command3
            #validate Ip change
            $logobj.newIps = ((gwmi Win32_NetworkAdapterConfiguration -ComputerName $($logobj.servername)).dnsserversearchorder | out-string).replace("`n",",")
    }
    return $logobj
}
    
if($csvfile -ne "" -and $csvfile -like "*.csv"){
    $csvfile = import-csv -Path $csvfile -Delimiter ';'
    if($csvfile -ne $null){
        ## start a job for every server, maximum of $maxjobs simultaneous jobs
        $cpt = 1
        $count = $csv.length
        $maxjobs = 10
        ##foreach line in csv
        foreach($item in $csv){
            write-host "$cpt out of $($servers.count)"
            while((get-job -state running).count -ge $maxjobs){
                sleep 2
            }
            $argument = $item
            ##start new job (servobj from csv)
            $job = Start-Job -ScriptBlock $command -ArgumentList $argument
            $job.name = "SetServer:$($item.servername)"
            $cpt++
        }
        ##receive job results
        ##get job results
        $done = $false
        $jobtimeout = 600 #seconds
        $finallist = @()
        while(!$done){
    
            $oldresults = $results
            $results = get-job -name "ScanServer:*"
            if($oldresults -ne $results){
                $startjobTime = get-date
            }
            #check for jobs done
            $resultsdone = $results | ?{$_.state -ne "Running"}
            foreach($result in $resultsdone){
                $resultdata = $null
                $resultdata = receive-job $result -ErrorAction SilentlyContinue
                $finallist += $resultdata
                remove-job $result
            }
            $jobTimeSpan = New-TimeSpan -Start $startjobtime -end $(Get-Date)
            # no results, jobs completed
            if($results.count -eq 0){
                $done = $true        
            }
            #if there were no changes for the past 10 minutes, cancel the remaining jobs
            elseif($jobTimeSpan.TotalSeconds -gt $jobtimeout){
                write-host "jobs took too long, cancelling remaining jobs"
                foreach($result in $results){
                    write-host "stopping job $($result.Name)"
                    $logobj = new-object psobject | select servername, os, site, isDC, hasDcOnSite, oldIps, newIps
                    $logobj.servername = $($result.name)
                    $logobj.os = $logobj.site = $logobj.isDC = $logobj.hasDcOnSite = $logobj.oldIps = $logobj.newIps = "N/A"
                    $finallist+= $logobj
                    stop-job $result
                    remove-job $result
                }
                $done = $true
           } #end elseif jobTimedOut
        sleep 2 #wait for jobs to catch on
        }#end job results
    }
    else{
        write-host "file doesn't exist, exiting"
    }
}
else{
    write-host "user didn't input a csv file, exiting"
}




