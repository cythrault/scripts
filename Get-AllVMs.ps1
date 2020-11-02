$vcenters = "pvavcs0a.bnquebec.ca","pvavcs0b.bnquebec.ca"

#Connect to the vCenter server defined above. Ignore certificate errors
Write-Host "Connecting to vCenter"
Connect-VIServer $vcenters -wa 0 | Out-Null

#Gather all VM's from vCenter
$vms = Get-VM | Sort

$i = 0; $startTime = Get-Date; $total = $vms.Count
foreach ($VM in $vms) {
	$i++
	$elapsedTime = New-TimeSpan -Start $startTime -End (Get-Date)
    $timeRemaining = New-TimeSpan -Seconds (($elapsedTime.TotalSeconds / ($i / $total)) - $elapsedTime.TotalSeconds)
	$act = "Gathering information for VMs"
	
	Write-Progress -CurrentOperation "Searching in Events" -Activity $act -PercentComplete $(($i / $total) * 100) -Status "$vm - $i of $total - Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -SecondsRemaining $timeRemaining.TotalSeconds
	
	$vmevents = Get-VIEvent $VM -MaxSamples([int]::MaxValue) | Where-Object {$_.FullFormattedMessage -like "Deploying*"} |Select CreatedTime, UserName, FullFormattedMessage
	if ($vmevents) { $type = "From Template" }
	if (!$vmevents) {
		$vmevents = Get-VIEvent $VM -MaxSamples([int]::MaxValue) | Where-Object {$_.FullFormattedMessage -like "Created*"} |Select CreatedTime, UserName, FullFormattedMessage
		$type = "From Scratch"
	}
	if (!$vmevents) {
		$vmevents = Get-VIEvent $VM -MaxSamples([int]::MaxValue) | Where-Object {$_.FullFormattedMessage -like "Clone*"} |Select CreatedTime, UserName, FullFormattedMessage
		$type = "Cloned"
	}
	if (!$vmevents) {
		$vmevents = Get-VIEvent $VM -MaxSamples([int]::MaxValue) | Where-Object {$_.FullFormattedMessage -like "Discovered*"} |Select CreatedTime, UserName, FullFormattedMessage
		$type = "Discovered"
	}
	if (!$vmevents) {
		$vmevents = Get-VIEvent $VM -MaxSamples([int]::MaxValue) | Where-Object {$_.FullFormattedMessage -like "* connected"} |Select CreatedTime, UserName, FullFormattedMessage
		$type = "Connected"
	}
	if (!$vmevents) { $type = $null }

	if ($vmevents) {
		$CreationDate = $vmevents[0].CreatedTime.ToString("dd/MM/yyyy")
		$Creator = $vmevents[0].Username
		$CreationMsg = $vmevents[0].FullFormattedMessage
	} else {
		$CreationDate = $null
		$Creator = $null
		$CreationMsg = "Not found in Events"
	}

	$elapsedTime = New-TimeSpan -Start $startTime -End (Get-Date)
    $timeRemaining = New-TimeSpan -Seconds (($elapsedTime.TotalSeconds / ($i / $total)) - $elapsedTime.TotalSeconds)
	Write-Progress -CurrentOperation "Writing Custom Object" -Activity $act -PercentComplete $(($i / $total) * 100) -Status "$vm - $i of $total - Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -SecondsRemaining $timeRemaining.TotalSeconds

	[pscustomobject]@{
		Name 			= $vm.Name
		PowerState 		= $vm.PowerState
		Guest			= $vm.Guest.OSFullName
		GuestId			= $vm.GuestId
		HardwareVersion = $vm.HardwareVersion
		Folder			= $vm.Folder
		NumCPU			= $vm.NumCPU
		MemoryGB		= [math]::Round($vm.MemoryGB, 0)
		UsedSpaceGB		= [math]::Round($vm.UsedSpaceGB, 1)
		ProvisionedSpaceGB	= [math]::Round($vm.ProvisionedSpaceGB, 1)
		Host			= $vm.VMHost.Name
		Cluster			= (Get-VMHost $vm.VMHost).Parent.Name
		Datacenter		= (Get-Datacenter -VMHost $vm.VMHost).Name
		vCenter			= [System.Net.Dns]::GetHostEntry((Get-VMHost $vm.VMHost |Get-View).Summary.ManagementServerIp).HostName
		DataStore		= (Get-Datastore -Id (Get-VM $vm).DatastoreIdList).Name
		CreationType	= $type
		CreationDate 	= $CreationDate
		Creator			= $Creator
		CreationMsg		= $CreationMsg
		Notes			= $vm.Notes
	}
}

Write-Progress -Activity $act -Completed

#|Export-Excel -FreezeTopRow -AutoSize -Path AllVMs-$(Get-Date -UFormat %Y%m%d).xlsx
