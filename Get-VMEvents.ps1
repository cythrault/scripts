param (
    [datetime]$beginTime = (Get-Date).Adddays(-60),
    [datetime]$endTime = (Get-Date)
)

$eventnumber = 1000 # max
$report = @()

$eventTypes = @{}
$eventTypes["vmClonedEvent"] = "vmClonedEvent"
$eventTypes["vmConnectedEvent"] = "vmConnectedEvent"
$eventTypes["vmCreatedEvent"] = "vmCreatedEvent"
$eventTypes["vmDeployedEvent"] = "vmDeployedEvent"
$eventTypes["vmDisconnectedEvent"] = "vmDisconnectedEvent"
$eventTypes["vmDiscoveredEvent"] = "vmDiscoveredEvent"
$eventTypes["vmMigratedEvent"] = "vmMigratedEvent"
$eventTypes["vmRegisteredEvent"] = "vmRegisteredEvent"
$eventTypes["vmRelocatedEvent"] = "vmRelocatedEvent"
$eventTypes["vmRemovedEvent"] = "vmRemovedEvent"
$eventTypes["vmRenamedEvent"] = "vmRenamedEvent"

$serviceInstance = get-view ServiceInstance
$eventMgr = Get-View eventManager

$efilter = New-Object VMware.Vim.EventFilterSpec
$efilter.time = New-Object VMware.Vim.EventFilterSpecByTime
$efilter.time.beginTime = $beginTime
$efilter.time.endtime = $endTime
$ecollectionImpl = Get-View ($eventMgr.CreateCollectorForEvents($efilter))
$ecollection = $ecollectionImpl.ReadNextEvents($eventnumber)

while($ecollection -ne $null){
    foreach($event in $ecollection){
        if($eventTypes[$event.gettype().name] -ne $null){
            $row = New-Object PSObject
            $row | Add-Member -name CreatedTime -value $event.CreatedTime -memberType NoteProperty
            $row | Add-Member -name VMname -value $event.Vm.Name -memberType NoteProperty
            $row | Add-Member -name VMid -value $event.Vm.Vm -memberType NoteProperty
            $row | Add-Member -name EventType -value $eventTypes[$event.gettype().name] -memberType NoteProperty
            $row | Add-Member -name FullFormattedMessage -value $event.FullFormattedMessage -memberType NoteProperty
            $row | Add-Member -name Host -value $event.Host.Name -memberType NoteProperty
            $row | Add-Member -name User -value $event.UserName -memberType NoteProperty
			Write-Output $row
        }
    }
	$ecollection = $ecollectionImpl.ReadNextEvents($eventnumber)
}

$ecollectionImpl.DestroyCollector()