function Get-VMotionDuration {
<#
.SYNOPSIS
 This function returns the list of VMotions and Storage VMotions that occured during a specified period.
.DESCRIPTION
 This function returns the list of VMotions and Storage VMotions that occured during a specified period.
 You can specify if you want DRS triggered VMotions, manual VMotions or both.
.NOTES 
    File Name   : Get-VMotionDuration.ps1 
    Author      : David Vekemans - based on an original script from  Alan Renouf
 Date        : 17/05/2013
 Version     : 1.00
 History     : 1.00 - 17/05/2013 - Initial release
.LINK
 http://geekdav.blogspot.be/2013/05/powercli-list-vmotion-storage-vmotion.html
 http://www.virtu-al.net/2012/09/20/vmotion-and-svmotion-details-with-powercli/
.PARAMETER Start
 This parameter specify the start date and time of the period to analyze.
 It must be a DATETIME type.
.PARAMETER Finish
 This parameter specify the end date and time of the period to analyze.
 It must be a DATETIME type.
.PARAMETER DaysOfHistory
 This is another way to give the start and end time of the period to analyze.
 This parameter specify the number of days from now to define the start time.
 The default value is 1 day.
 It must be a INTEGER type.
.PARAMETER Type
 This parameter specify the types pf VMotion that you want to display.
 Possible values are :
 DRS - for DRS triggered VMotions
 MANUAL - for any user triggered VMotions
 BOTH - for all VMotions
 Default value is BOTH.
.OUTPUTS
 The output is a special PSobject (or a table of PSobjects) with the following properties:
  VMName             : name of the VM that was moved
  StartTime          : start time of the vMotion
  EndTime            : end time of the vMotion
  Duration           : duration of the vMotion
  SourceHost         : original host of the VM
  DestinationHost    : new host of the VM
  Type               : type of move vMotion or svMotion
  Trigger            : trigger of move : DRS or MANUAL
  User               : admin who triggered the move (only in case of MANUAL)
.INPUTS
 See the paramater list.
.EXAMPLE
 Get-VMotionDuration -DaysOfHistory 3 -Type DRS
 Name            : VM1
 StartTime       : 16/05/2013 11:32:38
 EndTime         : 16/05/2013 11:36:02
 Duration        : 00:03:23.9130000
 SourceHost      : host06.local
 DestinationHost : host07.local
 Type            : vMotion
 Trigger         : DRS
 User            :
  
 Name            : VM2
 StartTime       : 16/05/2013 11:32:39
 EndTime         : 16/05/2013 11:35:54
 Duration        : 00:03:15.8170000
 SourceHost      : host08.local
 DestinationHost : host09.local
 Type            : vMotion
 Trigger         : DRS
 User            :
  
  
 This command will list all the VMotions done by DRS during the last 3 days.
.EXAMPLE
 Get-VMotionDuration -Start (get-date).addhours(-4) -Finish (get-date) -Type Manual
 VMName          : VM3
 StartTime       : 17/05/2013 09:35:56
 EndTime         : 17/05/2013 09:36:53
 Duration        : 00:00:56.8360000
 SourceHost      : host08.local
 DestinationHost : host09.local
 Type            : vMotion
 Trigger         : MANUAL
 User            : DOMAIN\admin_user
  
  
 This command will list all the VMotions done by any admin during the last 4 hours.
.EXAMPLE
 Get-VMotionDuration 1
 VMName          : VM4
 StartTime       : 16/05/2013 11:41:02
 EndTime         : 16/05/2013 11:42:10
 Duration        : 00:01:08.1000010
 SourceHost      : host07.local
 DestinationHost : host09.local
 Type            : vMotion
 Trigger         : DRS
 User            :
  
 VMName          : VM5
 StartTime       : 17/05/2013 09:35:56
 EndTime         : 17/05/2013 09:36:53
 Duration        : 00:00:56.8360000
 SourceHost      : host08.local
 DestinationHost : host09.local
 Type            : vMotion
 Trigger         : MANUAL
 User            : DOMAIN\admin_user
  
 VMName          : VM6
 StartTime       : 17/05/2013 11:25:57
 EndTime         : 17/05/2013 11:27:06
 Duration        : 00:01:09.0300010
 SourceHost      : host09.local
 DestinationHost : host09.local
 Type            : svMotion
 Trigger         : MANUAL
 User            : DOMAIN\admin_user
  
  
 This command will list all the VMotions during the last day.
#>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True,
    ParameterSetName="1",
    Position=0)]
    [datetime]$Start,
 [Parameter(Mandatory=$True,
    ParameterSetName="1",
    Position=1)]
    [datetime]$Finish,
 [Parameter(Mandatory=$True,
    ParameterSetName="2",
    Position=0)]
    [int]$DaysOfHistory = 1,
 [Parameter(Mandatory=$False)]
 [ValidateSet("DRS","Manual","Both")]
    [String]$Type="Both"
  )
 
  begin {
   write-verbose "Starting function"
 $Type = $Type.toupper()
 if ($DaysOfHistory){
  $Now = Get-Date
  $Start = $Now.AddDays(-$DaysOfHistory)
  $Finish = $Now
 }
 else {
  $DaysOfHistory = ($Finish - $Start).Days
 }
 Write-verbose "Number of days of history is : $DaysOfHistory"
 Write-verbose "Start time is  : $Start"
 Write-verbose "Finish time is  : $Finish"
 $Maxsamples = ($DaysOfHistory+1)*50000
 Write-verbose "Max samples  : $Maxsamples"
 Write-Verbose "Type is : $type"
 $ResultDRS = @()
 $ResultManual = @()
  }
 
  process {
    write-verbose "Beginning process loop"
 $events = get-vievent -start $Start -finish $Finish -maxsamples $Maxsamples
 if (($Type -eq "DRS") -or ($Type -eq "BOTH")) {
  $relocates = $events | where {($_.GetType().Name -eq "TaskEvent") -and ($_.Info.DescriptionId -eq "Drm.ExecuteVMotionLRO")}
  foreach($task in $relocates){
   $tEvents = $events | where {$_.ChainId -eq $task.ChainId} | Sort-Object -Property CreatedTime
   if($tEvents.Count){
    $obj = New-Object PSObject -Property @{
       VMName = $tEvents[0].Vm.Name
       Type = &{if($tEvents[0].Host.Name -eq $tEvents[-1].Host.Name){"svMotion"}else{"vMotion"}}
       Trigger = "DRS"
       StartTime = $tEvents[0].CreatedTime
       EndTime = $tEvents[-1].CreatedTime
       Duration = New-TimeSpan -Start $tEvents[0].CreatedTime -End $tEvents[-1].CreatedTime
       SourceHost = $tEvents[0].Host.Name
       DestinationHost = $tEvents[-1].Host.Name
       User = $tEvents[0].UserName
       }
    $ResultDRS+=$obj  
   }
  }
 }
 if (($Type -eq "MANUAL") -or ($Type -eq "BOTH")) { 
  $relocates = $events | 
    where {($_.GetType().Name -eq "TaskEvent") -and (($_.Info.DescriptionId -eq "VirtualMachine.migrate") -or ($_.Info.DescriptionId -eq "VirtualMachine.relocate"))}
  foreach($task in $relocates){
   $tEvents = $events | where {$_.ChainId -eq $task.ChainId} | Sort-Object -Property CreatedTime
   if($tEvents.Count){
    $obj = New-Object PSObject -Property @{
       VMName = $tEvents[0].Vm.Name
       Type = &{if($tEvents[0].Host.Name -eq $tEvents[-1].Host.Name){"svMotion"}else{"vMotion"}}
       Trigger = "MANUAL"
       StartTime = $tEvents[0].CreatedTime
       EndTime = $tEvents[-1].CreatedTime
       Duration = New-TimeSpan -Start $tEvents[0].CreatedTime -End $tEvents[-1].CreatedTime
       SourceHost = $tEvents[0].Host.Name
       DestinationHost = $tEvents[-1].Host.Name
       User = $tEvents[0].UserName
       }
    $ResultMANUAL+=$obj
   }
  }
 } 
  }
   
  end {
   write-verbose "Ending function"
 if ($Type -eq "DRS") {$Result = $ResultDRS}
 if ($Type -eq "MANUAL") {$Result = $ResultMANUAL}
 if ($Type -eq "BOTH") {$Result = $ResultDRS + $ResultMANUAL}
 $Result | Select-Object VMName,StartTime,EndTime,Duration,SourceHost,DestinationHost,Type,Trigger,User | Sort-Object -Property StartTime
  }
}
function Get-VIEventPlus {
<#   
.SYNOPSIS  Returns vSphere events    
.DESCRIPTION The function will return vSphere events. With
    the available parameters, the execution time can be
   improved, compered to the original Get-VIEvent cmdlet. 
.NOTES  Author:  Luc Dekens   
.PARAMETER Entity
   When specified the function returns events for the
   specific vSphere entity. By default events for all
   vSphere entities are returned. 
.PARAMETER EventType
   This parameter limits the returned events to those
   specified on this parameter. 
.PARAMETER Start
   The start date of the events to retrieve 
.PARAMETER Finish
   The end date of the events to retrieve. 
.PARAMETER Recurse
   A switch indicating if the events for the children of
   the Entity will also be returned 
.PARAMETER User
   The list of usernames for which events will be returned 
.PARAMETER System
   A switch that allows the selection of all system events. 
.PARAMETER ScheduledTask
   The name of a scheduled task for which the events
   will be returned 
.PARAMETER FullMessage
   A switch indicating if the full message shall be compiled.
   This switch can improve the execution speed if the full
   message is not needed.   
.EXAMPLE
   PS> Get-VIEventPlus -Entity $vm
.EXAMPLE
   PS> Get-VIEventPlus -Entity $cluster -Recurse:$true
#>
 
  param(
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$Entity,
    [string[]]$EventType,
    [DateTime]$Start,
    [DateTime]$Finish = (Get-Date),
    [switch]$Recurse,
    [string[]]$User,
    [Switch]$System,
    [string]$ScheduledTask,
    [switch]$FullMessage = $false
  )
 
  process {
    $eventnumber = 100
    $events = @()
    $eventMgr = Get-View EventManager
    $eventFilter = New-Object VMware.Vim.EventFilterSpec
    $eventFilter.disableFullMessage = ! $FullMessage
    $eventFilter.entity = New-Object VMware.Vim.EventFilterSpecByEntity
    $eventFilter.entity.recursion = &{if($Recurse){"all"}else{"self"}}
    $eventFilter.eventTypeId = $EventType
    if($Start -or $Finish){
      $eventFilter.time = New-Object VMware.Vim.EventFilterSpecByTime
    if($Start){
        $eventFilter.time.beginTime = $Start
    }
    if($Finish){
        $eventFilter.time.endTime = $Finish
    }
    }
  if($User -or $System){
    $eventFilter.UserName = New-Object VMware.Vim.EventFilterSpecByUsername
    if($User){
      $eventFilter.UserName.userList = $User
    }
    if($System){
      $eventFilter.UserName.systemUser = $System
    }
  }
  if($ScheduledTask){
    $si = Get-View ServiceInstance
    $schTskMgr = Get-View $si.Content.ScheduledTaskManager
    $eventFilter.ScheduledTask = Get-View $schTskMgr.ScheduledTask |
      where {$_.Info.Name -match $ScheduledTask} |
      Select -First 1 |
      Select -ExpandProperty MoRef
  }
  if(!$Entity){
    $Entity = @(Get-Folder -Name Datacenters)
  }
  $entity | %{
      $eventFilter.entity.entity = $_.ExtensionData.MoRef
      $eventCollector = Get-View ($eventMgr.CreateCollectorForEvents($eventFilter))
      $eventsBuffer = $eventCollector.ReadNextEvents($eventnumber)
      while($eventsBuffer){
        $events += $eventsBuffer
        $eventsBuffer = $eventCollector.ReadNextEvents($eventnumber)
      }
      $eventCollector.DestroyCollector()
    }
    $events
  }
}
 
function Get-MotionHistory {
<#   
.SYNOPSIS  Returns the vMotion/svMotion history    
.DESCRIPTION The function will return information on all
   the vMotions and svMotions that occurred over a specific
    interval for a defined number of virtual machines 
.NOTES  Author:  Luc Dekens   
.PARAMETER Entity
   The vSphere entity. This can be one more virtual machines,
   or it can be a vSphere container. If the parameter is a
    container, the function will return the history for all the
   virtual machines in that container. 
.PARAMETER Days
   An integer that indicates over how many days in the past
   the function should report on. 
.PARAMETER Hours
   An integer that indicates over how many hours in the past
   the function should report on. 
.PARAMETER Minutes
   An integer that indicates over how many minutes in the past
   the function should report on. 
.PARAMETER Sort
   An switch that indicates if the results should be returned
   in chronological order. 
.EXAMPLE
   PS> Get-MotionHistory -Entity $vm -Days 1
.EXAMPLE
   PS> Get-MotionHistory -Entity $cluster -Sort:$false
.EXAMPLE
   PS> Get-Datacenter -Name $dcName |
   >> Get-MotionHistory -Days 7 -Sort:$false
#>
 
  param(
    [CmdletBinding(DefaultParameterSetName="Days")]
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$Entity,
    [Parameter(ParameterSetName='Days')]
    [int]$Days = 1,
    [Parameter(ParameterSetName='Hours')]
    [int]$Hours,
    [Parameter(ParameterSetName='Minutes')]
    [int]$Minutes,
    [switch]$Recurse = $false,
    [switch]$Sort = $true
  )
 
  begin{
    $history = @()
    switch($psCmdlet.ParameterSetName){
      'Days' {
        $start = (Get-Date).AddDays(- $Days)
      }
      'Hours' {
        $start = (Get-Date).AddHours(- $Hours)
      }
      'Minutes' {
        $start = (Get-Date).AddMinutes(- $Minutes)
      }
    }
    $eventTypes = "DrsVmMigratedEvent","VmMigratedEvent"
  }
 
  process{
    $history += Get-VIEventPlus -Entity $entity -Start $start -EventType $eventTypes -Recurse:$Recurse |
    Select CreatedTime,
    @{N="Type";E={
        if($_.SourceDatastore.Name -eq $_.Ds.Name){"vMotion"}else{"svMotion"}}},
    @{N="UserName";E={if($_.UserName){$_.UserName}else{"System"}}},
    @{N="VM";E={$_.VM.Name}},
    @{N="SrcVMHost";E={$_.SourceHost.Name.Split('.')[0]}},
    @{N="TgtVMHost";E={if($_.Host.Name -ne $_.SourceHost.Name){$_.Host.Name.Split('.')[0]}}},
    @{N="SrcDatastore";E={$_.SourceDatastore.Name}},
    @{N="TgtDatastore";E={if($_.Ds.Name -ne $_.SourceDatastore.Name){$_.Ds.Name}}}
  }
 
  end{
    if($Sort){
      $history | Sort-Object -Property CreatedTime
    }
    else{
      $history
    }
  }
}