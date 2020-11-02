function Get-SnapshotSummary {
	param(
		$InputObject = $null
	)

	process {
		if ($InputObject -and $_) {
			throw 'ParameterBinderStrings\AmbiguousParameterSet'
			break
		} elseif ($InputObject) {
			$InputObject
		} elseif ($_) {
			$mySnaps = @()
			foreach ($snap in $_){
				$SnapshotInfo = Get-SnapshotExtra $snap
				$mySnaps += $SnapshotInfo
			}
			$mySnaps | Select VM,
				@{N="SnapName";E={[System.Web.HttpUtility]::UrlDecode($_.Name)}},
				@{N="DaysOld";E={[math]::round(((Get-Date) - $_.Created).Days,1)}},
				Creator,
				@{N="SizeGB";E={[math]::round($_.SizeGB -as [int],1)}},
				Created,
				Description -ErrorAction SilentlyContinue | Sort DaysOld
		} else {
			throw 'ParameterBinderStrings\InputObjectNotBound'
		}
	}
}

function Get-SnapshotTree{
	param($tree, $target)
	$found = $null

	foreach($elem in $tree){
		if($elem.Snapshot.Value -eq $target.Value){
			$found = $elem
			continue
		}
	}

	if($found -eq $null -and $elem.ChildSnapshotList -ne $null){
		$found = Get-SnapshotTree $elem.ChildSnapshotList $target
	}
	return $found
}

function Get-SnapshotExtra ($snap){
	$guestName = $snap.VM   # The name of the guest
	$tasknumber = 999	# Windowsize of the Task collector
	$taskMgr = Get-View TaskManager
	# Create hash table. Each entry is a create snapshot task
	$report = @{}
	$filter = New-Object VMware.Vim.TaskFilterSpec
	$filter.Time = New-Object VMware.Vim.TaskFilterSpecByTime
	$filter.Time.beginTime = (($snap.Created).AddSeconds(-5))
	$filter.Time.timeType = "startedTime"
	$collectionImpl = Get-View ($taskMgr.CreateCollectorForTasks($filter))
	$dummy = $collectionImpl.RewindCollector
	$collection = $collectionImpl.ReadNextTasks($tasknumber)

	while ($collection -ne $null) {
		$collection | where {$_.DescriptionId -eq "VirtualMachine.createSnapshot" -and $_.State -eq "success" -and $_.EntityName -eq $guestName} | %{
			$row = New-Object PsObject
			$row | Add-Member -MemberType NoteProperty -Name User -Value $_.Reason.UserName
			$vm = Get-View $_.Entity
			$snapshot = Get-SnapshotTree $vm.Snapshot.RootSnapshotList $_.Result
			if($snapshot){
				$key = $_.EntityName + "&" + ($snapshot.CreateTime.ToLocalTime().ToString())
				$report[$key] = $row
			} else {
				#Write-Host "No snapshot found for this event"
			}
		}
		$collection = $collectionImpl.ReadNextTasks($tasknumber)
	}
	$collectionImpl.DestroyCollector()
	# Get the guest's snapshots and add the user
	$snapshotsExtra = $snap | % {
		$key = $_.vm.Name + "&" + ($_.Created.ToString())
		if ( $report.ContainsKey($key) ){ $_ | Add-Member -MemberType NoteProperty -Name Creator -Value $report[$key].User }
		$_
	}
	$snapshotsExtra
}

# HTML formatting
$a  = "<style>"
$a += "BODY{background-color:white;}"
$a += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a += "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;foreground-color: black;background-color: LightBlue}"
$a += "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;foreground-color: black;background-color: white}"
$a += "</style>"

#Import-Module VMware.VimAutomation.Core
#Connect-VIServer pvavcs0a.bnquebec.ca,pvavcs0b.bnquebec.ca -user bnquebec\vclireadonly -password 4O6lX+q%

$date = Get-Date
$VMsSnaphots = Get-VM | Get-Snapshot
$body = $VMsSnaphots | Select-Object @{N='VM';E={$_.VM.Name}}, 
			Name,
			@{N="Datacenter"; E={ (Get-Datacenter -VM $_.VM.Name)}},
			@{N='DaysOld'; E={ [math]::round( ($date - $_.Created).TotalDays, 1) }},
			@{N='Creator'; E={ (Get-SnapshotExtra $_).Creator }},
			@{N="SizeGB"; E={ [math]::round( $_.SizeGB, 1) }},
			Created,
			Description | Sort-Object -Property Created | ConvertTo-HTML -head $a -body "<H2>VM Snapshot Report</H2>" | Out-String

$sender="vcenter@banq.qc.ca"
$recipients="martin.thomas@banq.qc.ca"
$subject="VM Snapshot Report"
$smtpserver="smtp.banq.qc.ca"
Send-MailMessage -To $recipients -BodyAsHtml -Body $body -From $sender -Subject $subject -SmtpServer $smtpserver