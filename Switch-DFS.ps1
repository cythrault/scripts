param (
    [string]$source = "Holt",
	[string]$dest = "Berri"
    #[string]$source = "Berri",
	#[string]$dest = "Holt"
	)

Function Get-TargetStatus {
	$dfspaths = Get-DfsnRoot bnquebec.ca | %{ Get-DfsnFolder -Path "$($_.Path)\*" }
	$srvberri="pbwfsp0t"
	$srvholt="phwfsp0k"
	$i=0
	$total = $dfspaths.Count
	$act = "Scanning DFS Targets"
	$dfspaths | %{
		$i++
		Write-Progress -Activity $act -PercentComplete $( $i / $total * 100 ) -Status "$i of $total"
		$path = $_.Path
		$targets = Get-DfsnFolderTarget -Path $path
		if ( ($targets.TargetPath -Match $srvberri) -and ($targets.TargetPath -Match $srvholt) ) {
			$berri = Get-DfsnFolderTarget $path | ?{$_.TargetPath -match $srvberri}
			$holt = Get-DfsnFolderTarget $path | ?{$_.TargetPath -match $srvholt}
			[pscustomobject]@{Path=$path; Berri=$berri.State; Holt=$holt.State; BerriTarget=$berri.TargetPath; HoltTarget=$holt.TargetPath}
		}
	}
	Write-Progress -Activity $act -Completed
}

Write-Host -ForegroundColor Black -BackgroundColor Green "Checking current state."
$dfstargets = Get-TargetStatus

if ( ($dfstargets.$dest -NotMatch "Online") -or ($dfstargets.$source -NotMatch "Offline") ) {
	$dfstargets | %{
		$sourceTarget = "$($source)Target"
		$destTarget = "$($dest)Target"
		Write-Host -ForegroundColor Black -BackgroundColor Green "Setting target for $($PSItem.Path) on $($dest) Online"
		Set-DfsnFolderTarget -Path $PSItem.Path -TargetPath $PSItem.$destTarget -State Online -Confirm:$false
		Write-Host -ForegroundColor Black -BackgroundColor Green "Setting target for $($PSItem.Path) on $($source) Offline"
		Set-DfsnFolderTarget -Path $PSItem.Path -TargetPath $PSItem.$sourceTarget -State Offline -Confirm:$false
	}
	Write-Host -ForegroundColor Black -BackgroundColor Green "Switch completed, checking new state."
	$dfstargets = Get-TargetStatus
}

if ( ($dfstargets.$dest -Match "Online") -and ($dfstargets.$source -Match "Offline") ) {
	Write-Host -ForegroundColor Black -BackgroundColor Green "All DFS targets are online on" $dest
	Write-Host -ForegroundColor Black -BackgroundColor Green "All DFS targets are offline on" $source
} else {
	Write-Host -ForegroundColor Black -BackgroundColor DarkRed "Switch incomplete."
}