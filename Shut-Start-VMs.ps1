param(
	[Parameter(Mandatory=$true)]
	[ValidateSet('Shutdown','Startup')]
	[string] $Action,
	[Parameter(Mandatory=$true)]
	[string] $Bloc
)

Start-Transcript -Path $action-$bloc-$($(Get-Date).ToString("yyyyMMddHHMMss")).log
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
}
$null = $FileBrowser.ShowDialog()

try { $VMs = Import-Excel -Path $FileBrowser.FileName -HeaderName "Nom","Status","Host","Notes","Bloc" -Sheet $action | ?{$_.Bloc -eq $bloc} }
catch { Write-Host -Foreground Red "Need list of VMs."; throw $_  }

#$VMs = Import-Excel -Path "\\bnquebec.ca\Tous\Informatique et télécommunications\Processus\Gestion des entretiens préventifs\2019\Mai\Sequence\201905GB.xlsx"

if ( $global:DefaultVIServers.Count -ne 2 ) {
	Write-Host "Connecting to vCenters..."
	try {
		If ((Get-VICredentialStoreItem -Host pvavcs0a.bnquebec.ca -ErrorAction SilentlyContinue) -and (Get-VICredentialStoreItem -Host pvavcs0b.bnquebec.ca -ErrorAction SilentlyContinue)) {
			Connect-VIServer -Server pvavcs0a.bnquebec.ca, pvavcs0b.bnquebec.ca
		} Else {
			Connect-VIServer pvavcs0a.bnquebec.ca, pvavcs0b.bnquebec.ca -Credential (Get-Credential -Message "vCenter credentials") -SaveCredentials
		}
	}	
	catch { Write-Host -Foreground Red "Need connection to both vCenters."; throw $_ }
}

If ( $VMs.Count -eq 0 ) { Write-Host "Nothing to do."; exit }

Write-Host "Found $($VMs.Count) machines to $action for $bloc"
Read-Host -Prompt "Press any key to continue or CTRL+C to quit"

$i = 0
$trottle = 0
$total = $VMs.Count
$startTime = Get-Date
foreach ($vm in $VMs) {
	$i++
	$elapsedTime = New-TimeSpan -Start $startTime -End (Get-Date)
	$timeRemaining = New-TimeSpan -Seconds (($elapsedTime.TotalSeconds / ($i / $total)) - $elapsedTime.TotalSeconds)
	$act = "$action of $($vm.nom)"
	Write-Progress -Activity $act -PercentComplete $( $i / $total * 100 ) -SecondsRemaining $timeRemaining.TotalSeconds -Status "$i of $total - Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))"
	$mv = $vm.Nom
	$vm = Get-VM -Name $vm.nom -ErrorAction SilentlyContinue
	
	if ( !$vm ) { Write-Host -Foreground Red "$mv not found in vCenters."; continue}
	
	$vmdc = (Get-VM $mv -ErrorAction SilentlyContinue) | Get-Datacenter
	if ( $vmdc.Name -ne "Berri" ) { Write-Host -Foreground Black -Background DarkRed "$mv not in Berri."; $vm = $null}
	
	if ( ($vm) -and ($action -eq "Shutdown") ) {
		If ($vm.PowerState -eq "PoweredOn") {
			if ($vm.Guest.State -eq "Running"){
				Write-Host -Foreground Green "$vm.Name is $vm.PowerState and tools are running - shutting down guest."
				Shutdown-VMGuest -VM $vm -Confirm:$false -Verbose #-WhatIf 
			} else {
				Write-Host -Foreground Yellow "$vm.Name is $vm.PowerState and tools not running - stopping the vm."
				Stop-VM -VM $vm -Confirm:$false -Verbose #-WhatIf
			}
			if ( $bloc -eq "Bloc3J" ) { Write-Host "Bloc 3J: Sleeping 5 seconds between VMs"; Start-Sleep -Seconds 5 }
			$trottle++
			if ( $trottle -gt 4 ) {
				Write-Host "Trottling: Sleeping 2 seconds.";
				Start-Sleep -Seconds 2
				$trottle = 0
			}
		} Else {
			if ( $vm.Name ) { Write-Host -Foreground Red "$($vm.Name) is $($vm.PowerState)." }
		}
	}

	if ( ($vm) -and ($action -eq "Startup") ) {
		If ($vm.PowerState -eq "PoweredOff") {
			Write-Host -Foreground Green "$($vm.Name) is $($vm.PowerState). Starting."
			Start-VM $vm.Name -Confirm:$false -Verbose #-WhatIf
			if ( $bloc -eq "Bloc3J" ) { Write-Host "Bloc 3J: Sleeping 5 seconds between VMs"; Start-Sleep -Seconds 5 }
			$trottle++
			if ( $trottle -gt 4 ) {
				Write-Host "Trottling: Sleeping 2 seconds.";
				Start-Sleep -Seconds 2
				$trottle = 0
			}
		} Else {
			if ( $vm.Name ) { Write-Host -Foreground Red "$($vm.Name) is $($vm.PowerState)." }
		}
	}
}
Write-Progress -Activity $act -Completed

#post action checks
$startTime = Get-Date
if ( $action -eq "Startup" ) { $desiredstate = "PoweredOn" }
if ( $action -eq "Shutdown" ) { $desiredstate = "PoweredOff" }
do {
	Write-Host "Checking state of $($VMs.Count) VMs for desired state: $desiredstate."
	Start-Sleep -Seconds 5
	$MVs = @()
	$vm = $null
	foreach ($vm in $VMs) {
		$mv = $vm.Nom
		$vm = Get-VM -Name $vm.nom -ErrorAction SilentlyContinue
		if ( !$vm ) { Write-Host -Foreground Red "$mv not found in vCenters."; continue}
		If ( [string]$vm.PowerState -ne $desiredstate ) {
			$MVs += $vm
			Write-Host "$($vm.Name) is $($vm.PowerState). Tools are $($vm.Guest.State)."
		}
	}
	$elapsedTime = New-TimeSpan -Start $startTime -End (Get-Date)
	Write-Host -Foreground Yellow "[$($elapsedTime.ToString('hh\:mm\:ss'))] $($MVs.Count) VMs are not in desired state: $desiredstate."
} while ( $MVs )
Write-Host "Completed in $($elapsedTime.ToString('hh\:mm\:ss'))"
Stop-Transcript