param (
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$vm,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[switch]$Console = $false,
	[Parameter(Mandatory=$false)]
	[switch]$Full = $false
)

function RunScript {
	param ([string]$ScriptText)
	Write-Host -BackgroundColor Green -ForegroundColor Black "Executing: $ScriptText"
	Invoke-VMScript -ScriptType PowerShell -ScriptText $ScriptText -VM $template -GuestCredential $localcreds
}

Write-Host -BackgroundColor Green -ForegroundColor Black "Waiting for VM Tools to be available."
Do {Sleep 5; Write-Host "." -NoNewline} Until ( (Get-VM -Name $vm).ExtensionData.Guest.ToolsRunningStatus -eq "guestToolsRunning" )
Write-Host "done"

Write-Host -BackgroundColor Green -ForegroundColor Black "Setting credentials to connect."
$admin = "administrator"
$adminPwd = ConvertTo-SecureString "Scaphandre12" -AsPlainText -Force
$localcreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $admin, $adminPwd

if ($Full.IsPresent) {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Running full update on $vm."
	#Update-Tools -VM $vm
	$result = Invoke-VMScript -ScriptType PowerShell -ScriptText "c:\tools\scripts\PrepareTemplate.ps1" -VM $vm -GuestCredential $localcreds
} else {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Installing Microsoft Updates only on $vm."
	$result = RunScript 'Get-WUInstall -MicrosoftUpdate -AcceptAll -IgnoreReboot -Install -Verbose -Confirm:$false'
	if (!$result) {
		Write-Host -BackgroundColor Red -ForegroundColor Black "Failed to run script."
	} else {
		$result.ScriptOutput | Out-File "$vm-$date.log"
		if ($result.Contains("Found [0] Updates")) { Write-Warning "No updates. Skipping replication."} else { [bool]$repl = $true }
		if ($result.Contains("Reboot is required")) {
			Do {
				Write-Warning "Reboot necessary."
				RunScript 'Restart-Computer -Force'
				WaitforVM
				$result = RunScript 'Get-WUInstall -Install -AcceptAll -IgnoreReboot -Verbose -Confirm:$false'
				$result.ScriptOutput | Out-File -Append "$vm-$date.log"
			} Until ( !$result.Contains("Reboot is required") )
		}
		Write-Host -BackgroundColor Green -ForegroundColor Black "Completed executing script. Results in $vm-$date.log."
		Invoke-Item "$vm-$date.log"
	}
}