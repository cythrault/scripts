param (
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$Template,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[switch]$Console = $false,
	[Parameter(Mandatory=$false)]
	[switch]$Full = $false
)

function ToVM {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Converting template to VM."
	Set-Template -Template $(Get-Template -Location Berri -Name $template) -ToVM -Confirm:$false
	Write-Host -BackgroundColor Green -ForegroundColor Black "Starting VM."
	Start-VM -VM $template -Confirm:$false
	if ($console.IsPresent) {
		Write-Host -BackgroundColor Green -ForegroundColor Black "Connecting to console."
		Get-VM $template | Open-VMConsoleWindow
	}
}

function ToTemplate {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Shutting down Guest VM."
	Shutdown-VMGuest -VM $template -Confirm:$false
	Write-Host -BackgroundColor Green -ForegroundColor Black "Waiting for VM to be Powered Off."
	Do {
		$i++
		Sleep 5; Write-Host "." -NoNewline
		if ($i -gt 6) {
			Write-Host "."
			Write-Warning "Still waiting for VM to be Powered Off. Sending another shutdown request."
			Shutdown-VMGuest -VM $template -Confirm:$false
			$i=0
		}
	} Until ( (Get-VM $template).PowerState -eq "PoweredOff" )
	Write-Host "done"
	Write-Host -BackgroundColor Green -ForegroundColor Black "Converting VM to template."
	Set-VM -VM $template -ToTemplate -Confirm:$false
}

function Replicate {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Replicating template to secondary site."
	$destclu = Get-Datastore -Location Holt|?{$_.Name -like "*ISO-TEMPLATES"}
	$netpg = Get-VDPortgroup -Name DvPG-Gestion-332 | ?{$_.Datacenter -like "Holt"}
	$destfolder = Get-Folder -Location Holt -Name Templates
	$vmhost = (Get-VMHost -Location Berri|Get-Random)
	$destvmhost = (Get-VMHost -Location Holt|Get-Random)
	if (Get-Template -Location Holt -Name $template -ErrorAction Ignore) {Get-Template -Location Holt -Name $template|Remove-Template -DeletePermanently -Confirm:$false}
	New-VM -Name $template -Template $template -VMHost $vmhost
	Move-VM -VM (Get-VM $template) -Destination $destvmhost -Datastore $destclu -PortGroup $netpg -InventoryLocation $destfolder -DiskStorageFormat Thin
	Set-VM -VM $template -ToTemplate -Confirm:$false
}

function RunScript {
	param ([string]$ScriptText)
	Write-Host -BackgroundColor Green -ForegroundColor Black "Executing: $ScriptText"
	Invoke-VMScript -ScriptType PowerShell -ScriptText $ScriptText -VM $template -GuestCredential $localcreds
}

function WaitforVM {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Waiting for VM Tools to be available."
	Do {Sleep 5; Write-Host "." -NoNewline} Until ( (Get-VM -Name $template).ExtensionData.Guest.ToolsRunningStatus -eq "guestToolsRunning" )
	Write-Host "done"
}

# Install PS-Menu if not available
if (!(Get-Module -Name PS-Menu -ListAvailable)) {Write-Host -BackgroundColor Green -ForegroundColor Black "Installing PS-Menu.."; Install-Module PS-Menu -Force}

add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
  public bool CheckValidationResult(
  ServicePoint srvPoint, X509Certificate certificate,
  WebRequest request, int certificateProblem) {
  return true;
  }
}
"@

[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

$date = Get-Date -Format yyyyMMddhhmm

Write-Host -BackgroundColor Green -ForegroundColor Black "Locating template $template."

if ($template) {
	$template = Get-Template -Name $template -Location "Berri"
} else {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Please select template to update."
	$template = menu @(Get-Template -Location "Berri" | Sort)
}

if ([string]::IsNullOrEmpty($template)) {
	Write-Host -BackgroundColor Red -ForegroundColor Black "A template need to be specified."
	return
}

ToVM

Write-Host -BackgroundColor Green -ForegroundColor Black "Waiting for VM Tools to be available."
Do {Sleep 5; Write-Host "." -NoNewline} Until ( (Get-VM -Name $template).ExtensionData.Guest.ToolsRunningStatus -eq "guestToolsRunning" )
Write-Host "done"

Write-Host -BackgroundColor Green -ForegroundColor Black "Setting credentials to connect."
$admin = "administrator"
$adminPwd = ConvertTo-SecureString "Scaphandre12" -AsPlainText -Force
$localcreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $admin, $adminPwd

if ($Full.IsPresent) {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Running full update on $template."
	#Update-Tools -VM $template
	$result = Invoke-VMScript -ScriptType PowerShell -ScriptText "c:\tools\scripts\PrepareTemplate.ps1" -VM $template -GuestCredential $localcreds
} else {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Installing Microsoft Updates only on $template."
	$result = RunScript 'Get-WUInstall -Install -AcceptAll -IgnoreReboot -Verbose -Confirm:$false'
	if (!$result) {
		Write-Host -BackgroundColor Red -ForegroundColor Black "Failed to run script."
	} else {
		$result.ScriptOutput | Out-File "$template-$date.log"
		if ($result.Contains("Found [0] Updates")) { Write-Warning "No updates. Skipping replication."} else { [bool]$repl = $true }
		if ($result.Contains("Reboot is required")) {
			Do {
				Write-Warning "Reboot necessary."
				RunScript 'Restart-Computer -Force'
				WaitforVM
				$result = RunScript 'Get-WUInstall -Install -AcceptAll -IgnoreReboot -Verbose -Confirm:$false'
				$result.ScriptOutput | Out-File -Append "$template-$date.log"
			} Until ( !$result.Contains("Reboot is required") )
		}
		Write-Host -BackgroundColor Green -ForegroundColor Black "Completed executing script. Results in $template-$date.log."
		Invoke-Item "$template-$date.log"
	}
}

ToTemplate
if ($repl) {Replicate}

Write-Host -BackgroundColor Green -ForegroundColor Black "Completed!"