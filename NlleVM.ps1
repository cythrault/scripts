#Requires -Version 5

<#
.SYNOPSIS
	Deploy VM from command line.

.DESCRIPTION
	NlleVM  validate and execute the VM creation on vCenter.

.PARAMETER Name
	Specify the VM name to use.

.PARAMETER Description
	Specify the VM description to use.

.PARAMETER NbCPU
	Specify the number of CPUs to assign.

.PARAMETER NbRAM
	Specify the size in GB of RAM to assign.

.PARAMETER DiskFormat
	Specify the type of disk format to use (Thick or Thin.)

.PARAMETER Site
	Specify in which site to deploy the VM.

.PARAMETER Template
	Specify the template to use.

.PARAMETER DatastoreCluster
	Specify the Datastore to use.

.PARAMETER NetworkPG
	Specify the Network to use.

.PARAMETER WSUSGroup
	Specify the WSUS Group to use.

.PARAMETER Folder
	Specify the Folder to store the VM.

.PARAMETER IPAddress
	Specify the IP Address to use. Enter DHCP to use dynamic addressing.

.PARAMETER IPSubnetMask
	Specify the subnet mask to use if not 255.255.255.0.

.PARAMETER IPGateway
	Specify the gateway to use of not .1.

.PARAMETER IPDNSServers
	Specify the DNS servers to use. Specifiy one or more sperated my commas.

.PARAMETER ADDomain
	Specify the domain to join.

.PARAMETER orgName
	Specify the organisation name to use in Windows.

.PARAMETER domainCredentials,
	Specify the credentials to use when joining a domain.

.PARAMETER vCenters
	Specify the vCenter servers to connect.

.NOTES
	Version:	1.3
	Author:		Martin Thomas

.EXAMPLE
	NlleVM
	NlleVM -Name Test2
	NlleVM -Name Test2 -NbRAM 16 -Site Berri -Template W2K19-DATACTR-GUI-TPL-1.2 -Folder Test
	NlleVM -Name test2 -NbCPU 4 -NbRAM 16 -Site Berri -Template W2K19-DATACTR-GUI-TPL-1.2 -Folder Test -NetworkPG DvPG-Infra-336
	NlleVM -Name PVWADS0z -ADDomain "banqpublic.ca" -IPDNSServer1 172.16.50.25 -IPDNSServer2 172.16.50.26 -NetworkPG DvPG-Public-650 -IPAddress 172.16.50.254 -NbCPU 2 -NbRAM 8 -Site Berri -Template W2K19-DATACTR-CORE-TPL-1.2 -Folder AD
	NlleVM -Name avwxen0t -NetworkPG DvPG-Infra-336 -IPAddress 10.9.36.90 -NbCPU 4 -NbRAM 16 -Site Berri -Template W2K16-DATACTR-GUI-TPL-1.2 -Folder Citrix
#>

[CmdletBinding(SupportsShouldProcess)]
param (
	[Parameter(Mandatory=$true, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[ValidatePattern("^(?![0-9]{1,15}$)[a-zA-Z0-9-]{1,15}$")]
	[string]$Name,
	[Parameter(Mandatory=$true, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$Description,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[int]$NbCPU = 1,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[int]$NbRAM = 4,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$DiskFormat = "Thin",
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$Site,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$Template,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$Folder,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$DatastoreCluster,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$NetworkPG,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$WSUSGroup,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[ValidateScript({$_ -match [IPAddress]$_})]
	[string]$IPAddress,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[ValidateScript({$_ -match [IPAddress]$_})]
	[string]$IPSubnetMask = "255.255.255.0",
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[ValidateScript({$_ -match [IPAddress]$_})]
	[string]$IPGateway,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[ValidateScript({$_ -match [IPAddress]$_})]
	[array]$IPDNSServers, # = ("10.9.35.11", "10.9.35.62"),
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$ADDomain = "bnquebec.ca",
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$OUpath,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$admin = "administrator",
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$adminPwd = "Scaphandre12",
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string]$orgName = "BAnQ",
	[Parameter(Mandatory=$False, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[System.Management.Automation.PSCredential]
	$domainCredentials,
	[Parameter(Mandatory=$false, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[array]$vCenters = ("pvavcs0a.bnquebec.ca", "pvavcs0b.bnquebec.ca")
)

# Required to use Invoke-VMScript
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {public bool CheckValidationResult(ServicePoint srvPoint, X509Certificate certificate, WebRequest request, int certificateProblem) {return true;}}
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

$localpwd = ConvertTo-SecureString $adminPwd -AsPlainText -Force
function RunVMScript {
	param ([string]$ScriptText)
	Write-Host -BackgroundColor Green -ForegroundColor Black "Executing: $ScriptText"
	Invoke-VMScript -ScriptType PowerShell -ScriptText $ScriptText -VM $nouvellevm -GuestUser $admin -GuestPassword $localpwd
}

#$ConfirmPreference = "High"

# Install PS-Menu if not available
if (!(Get-Module -Name PS-Menu -ListAvailable)) {Write-Host -BackgroundColor Green -ForegroundColor Black "Installing PS-Menu.."; Install-Module PS-Menu -Force}
Import-Module PS-Menu # -Verbose

# Connect to vCenter
if ( ((Get-PowerCLIConfiguration -Scope User).DefaultVIServerMode -ne "Multiple") -and ($vCenters.Count -gt 1) ) {
	Write-Warning "PowerCLI configuration set to single connection and multiple vCenters were specified: configuring for multiple vCenters."
	Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false
}

# determining if we have a connection to all specified vCenter server(s) specified
$missingVC = @()
if ( $global:DefaultVIServers ) {
	ForEach ($vCenter in $vCenters) {
		if( ($global:DefaultVIServers.Name -notcontains $vCenter)) { $missingVC += $vcenter }
	}
} else { $missingVC = $vCenters }

if ( $missingVC ) {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Connecting to vCenter(s)"
	$viuser = "$env:UserDomain\$env:UserName"
	
	# checking if we have credentials for missing connections
	ForEach ($vCenter in $missingVC) {
		Write-Warning "Checking for stored credentials for $vCenter"
		if (Get-VICredentialStoreItem -Host $vCenter -User $viuser -ErrorAction Ignore) {
			Write-Host -ForegroundColor Green "Got credentials for $vCenter"
		} else {
			Write-Warning "Missing credentials for $vCenter"
			$vipwd = (Get-Credential -UserName $viuser -Message "Enter password for $vCenter").Password
			$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($vipwd)
			$vipwd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
			New-VICredentialStoreItem -Host $vCenter -User $viuser -Password $vipwd -Verbose
		}
	}
	# attempting to connect to missing servers
	try { Connect-VIServer -Server $missingVC }
	catch {
		Write-Host -BackgroundColor Red -ForegroundColor Black "Could not connect to specified vCenter server(s.)"
		Write-Error $_
		return
	}
}

if ( (($site -eq "Holt") -and ($global:DefaultVIServers.Name -notcontains "pvavcs0b.bnquebec.ca")) -or
(($site -eq "Berri") -and ($global:DefaultVIServers.Name -notcontains "pvavcs0a.bnquebec.ca"))
) { Write-Host -BackgroundColor Red -ForegroundColor Black "Site/vCenter mismatch: $site/$global:DefaultVIServers"; return }

# check provded values or ask for mandatory variables
if ($IPAddress){ if (Test-Connection $IPAddress -count 1 -Quiet) { Write-Host -BackgroundColor Red -ForegroundColor Black "$IPAddress is not free."; return }}

if (!$IPDNSServers){ 
	Switch ($site)
	{
		"Berri" { [array]$IPDNSServers = ("10.9.35.11", "10.9.35.62") }
		"Holt" { [array]$IPDNSServers = ("10.9.35.62", "10.9.35.11") }
	}
}

Write-Host -BackgroundColor Green -ForegroundColor Black "Testing credentials for $ADDomain."
if ((!$domainCredentials) -or ($domainCredentials.UserName.Split("\")[0] -notlike $addomain.Split(".")[0]) ) {
	$domainCredentials = Get-Credential -Message "Enter credentials for $ADDomain"
	if (!$domainCredentials) {
		Write-Host -BackgroundColor Red -ForegroundColor Black "Empty credentials."
		return
	}
}
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain
$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $ct, $ADDomain
#$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
if (!$DS.ValidateCredentials($domainCredentials.UserName.Split("\")[1], $domainCredentials.GetNetworkCredential().Password)) {
	Write-Host -BackgroundColor Red -ForegroundColor Black "Unable to authenticate with provided credentials for $ADDomain."
	return
}

try {$adcomputer=Get-ADComputer -Server $ADDomain -Identity $name -Server $ADDomain -Credential $domainCredentials} catch {}
if ($adcomputer) {Write-Host -BackgroundColor Red -ForegroundColor Black "$name already exist in $ADDomain."; return}

if (!$OUpath) {
	$OUpath = "OU=Serveurs,OU=1.Gestion"
	ForEach ( $dc in $($addomain.Split(".")) ) {
		$OUpath += ",DC=$dc"
	}
}

try {$adou=Get-ADOrganizationalUnit -Identity $OUpath -Server $ADDomain -Credential $domainCredentials} catch {}
if (!$adou) {Write-Host -BackgroundColor Red -ForegroundColor Black "$OUpath does not exist in $ADDomain."; return}

if (Get-VM -Name $name -ErrorAction Ignore) { Write-Host -BackgroundColor Red -ForegroundColor Black "$name already exist."; return }
if ($site) { $site = Get-Datacenter -Name $site } else { Write-Host -BackgroundColor Green -ForegroundColor Black "Select site";  $site = menu @(Get-Datacenter | Sort) }
if ($template) { $template = Get-Template -Name $template -Location $site } else { Write-Host -BackgroundColor Green -ForegroundColor Black "Select template"; $template = menu @(Get-Template -Location $site | Sort) }
if ($folder) { $folder = Get-Folder -Name $folder -Type VM -Location $site } else { Write-Host -BackgroundColor Green -ForegroundColor Black "Select VM Folder"; $folder = menu @(Get-Folder -Type VM -Location $site | Sort) }
if ($DatastoreCluster) { $DatastoreCluster = Get-DatastoreCluster -Name $DatastoreCluster -Location $site } else { $DatastoreCluster = Get-DatastoreCluster -Location $site | ?{$_.Name -like "*OS"} }
if ($NetworkPG) { $NetworkPG = Get-VDPortgroup -Name $NetworkPG | ?{$_.Datacenter -like $site} } else { Write-Host -BackgroundColor Green -ForegroundColor Black "Select Network Port Group"; $NetworkPG = menu @(Get-VDPortgroup | ?{$_.Datacenter -like $site}) }

$wsus = Get-WsusServer -Name pvwbat0b -PortNumber 8530
$WSUSGroupTarget = $wsus.GetComputerTargetGroups() | ?{$_.Name -eq $WSUSGroup}
if ($WSUSGroupTarget -eq $null) {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Select WSUS Group"
	$wsusGroups = $wsus.GetComputerTargetGroups().Name
	$WSUSGroup = menu @( $wsusGroups )
}

Write-Host -BackgroundColor Green -ForegroundColor Black "Locating an ESX host in $site."
$vmhost = Get-Cluster -Location $site | Get-VMHost | Get-Random

#Write-Host -BackgroundColor Green -ForegroundColor Black "Locating least used (RAM) ESX host."
#$vmhost = $(
#	$metrics = "mem.usage.average"
#	$esx = Get-Cluster -Location $site | Get-VMHost
#	$stats = Get-Stat -Entity $esx -Stat $metrics -Realtime -MaxSamples 1
#	$avg = $stats | Measure-Object -Property Value -Average | Select -ExpandProperty Average
#	$stats | where{$_.Value -lt $avg} | Get-Random | Select -ExpandProperty Entity
#)

if (
	[string]::IsNullOrWhiteSpace($site) -or
	[string]::IsNullOrWhiteSpace($vmhost) -or
	[string]::IsNullOrWhiteSpace($template) -or
	[string]::IsNullOrWhiteSpace($folder) -or
	[string]::IsNullOrWhiteSpace($DatastoreCluster) -or
	[string]::IsNullOrWhiteSpace($NetworkPG) -or
	[string]::IsNullOrWhiteSpace($NbCPU) -or
	[string]::IsNullOrWhiteSpace($NbRAM) -or
	[string]::IsNullOrWhiteSpace($DiskFormat)
) { Write-Host -BackgroundColor Red -ForegroundColor Black "One or more values are not correct."; return }

try {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Creating OS Customization Spec."
	if (Get-OSCustomizationSpec -Name $name -ErrorAction Ignore) {
		Write-Warning "An OS Customization Spec already exist for $name. Clearing before proceeding."
		Remove-OSCustomizationSpec -CustomizationSpec $name -Confirm:$false
	}
	$OSSpec = New-OSCustomizationSpec -OSType "Windows" -Name $name -Type NonPersistent `
-Domain $addomain `
-DomainCredentials $domainCredentials `
-FullName $admin `
-AdminPassword $adminPwd `
-AutoLogonCount 1 `
-OrgName $orgName `
-TimeZone 035 `
-ChangeSid `
-ErrorAction Stop `
-Confirm:$false

	$NicSpecsProperties = @{OSCustomizationNicMapping = Get-OSCustomizationNicMapping -OSCustomizationSpec $OSSpec }
	if (!$IPAddress){
		Write-Warning "IP not provided, using DHCP"
		$NicSpecsProperties.IpMode = "UseDHCP"
	} else {
		if (!$IPGateway) {
			$IPByte = $IPAddress.Split(".")
			$IPGateway = $IPByte[0] + "." + $IPByte[1] + "." + $IPByte[2] + ".1"
		}
		$NicSpecsProperties.IpMode = "UseStaticIP"
		$NicSpecsProperties.IpAddress = $IPAddress
		$NicSpecsProperties.SubNetMask = $IPSubnetMask
		$NicSpecsProperties.DefaultGateway = $IPGateway
		$NicSpecsProperties.dns = $IPDNSServers
	}
	Set-OSCustomizationNicMapping @NicSpecsProperties -Confirm:$false
}
catch {
	Write-Host -BackgroundColor Red -ForegroundColor Black "Error creating OS Customization Spec."
	Write-Error $_
	return
}

Write-Host -BackgroundColor Green -ForegroundColor Black "New VM specifications:"
Write-Host -ForegroundColor Gray "Name: " -NoNewline; Write-Host -ForegroundColor Green $name
Write-Host -ForegroundColor Gray "Description: " -NoNewline; Write-Host -ForegroundColor Green $Description
Write-Host -ForegroundColor Gray "Number of CPU(s): " -NoNewline; Write-Host -ForegroundColor Green $NbCPU
Write-Host -ForegroundColor Gray "Memory Size (GB): " -NoNewline; Write-Host -ForegroundColor Green $NbRAM
Write-Host -ForegroundColor Gray "Site: " -NoNewline; Write-Host -ForegroundColor Green $site
Write-Host -ForegroundColor Gray "ESX Host: " -NoNewline; Write-Host -ForegroundColor Green $vmhost
Write-Host -ForegroundColor Gray "Datastore Cluster: " -NoNewline; Write-Host -ForegroundColor Green $DatastoreCluster
Write-Host -ForegroundColor Gray "Template: " -NoNewline; Write-Host -ForegroundColor Green $template
Write-Host -ForegroundColor Gray "Folder: " -NoNewline; Write-Host -ForegroundColor Green $folder
Write-Host -ForegroundColor Gray "WSUS Group: " -NoNewline; Write-Host -ForegroundColor Green $WSUSGroup
Write-Host -ForegroundColor Gray "Disk Storage Format: " -NoNewline; Write-Host -ForegroundColor Green $DiskFormat
Write-Host -ForegroundColor Gray "Active Directory Domain to join: " -NoNewline; Write-Host -ForegroundColor Green $ADDomain
Write-Host -ForegroundColor Gray "Organisational Unit (OU): " -NoNewline; Write-Host -ForegroundColor Green $OUpath
Write-Host -ForegroundColor Gray "Network Port Group: " -NoNewline; Write-Host -ForegroundColor Green $NetworkPG

if (!$IPAddress){
	Write-Host -ForegroundColor Gray "IP Address: " -NoNewline; Write-Host -ForegroundColor Green "DHCP"
} else {
	Write-Host -ForegroundColor Gray "IP Address: " -NoNewline; Write-Host -ForegroundColor Green $IPAddress
	Write-Host -ForegroundColor Gray "IP Subnet Mask: " -NoNewline; Write-Host -ForegroundColor Green $IPSubnetMask
	Write-Host -ForegroundColor Gray "IP Gateway: " -NoNewline; Write-Host -ForegroundColor Green $IPGateway
	Write-Host -ForegroundColor Gray "IP DNS Servers: " -NoNewline; Write-Host -ForegroundColor Green $IPDNSServers
}

$response = read-host "Enter y to continue, or any other key to abort"
if ($response -ne "y") {return}

Write-Host -BackgroundColor Green -ForegroundColor Black "Creating AD Computer object."
try {
	New-ADComputer -Name $name -SamAccountName $name -Description $Description -Path $OUpath -Server $ADDomain -Credential $domainCredentials
	$newsrv = Get-ADComputer -Identity $name -Server $ADDomain -Credential $domainCredentials
	Get-ADGroup -Identity WSUS.GPO.FILTER -Server $ADDomain -Credential $domainCredentials | Add-ADGroupMember -Members $newsrv
}
catch {
	Write-Host -BackgroundColor Red -ForegroundColor Black "Error creating AD Computer object."
	Write-Error $_
	return
}

Write-Host -BackgroundColor Green -ForegroundColor Black "Creating VM."
try {
	$nouvellevm = New-VM -Name $Name -Notes $Description -OSCustomizationSpec $OSSpec -DiskStorageFormat $DiskFormat -VMHost $vmhost -Template $template -Datastore $DatastoreCluster -Location $folder -Verbose -ErrorAction Stop -Confirm:$false
}
catch {
	Write-Host -BackgroundColor Red -ForegroundColor Black "Error provisioning VM."
	Write-Error $_
	return
}

if ($nouvellevm) {
	Write-Host -BackgroundColor Green -ForegroundColor Black "Setting Network Port Group."
	$nouvellevm | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName $NetworkPG -Verbose -Confirm:$false #| Out-Null
	Write-Host -BackgroundColor Green -ForegroundColor Black "Setting number of CPUs and RAM."
	$nouvellevm | Set-VM -NumCpu $NbCPU -MemoryGB $NbRAM -Verbose -Confirm:$false #| Out-Null
	Write-Host -BackgroundColor Green -ForegroundColor Black "Starting VM and opening console."
	$nouvellevm | Start-VM | Open-VMConsoleWindow
	Write-Host -BackgroundColor Green -ForegroundColor Black "Starting VM Customization."
	$scripttorun = $(Split-Path $script:MyInvocation.MyCommand.Path) + "\WaitVmCustomization.ps1"
	Sleep 30
	& $scripttorun -vmList $nouvellevm -timeoutSeconds 600
	Write-Host -BackgroundColor Green -ForegroundColor Black "Customization completed. Starting post-deployment customization."

	if ($template -match "core") {
		Write-Host "Windows Server Core - Setting Shell to PowerShell."
		RunVMScript "Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name Shell -Value 'PowerShell.exe -NoExit'"
	}

### Test IP connectivity

if (!$IPAddress){
	if ( (Get-VM -Name $nouvellevm).Guest.IPAddress -match "169.254." ) {
		Write-Host -BackgroundColor Red -ForegroundColor Black "Detected auto-assigned IP. Abording. Check the DHCP scope for $((Get-NetworkAdapter -VM $nouvellevm).NetworkName)"
		return
	}
} else {	
	$testGW = RunVMScript "Test-Connection (Get-NetRoute -DestinationPrefix 0.0.0.0/0 | Select-Object -ExpandProperty Nexthop) -Quiet -Count 1"
	if (!$testGW) {
		Write-Host -BackgroundColor Red -ForegroundColor Black "Could not ping gateway. Abording."
		return
	}
}

	Write-Host "Installing extra PowerShell Modules."
	RunVMScript "Install-PackageProvider -Name NuGet -Force -Verbose"
	RunVMScript "Install-Module -Name NuGet -Force -Verbose"
	RunVMScript "Install-Module -Name PSWindowsUpdate -Force -Verbose"
	
	Write-Host "Configuring Windows Updates."
	$registryPath = "'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'"
	$wsusurl = "'http://pvwbat0b.bnquebec.ca:8530'"
	RunVMScript "if (!(Test-Path $registryPath)) { New-Item -Path $registryPath -Force -Verbose }"
	RunVMScript "New-ItemProperty -Path $registryPath -Name 'WUServer' -Value $wsusurl -PropertyType STRING -Force -Verbose"
	RunVMScript "New-ItemProperty -Path $registryPath -Name 'WUStatusServer' -Value $wsusurl -PropertyType STRING -Force -Verbose"
	RunVMScript "Get-WUList"

	Write-Host -BackgroundColor Green -ForegroundColor Black "Adding VM $($Name) to WSUS group $($WSUSGroup)."
	$wsus | Get-WsusComputer -NameIncludes $name | Add-WsusComputer -TargetGroupName $WSUSGroup
	
	Write-Host "Running Windows Updates."
	RunVMScript "Get-WUInstall -Install -AcceptAll -IgnoreReboot -Verbose -Confirm:`$false"

	Write-Host "Updating PowerShell Help."
	RunVMScript "Update-Help"
	
	Write-Host "Enabling Remote Desktop."
	RunVMScript "Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name 'fDenyTSConnections' -Value 0 -Verbose"
	
	Write-Host "Disabling NetBios"
	RunVMScript "gwmi Win32_NetworkAdapterConfiguration|?{`$_.ServiceName -like 'vmxnet*'}|%{if(`$_.TcpipNetbiosOptions -ne 2){`$_.SetTcpipNetbios(2)}}"

	Write-Host "Disabling WINS Globaly"
	RunVMScript "(Get-WmiObject -list Win32_NetworkAdapterConfiguration).enablewins(`$false,`$false)"

	Write-Host "Enabling Firewall Rules for Remote Management Tools"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Remote Desktop' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Remote Event Log Management' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Remote Service Management' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'File and Printer Sharing' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Performance Logs and Alerts' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Remote Volume Management' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Windows Defender Firewall Remote Management' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Windows Management Instrumentation (WMI)' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'Windows Remote Management' -Verbose"
	RunVMScript "Enable-NetFirewallRule -DisplayGroup 'SNMP Service' -Verbose"

	Write-Host "Setting Time Zone"
	RunVMScript "Set-TimeZone -Name 'Eastern Standard Time' -Verbose"

	Write-Host "Changing drive letter for DVD to X:"
	RunVMScript "gwmi Win32_Volume -Filter `"DriveType = '5'`" | swmi -Arguments @{DriveLetter = 'X:'}"

	Write-Host "Disabling indexing on all drives."
	RunVMScript "gwmi Win32_Volume -Filter `"IndexingEnabled=`$true`" | swmi -Arguments @{IndexingEnabled=`$false}"
	
	Write-Host "Disabling Hibernation and enabling High Performance Profile."
	RunVMScript "c:\windows\system32\powercfg.exe /HIBERNATE off"
	RunVMScript "c:\windows\system32\powercfg.exe /SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"

	Write-Host "Running Windows Updates."
	RunVMScript "Get-WUInstall -Install -AcceptAll -AutoReboot -Verbose -Confirm:`$false"

	Write-Host "Rebooting VM. Ignore error."
	RunVMScript "Restart-Computer -Force"
	
	Write-Host -BackgroundColor Green -ForegroundColor Black "Creation completed. Server will be available after the reboot sequence."
	
} else {
	Write-Warning "$name was not created."
}

# c:\scripts\NlleVM-test.ps1 -Name testmartin2 -Description "test" -Site Berri -Template W2K19-DATACTR-core-TPL-1.2 -Folder test -NetworkPG DvPG-Infra-336 -IPAddress 10.9.36.74 -ADDomain bnquebec.ca -domainCredentials $creds -WSUSGroup Test
# $vm = "testmartin";Stop-VM -VM $vm -Confirm:$false -Verbose; Remove-VM -VM $vm -DeleteFromDisk -Confirm:$false -Verbose; Remove-ADComputer $vm -Confirm:$false -Verbose;$fqdn=$vm+".bnquebec.ca";$client = $wsus.GetComputerTargetByName($fqdn); $client.Delete()