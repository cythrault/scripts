#Requires -Version 5

<#
.SYNOPSIS
	Setup Basic configuration of the server.

.DESCRIPTION
	Set-ServerConfig configure the network, HDD and Pagefile settings of the server.

.PARAMETER IP
	Specify the Network IP to use for the new server.

.PARAMETER Name
	Specify the machine name to use when joining to the domain.

.PARAMETER Credentials
	Specify the credentials to use when joining this machine to the domain.

.NOTES
	Version:	1.0
	Author:		Yann Bourgault

	Version:	2.0
	Author:		Martin Thomas
	Creation Date:	April 26th 2019
	Purpose/Change:	Update for Windows 2019

	Version:	2.1
	Author:		Martin Thomas
	Creation Date:	July 31st 2019
	Purpose/Change:	Add various base configuration.

.EXAMPLE
	Set-ServerConfig -Name PVWAPP0a -IP 10.9.36.249
#>


[CmdletBinding()]
param(
	[Parameter(Mandatory=$True)]
	[ValidateScript({$_ -match [IPAddress]$_})]
	[string] $IP,
	[Parameter(Mandatory=$True)]
	[ValidatePattern("^(?![0-9]{1,15}$)[a-zA-Z0-9-]{1,15}$")]
	[string] $Name,
	[Parameter(Mandatory=$True)]
	[ValidateNotNull()]
	[System.Management.Automation.PSCredential]
	[System.Management.Automation.Credential()]
	$Credentials
)

$DNSServers = "10.9.35.11","10.9.35.62"
$DNSSuffix = "bnquebec.ca"
$DNSSearchSuffix = "bnquebec.ca","biblio.bnquebec.ca","archives.bnquebec.ca"
$IPByte = $IP.SPlit(".")
$Gateway = $IPByte[0] + "." + $IPByte[1] + "." + $IPByte[2] + ".1"

Write-Host "Configuring Network"
New-NetIPAddress -InterfaceAlias Ethernet0 -IPAddress $IP -DefaultGateway $Gateway -AddressFamily IPv4 -PrefixLength 24
Set-DnsClientServerAddress -InterfaceAlias Ethernet0 -ServerAddresses $DNSServers
Set-DnsClientGlobalSetting -SuffixSearchList $DNSSearchSuffix
Set-DnsClient -InterfaceAlias Ethernet0 -ConnectionSpecificSuffix $DNSSuffix

Write-Host "Disabling NetBios"
$nics = gwmi Win32_NetworkAdapterConfiguration | Where-Object {$_.ServiceName -like 'vmxnet*'}
get If ($nic.TcpipNetbiosOptions -ne 2) { $nic.SetTcpipNetbios(2) } }

Write-Host "Disabling WINS Globaly"
$nicClass = Get-WmiObject -list Win32_NetworkAdapterConfiguration
$nicClass.enablewins($false,$false)

Write-Host "Enabling Firewall Rules for Remote Management Tools"
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
Enable-NetFirewallRule -DisplayGroup "Remote Event Log Management"
Enable-NetFirewallRule -DisplayGroup "Remote Service Management"
Enable-NetFirewallRule -DisplayGroup "File and Printer Sharing"
Enable-NetFirewallRule -DisplayGroup "Performance Logs and Alerts"
Enable-NetFirewallRule -DisplayGroup "Remote Volume Management"
Enable-NetFirewallRule -DisplayGroup "Windows Defender Firewall Remote Management"
Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)"
Enable-NetFirewallRule -DisplayGroup "Windows Remote Management"
Enable-NetFirewallRule -DisplayGroup "SNMP Service"

Write-Host "Changing drive letter for DVD to X:"
gwmi Win32_Volume -Filter "DriveType = '5'" | swmi -Arguments @{DriveLetter = 'X:'}

Write-Host "Disabling indexing on all drives."
gwmi Win32_Volume -Filter "IndexingEnabled=$true" | swmi -Arguments @{IndexingEnabled=$false}

Write-Host "Enabling Remote Desktop."
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0

Write-Host "Disabling Hibernation and enabling High Performance Profile."
c:\windows\system32\powercfg.exe /HIBERNATE off
c:\windows\system32\powercfg.exe /SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

Write-Host "Setting TimeZone"
Set-TimeZone -Id "Eastern Standard Time" -Verbose

Write-Host "Updating PowerShell Help."
Update-Help

Write-Host "Installing extra PowerShell Modules."
Install-PackageProvider -Name NuGet -Force
Install-Module -Name NuGet -Force
Install-Module -Name PSWindowsUpdate -Force

Write-Host "Installing Windows Updates."
if (!(Test-Path $registryPath)) { New-Item -Path $registryPath -Force | Out-Null }
New-ItemProperty -Path $registryPath -Name "WUServer" -Value "http://pvwbat0b.bnquebec.ca:8530" -PropertyType DWORD -Force | Out-Null
New-ItemProperty -Path $registryPath -Name "WUStatusServer" -Value "http://pvwbat0b.bnquebec.ca:8530" -PropertyType DWORD -Force | Out-Null
Get-WUInstall -MicrosoftUpdate -Install -AcceptAll -IgnoreReboot -Confirm:$false -Verbose

Write-Host "Renaming to $Name."
Rename-Computer -NewName $Name -Force -Verbose

Write-Host "Joining $Name to $DNSSuffix."
Add-Computer -DomainName $DNSSuffix -Credential $Credentials -Verbose -Restart