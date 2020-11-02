#Requires -Version 5

# https://docs.microsoft.com/en-ca/previous-versions/windows/desktop/legacy/hh846315(v=vs.85)
#if (Get-ItemProperty -Path 'HKLM:Software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -Name ServerCore -ErrorAction Ignore) {
if ((Get-ItemProperty -Path 'HKLM:Software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -Name ServerCore -ErrorAction Ignore) -and ((Get-ItemPropertyValue -Path 'HKLM:Software\Microsoft\Windows NT\CurrentVersion\Server\ServerLevels' -Name Server-Gui-Shell -ErrorAction Ignore) -ne 1)) {
	Write-Host "Setting Shell to CMD."
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name Shell -Value 'cmd.exe /c "cd /d "%USERPROFILE%" & start cmd.exe /k runonce.exe /AlternateShellStartup"'
}

Write-Host "Disabling NetBios"
$nics = gwmi Win32_NetworkAdapterConfiguration | Where-Object {$_.ServiceName -like 'vmxnet*'}
foreach ($nic in $nics) { If ($nic.TcpipNetbiosOptions -ne 2) { $nic.SetTcpipNetbios(2) } }

Write-Host "Disabling WINS Globaly"
$nicClass = Get-WmiObject -list Win32_NetworkAdapterConfiguration
$nicClass.enablewins($false,$false)

Write-Host "Enabling SNMP Service"
Install-WindowsFeature SNMP-Service

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

Write-Host "Setting Time Zone"
Set-TimeZone -Name "Eastern Standard Time"

Write-Host "Changing drive letter for DVD to X:"
gwmi Win32_Volume -Filter "DriveType = '5'" | swmi -Arguments @{DriveLetter = 'X:'}

Write-Host "Disabling indexing on all drives."
gwmi Win32_Volume -Filter "IndexingEnabled=$true" | swmi -Arguments @{IndexingEnabled=$false}

Write-Host "Enabling Remote Desktop."
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0

Write-Host "Disabling Hibernation and enabling High Performance Profile."
c:\windows\system32\powercfg.exe /HIBERNATE off
c:\windows\system32\powercfg.exe /SETACTIVE 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

Write-Host "Updating PowerShell Help."
Update-Help

Write-Host "Installing extra PowerShell Modules."
Install-PackageProvider -Name NuGet -Force
Install-Module -Name NuGet -Force
Install-Module -Name PSWindowsUpdate -Force

Write-Host "Installing Windows Updates."
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$wsusurl = "http://pvwbat0b.bnquebec.ca:8530"
if (!(Test-Path $registryPath)) { New-Item -Path $registryPath -Force }
New-ItemProperty -Path $registryPath -Name "WUServer" -Value $wsusurl -PropertyType STRING -Force
New-ItemProperty -Path $registryPath -Name "WUStatusServer" -Value $wsusurl -PropertyType STRING -Force
Get-WUInstall -Install -AcceptAll -IgnoreReboot -Verbose -Confirm:$false

Write-Host "Clearing Event Logs."
wevtutil el | Foreach-Object {wevtutil cl "$_"}

Write-Host "Resetting Windows Update database."
dism /online /Cleanup-Image /StartComponentCleanup /ResetBase