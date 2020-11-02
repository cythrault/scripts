#C:\Users\mthomas\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
#Write-Host -Fore Black -Back Green "Loading functions."
#Get-ChildItem -Path C:\Users\martin.thomas\OneDrive\scripts\autoload\*.ps1 | Foreach-Object{ . $_.FullName } 
#Write-Host -Fore Black -Back Green "Done."

Write-Host -Foreground Black -Background Green "Loading credentials.."
#$creds = creds $env:USERDOMAIN\$env:USERNAME
#$ocreds = creds mthomas@banq.qc.ca

# Install-Module -Name CredentialManager
if (!(Get-InstalledModule CredentialManager)) {Write-Warning "Please install CredentialManager."; return}

# New-StoredCredential -Target admin -Persist LocalMachine -UserName bnquebec\mthomas -Password ...
# New-StoredCredential -Target o365 -Persist LocalMachine -UserName martin.thomas@banq.qc.ca -Password ...
$creds = Get-StoredCredential -target admin
$ocreds = Get-StoredCredential -target o365

if (!$creds) {Write-Warning "Empty credentials (creds)"; return}
if (!$ocreds) {Write-Warning "Empty credentials (ocreds)"; return}

Function Connect-Exchange {
	Write-Host -Foreground Black -Background Green "Connecting to Exchange."
	$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://PVWEXC0f.bnquebec.ca/PowerShell/ -Authentication Kerberos -Credential $creds
	Import-PSSession $ExchangeSession -DisableNameChecking
	Set-ADServerSettings -ViewEntireForest $true
}

#Function Connect-ExchangeOnline {
#	Write-Host -Foreground Black -Background Green "Connecting to Exchange Online."
#	$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $ocreds -Authentication Basic -AllowRedirection
#	Import-PSSession $ExchangeOnlineSession -DisableNameChecking
#}

Function Connect-Lync {
	Write-Host -Foreground Black -Background Green "Connecting to Lync."
	$LyncSession = New-PSSession -ConnectionUri https://skypepoolfe01.banq.qc.ca/OcsPowershell -Credential $creds
	Import-PSSession $LyncSession -DisableNameChecking
}

Function Connect-VI {
	[array]$vCenters = ("pvavcs0a.bnquebec.ca", "pvavcs0b.bnquebec.ca")
	if ( ((Get-PowerCLIConfiguration -Scope User).DefaultVIServerMode -ne "Multiple") -and ($vCenters.Count -gt 1) ) {
		Write-Warning "PowerCLI configuration set to single connection and multiple vCenters were specified: configuring for multiple vCenters."
		Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false
	}
	Connect-VIServer -Server $vCenters -Credential $creds
}

Function Connect-O365 {
	Write-Host -Foreground Black -Background Green "Connecting to AzureAD and MsolService."
	Connect-AzureAD #-Credential $ocreds
	Connect-MsolService #-Credential $ocreds
}