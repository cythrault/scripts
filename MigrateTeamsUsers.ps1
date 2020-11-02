# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
# -File C:\tools\scripts\MigrateTeamsUser.ps1
$members = Get-ADGroupMember -Identity M365-License-Teams
$members | %{
	
	$user = Get-ADUser -Server "$((Get-ADForest).DomainNamingMaster):3268" -Identity $_.SamAccountName
	[string]$dc = switch -wildcard ( $user.DistinguishedName ) {
		"*DC=archives,DC=bnquebec,DC=ca" { (Get-ADDomain archives.bnquebec.ca).PDCEmulator }
		"*DC=biblio,DC=bnquebec,DC=ca" { (Get-ADDomain biblio.bnquebec.ca).PDCEmulator }
		default { (Get-ADDomain bnquebec.ca).PDCEmulator }
	}
	
	$enabled = (Get-ADUser $user.distinguishedName -Server $dc -Properties msRTCSIP-UserEnabled)."msRTCSIP-UserEnabled"
	if ($enabled -ne $true) {
		if ($user.UserPrincipalName -notmatch $newSuffix) {
				$newUpn = $user.UserPrincipalName.Split('@')[0]  + "@" + $newSuffix
				$user | Set-ADUser -Server $dc -UserPrincipalName $newUpn
		}
		Get-CsAdUser -Identity $user.DistinguishedName | Enable-CsUser -SipAddressType SamAccountName -SipDomain banq.qc.ca -RegistrarPool skypepoolfe01.banq.qc.ca -Verbose
		Invoke-Command -ComputerName pvwaad0a -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
		Start-Sleep -Seconds 120
	}

	$provider = (Get-CsUser -Identity $user.DistinguishedName).HostingProvider
	if ($provider -ne "sipfed.online.lync.com") {
		Do  {
			Get-CsUser -Identity $user.DistinguishedName | Move-CsUser -Target sipfed.online.lync.com -UseOAuth -MoveToTeams -Confirm:$false -Verbose -DomainController $dc
			Start-Sleep -Seconds 10
			$provider = (Get-CsUser -Identity $user.DistinguishedName).HostingProvider
			If ($provider -ne "sipfed.online.lync.com") {
				Invoke-Command -ComputerName pvwaad0a -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
				Start-Sleep -Seconds 120
			}
		} Until ($provider -eq "sipfed.online.lync.com")
	}
}
