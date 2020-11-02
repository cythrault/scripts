function Set-RMAccess {
Param (
       [Parameter(Mandatory=$true, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
       $mailbox
)
Process {
		if ($mbx) {Remove-Variable mbx -ErrorAction SilentlyContinue}
		$mbx = Get-Mailbox $mailbox
		if (!$mailbox) {write-host "no mailbox"; return}
		$FAgrp = "$($mbx.Alias)-FullAccess"
		$SAgrp = "$($mbx.Alias)-SendAs"

		Write-Host -ForegroundColor Black -BackgroundColor Green "Executing on" $mbx

### Groups
       
		Write-Host -ForegroundColor Black -BackgroundColor Green "Creating groups."
		New-ADGroup -Name $FAgrp -GroupCategory Security -GroupScope Universal -Path "OU=Exchange,OU=Applications,OU=Groupes,OU=1.Gestion,DC=bnquebec,DC=ca"
		New-ADGroup -Name $SAgrp -GroupCategory Security -GroupScope Universal -Path "OU=Exchange,OU=Applications,OU=Groupes,OU=1.Gestion,DC=bnquebec,DC=ca"

		Write-Host -ForegroundColor Black -BackgroundColor Green "Adding $($SAgrp) to $($FAgrp)"
		Add-ADGroupMember -Identity $FAgrp -Members $SAgrp 

### Mailbox Permissions

		Write-Host -ForegroundColor Black -BackgroundColor Green "Adding Send As persmissions to $($SAgrp)"
		Add-ADPermission -Identity $mbx.Name -User $SAgrp -AccessRights ExtendedRight -ExtendedRights "Send As"
		Write-Host -ForegroundColor Black -BackgroundColor Green "Adding FullAccess persmissions for $($FAgrp)"
		Add-MailboxPermission -Identity $mbx.Name -User $FAgrp -AccessRights FullAccess -AutoMapping:$false
		Write-Host -ForegroundColor Black -BackgroundColor Green "Removing and Helpdesk permissions"
		Remove-MailboxPermission -Identity $mbx.Name -User HelpDesk -AccessRights FullAccess -Confirm:$false
		Write-Host -ForegroundColor Black -BackgroundColor Green "Adding and Helpdesk permissions with AutoMapping disabled"
		Add-MailboxPermission -Identity $mbx.Name -User HelpDesk -AccessRights FullAccess -AutoMapping:$false

		$mbxperms = Get-MailboxPermission -Identity $mbx |?{
              ($_.IsInherited -eq $false) -and
              ($_.User.RawIdentity -ne "NT AUTHORITY\SELF") -and
              ($_.User.RawIdentity -ne "BIBLIO\HelpDesk") -and
              ($_.User.RawIdentity -ne "BNQUEBEC\$FAgrp")
		}

		Write-Host -ForegroundColor Black -BackgroundColor Green "Current mailbox permissions"
		$mbxperms | ft -a

		Write-Host -ForegroundColor Black -BackgroundColor Green "Removing SIDs"
		$mbxperm | ?{($_.IsInherited -eq $false) -and ($_.User -like "S-1-5-*")} | Remove-MailboxPermission -Confirm:$false -Verbose

		Write-Host -ForegroundColor Black -BackgroundColor Green "Removing users from mailbox and adding to the group"
		$mbxperms |%{
			Remove-MailboxPermission -Identity $mbx -User $_.User -AccessRights $_.AccessRights -Confirm:$false -Verbose -ErrorAction SilentlyContinue
			$user = Get-ADUser $_.User.RawIdentity.Split("\")[1] -Server ldap.bnquebec.ca:3268
			Add-ADGroupMember -Identity $SAgrp -Members $user -ErrorAction SilentlyContinue 
		}

## msExchDelegateListlink

		Write-Host -ForegroundColor Black -BackgroundColor Green "Adding users to msExchDelegateListlink"
		Set-ADUser -Identity $mbx.distinguishedName -Clear msExchDelegateListlink
		Get-ADGroupMember -Identity $FAgrp -Recursive |select distinguishedName | %{
			Set-ADUser -Identity $mbx.distinguishedName -Add @{msExchDelegateListlink=$_.distinguishedName}
		}
		Write-Host -ForegroundColor Black -BackgroundColor Green "msExchDelegateListlink on $mbx.distinguishedName"
		Get-ADUser $mbx.distinguishedName -Properties msExchDelegateListlink | Select -ExpandProperty msExchDelegateListlink
		
		$mbxperms = Get-MailboxPermission -Identity $mbx |?{
              ($_.IsInherited -eq $false) -and
              ($_.User.RawIdentity -ne "NT AUTHORITY\SELF") -and
              ($_.User.RawIdentity -ne "BIBLIO\HelpDesk") -and
              ($_.User.RawIdentity -ne "BNQUEBEC\$FAgrp")
		}
		
		Write-Host -ForegroundColor Black -BackgroundColor Green "New mailbox permissions"
		$mbxperms | ft -a

		Write-Host -ForegroundColor Black -BackgroundColor Green "Members of $($FAgrp)"
		(Get-ADGroupMember -Identity $FAgrp).distinguishedName

### AD Permissions

		$adperms = Get-Mailbox -identity $mbx | Get-ADPermission |?{
			($_.IsInherited -eq $false) -and
			($_.User -notlike "S-1-5-32*") -and
			($_.User.RawIdentity -notlike "NT AUTHORITY\*") -and
			($_.User.RawIdentity -ne "Everyone") -and
			($_.User.RawIdentity -ne "BNQUEBEC\Domain Admins") -and
			($_.User.RawIdentity -ne "BNQUEBEC\Cert Publishers") -and
			($_.User.RawIdentity -ne "BNQUEBEC\RAS and IAS Servers") -and
			($_.User.RawIdentity -ne "BIBLIO\HelpDesk") -and
			($_.User.RawIdentity -ne "BNQUEBEC\$SAgrp")
		}

		Write-Host -ForegroundColor Black -BackgroundColor Green "Current AD permissions"
		$adperms | ft -a

		Write-Host -ForegroundColor Black -BackgroundColor Green "Removing users from AD account and adding to the group"
		$adperms | Remove-ADPermission -Confirm:$false -Verbose
		$adperms | %{
			$user = Get-ADUser $_.User.RawIdentity.Split("\")[1] -Server ldap.bnquebec.ca:3268
			Add-ADGroupMember -Identity $SAgrp -Members $user -ErrorAction SilentlyContinue 
		}
		
		$adperms = Get-Mailbox -identity $mbx | Get-ADPermission |?{
			($_.IsInherited -eq $false) -and
			($_.User -notlike "S-1-5-32*") -and
			($_.User.RawIdentity -notlike "NT AUTHORITY\*") -and
			($_.User.RawIdentity -ne "Everyone") -and
			($_.User.RawIdentity -ne "BNQUEBEC\Domain Admins") -and
			($_.User.RawIdentity -ne "BNQUEBEC\Cert Publishers") -and
			($_.User.RawIdentity -ne "BNQUEBEC\RAS and IAS Servers") -and
			($_.User.RawIdentity -ne "BIBLIO\HelpDesk") -and
			($_.User.RawIdentity -ne "BNQUEBEC\$SAgrp")
		}
		
		Write-Host -ForegroundColor Black -BackgroundColor Green "New AD permissions"
		$adperms | ft -a

		Write-Host -ForegroundColor Black -BackgroundColor Green "Members of $($SAgrp)"
		(Get-ADGroupMember -Identity $SAgrp).distinguishedName
	}
}