#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.0.8
# Created on:   2/11/2013 5:53 PM
# Created by:   michael.england
# Last Modified By: Joe Stocker
# Last Modified Date: 1/27/2018
# Change History:
# Change 1 - On line 68 and 78 changed get and set commands to MailUser instead of Mailbox
# Change 2 - On line 62 and 63 changed the $user variable to strip off the last 3 hex characters imposed by Exchange 2010 SP1 Rollup 6
# Reference Blog Article 1 http://www.righthandedexchange.com/2013/02/fixing-imceaex-ndrs-missing-x500.html#!/2013/02/fixing-imceaex-ndrs-missing-x500.html
# Reference Blog Article 2 https://eightwone.com/2013/08/12/legacyexchangedn-attribute-myth/
#========================================================================


#X500 IMCEA Problems - or getting NDRs... without the NDRs

PARAM([switch]$AutoHeal,[int]$days=1,[string]$Servers)
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction 'SilentlyContinue' 

function Find-X500Failures([int]$days=1,[string]$servers,[switch]$AutoHeal){
	$start=("{0:MM/dd/yyyy}" -f (Get-Date).AddDays($Days * -1))
	$results=@()
	Write-Output ("Searching for messages sent after: {0}" -f $start)
	foreach ($server in (Get-TransportServer $servers*|Sort Name)) {
		Write-output ("Checking server - {0}" -f $server.Name)
		$results+=get-messagetrackinglog -Server $server -eventID FAIL -start $Start -ResultSize Unlimited |Where {$_.Recipients -match "IMCEAEX"}
	}
	
	$UniqueUsers=@()
	#Go through each result and recipient
	foreach ($result in $results){
		foreach ($Recipient in ($result.Recipients|where {$_ -match "IMCEAEX"})) {
			#remove Invalid Characters
			$Recipient=$Recipient.ToUpper()
			$Recipient=$Recipient.Remove(0,8)
			$Recipient=$Recipient.Replace("+20"," ")
			$Recipient=$Recipient.Replace("+28","(")
			$Recipient=$Recipient.Replace("+29",")")
			$Recipient=$Recipient.Replace("+2E",".")
			$Recipient=$Recipient.Replace("_O=","/O=")
			$Recipient=$Recipient.Replace("_OU=","/OU=")
			$Recipient=$Recipient.Replace("_CN=","/CN=")
			$Recipient=$Recipient.Replace("_cn=","/cn=")
			$Recipient=$Recipient.Remove($Recipient.IndexOf("@"))
			$UniqueUsers+=$Recipient
		}
	}
	#Remove DuplicateEntries
	$UniqueUsers=$UniqueUsers|sort -unique
	
	Write-Output ("Found {0} Unique user(s)" -f $UniqueUsers.Count)
	
	$Report=@()
	#Go through each entry and check for an existing user
	foreach ($usr in $UniqueUsers) {
		
		$user=$null
		#Pull User Name out of X500
		if ($usr.Contains("/CN")){
			$user=$usr.Remove(0,$usr.IndexOf("/CN="))
			$user=$user.Remove(0,$user.IndexOf("/CN=",3)+4 )
			Write-Host "looking for: " $user.Substring(0,$user.Length-3)
			$user =  $user.Substring(0,$user.Length-3)
		}
		
		if ($user) {
			#search for user
			$mbx=get-mailuser $user -ErrorAction 'SilentlyContinue' -ResultSize 1|where {$_.RecipientTypeDetails -ne "LegacyMailbox"}
			if ($mbx){

				$userTemp=$mbx|Select-Object Name, Alias, Status
				$userTemp.Status="Not Changed"
				if ($AutoHeal){
					$exists=$mbx.Emailaddresses|where {$_.addressstring -match $usr}
					if (!$exists){
						
$mbx.EmailAddresses+=("X500:{0}" -f $usr)
						Set-MailUser $mbx.identity -EmailAddresses $mbx.Emailaddresses
						$userTemp.Status="Updated"
					}
				}
				$Report+=$userTemp 
			}
		}
	}
	Write-Output "---------------------------------------`nResults`n---------------------------------------"
	if ($Report.count -gt 0){Write-Output $report|Sort Name}
	else{"Nothing to fix";$UniqueUsers }
}

Find-X500Failures -days $days -servers $servers -AutoHeal:$AutoHeal



