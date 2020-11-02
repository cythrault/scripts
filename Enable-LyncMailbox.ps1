function Get-FreeSkypeNumber {
    Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("Site Code")]
		[string]$sitecode
    )
    # Query Lync or failback to local CSV copy
    $sn = Get-CsUnassignedNumber | where {$_.identity -like "$sitecode*"} | select Identity, NumberRangeStart, NumberRangeEnd
    if (-Not $sn ) { $sn = Import-CSV -path C:\inetpub\ADUsers\App_Data\listUnassigned.csv }
    $numberrange=@()
    $assignedto=$null
    ($sn -match $sitecode)|%{
        [int]$areacode = $_.NumberRangeStart.TrimStart("tel:+1").SubString(0,3)
        [int]$start = $_.NumberRangeStart.TrimStart("tel:+1").SubString(3,7)
        [int]$end = $_.NumberRangeEnd.TrimStart("tel:+1").SubString(3,7)
        $numberrange += $start..$end
    }
    foreach ($suffix in $numberrange ) {
        [int64]$number="$areacode$suffix"
        $ldapfilter = "(|(proxyAddresses=eum:$number*)(msRTCSIP-Line=tel:+1$number))"
        $assignedto = (Get-ADObject -LDAPFilter $ldapfilter).name
        if ( -Not $assignedto ) { break } # bingo!
    }
    return $number
}

Start-Transcript -Path c:\ops\logs\Enable-LyncMailbox.log

$ErrorActionPreference = "SilentlyContinue"
[string]$dc = "dccan100dom21.corp.pbwan.net"

# define automated OU and get accounts. break if no accounts to process.
$last30days=((Get-Date).AddDays(-30)).Date
$searchBase = "OU=Accounts - Automated,OU=Canada,DC=corp,DC=pbwan,DC=net"
$targetUsers = Get-ADUser -SearchBase $searchBase -Properties * -Server $dc -Filter { whenCreated -ge $last30days }
Write-Host "New accounts found:" $targetUsers.Count
if ( $targetUsers.Count -eq 0 ) { Stop-Transcript; break }

# load sites/OUs details into hash table
$sites=@()
([xml](Get-Content C:\inetpub\ADUsers\App_Data\Lieu_travailCORP.xml)).Adresse.Element | %{$sites += $_}
$sitesOU=@{}
for ($i=0; $i -lt $sites.Count; $i++) { $sitesOU.Add( $sites[$i].scADSiteCode, $sites[$i].scOU ) }
# load sites/provinces details into hash table
$sitesProvince=@{}
for ($i=0; $i -lt $sites.Count; $i++) { $sitesProvince.Add( $sites[$i].scADSiteCode, $sites[$i].scProvince ) }
# define east/west provinces
$EastWest = @{
    'Newfoundland and Labrador' = 'CAeast'
    'Prince Edward Island' = 'CAeast'
    'Nova Scotia' = 'CAeast'
    'Québec' = 'CAeast'
    'Ontario' = 'CAeast'
    'Alberta' = 'CAwest'
    'Manitoba' = 'CAwest'
    'Saskatchewan' = 'CAwest'
    'British Columbia' = 'CAwest'
}
# site Lync pool/regions
$LyncPool = @{
    'CAeast' = 'enaarlyncpool.corp.pbwan.net'
    'CAwest' = 'endenlyncpool.corp.pbwan.net'
    'AU' = 'apbnelyncpool.corp.pbwan.net'
    'GB' = 'ealonlyncpool.corp.pbwan.net'
    'KR' = 'apbnelyncpool.corp.pbwan.net'
    'CN' = 'apbnelyncpool.corp.pbwan.net'
    'SE' = 'eastolyncpool.corp.pbwan.net'
    'TW' = 'apbnelyncpool.corp.pbwan.net'
}
# define region/Exchange DAGs
$exchdags = @{
	'AU' = 'APDAG02'
	'CA' = 'CADAG01'
	'CN' = 'ASDAG02'
	'GB' = 'EADAG01'
	'KR' = 'ASDAG02'
	'SE' = 'SEDAG01'
	'TW' = 'ASDAG02'
	'US' = 'ENDAG01'
	'ZA' = 'ZADAG01'
}

# Remove Stale PSS (testing phase only)
$stalePSS = Get-PSSession | ?{$_.Availability -eq "None"}
if ( $stalePSS ) { Write-Host "Removing stale PS sessions."; Remove-PSSession $stalePSS }

# Check if commands are available
#if (-Not (Get-Command Get-CSUser)) { Get-PSSession|where {$_.ComputerName -like "*lyncpool*" -or $_.ComputerName -like "*fepool*"} | Remove-PSSession }
#if (-Not (Get-Command Get-Mailbox)) { Get-PSSession|where ConfigurationName -eq "Microsoft.Exchange" | Remove-PSSession }

# load Exchange and Lync PSSessions
try {
    if (-Not $(Get-PSSession|where ConfigurationName -eq "Microsoft.Exchange")) {
        Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Exchange Session."
        $ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://dccan100exc01/PowerShell/ -Authentication Kerberos
        Import-PSSession $ExchSession | Out-Null
    }

# TODO add exchange online session

	if (-Not $(Get-PSSession|where {$_.ComputerName -like "*lyncpool*" -or $_.ComputerName -like "*fepool*"}) ) {
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Lync Session."
		#$LyncSession = New-PSSession -ConnectionUri https://enaarlyncpool.corp.pbwan.net/OcsPowerShell -Authentication NegotiateWithImplicitCredential 
		$LyncSession = New-PSSession -ConnectionUri https://glb-s4b-fepool-dccan100.corp.pbwan.net/OcsPowerShell -Authentication NegotiateWithImplicitCredential
		Import-PSSession $LyncSession | Out-Null
    }
}
catch {
    Write-Host -ForegroundColor Black -BackgroundColor Red "Error loading PS remote sessions - Can't proceed without valid Exchange and Lync sessions."
    break
}

# Check if commands are available
if (-Not (Get-Command Get-CSUser) -or -Not (Get-Command Get-Mailbox)) { 
    Write-Host -ForegroundColor Black -BackgroundColor Red "Commands are not availables - Can't proceed without valid Exchange and Lync sessions."
    break
}

foreach ($newuser in $targetusers) {
    #[bool]$ok = $True
    [bool]$mailboxenabled = $False
    [bool]$evenabled = $False
	[bool]$lyncenabled = $False
    [bool]$UMStatus = $False
    [bool]$triedEV = $False
    [bool]$triedUM = $False

    # Account validation
    if ( $newuser.extensionAttribute9 -eq $null -or $newuser.extensionAttribute9 -eq "CA-UNASSIGNED-POOL" ) { Write-Host $newuser.samAccountName "does not have a valid EA9"; continue }
    $siteCode = $newuser.extensionAttribute9
    Write-Host "Processing account $($newuser.samAccountName) in $siteCode from $($sitesProvince[$sitecode]), $($newuser.co)."

    # send email?
    # validate UPN - does not contains unwanted characters
    $id = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($newuser.UserPrincipalName))
    if ( $id -ne $($newuser.UserPrincipalName) ) { Write-Host "Account $($newuser.samAccountName) UPN $($newuser.UserPrincipalName) contains invalid characters. Corrected to: $id."; continue }
    # validate UPN suffix
    if ( $($newuser.UserPrincipalName) -notlike "*@wsp.com" ) { Write-Host "Account $($newuser.samAccountName) use and invalid UPN suffix: $($newuser.UserPrincipalName)"; continue }
    Write-Host "UPN for $($newuser.samAccountName) is $id."
    $employeeNumber = $newuser.samAccountName.Substring($newuser.samAccountName.Length-6)

    # Enable Mailbox
    # Check if already mailbox enabled. Detect if mailsomething else enabled and abort if so.
	if ( $newuser.msExchRecipientTypeDetails -eq 1 ) { 
		if ( $newuser.mail -eq $id ) { Write-Host "Account $($newuser.samAccountName) already has a mailbox with mail: $($newuser.mail)" }
		else { Write-Host "Account $($newuser.samAccountName) already has a mailbox but with wrong mail: $mail (should be $id)"; continue }	
	} elseif ( $newuser.msExchRecipientTypeDetails ) {
		Write-Host "Account $($newuser.samAccountName) already has a mail enabled account (msExchRecipientTypeDetails = $($newuser.msExchRecipientTypeDetails).)"; continue
	}
	else {
		# selecting a DB based on the user's region DAG or defaults to CA
		if ( $exchdags[$newuser.c] ) {
			$exchdb = $(Get-MailboxDatabase | where {($_.MasterServerOrAvailabilityGroup -eq $exchdags[$newuser.c]) -and ($_.IsExcludedFromProvisioning -eq $False) -and ($_.IsSuspendedFromProvisioning -eq $False -and ($_.ServerName -like "DCCAN100*")) }|Get-Random).Name
		} else {
			$exchdb = $(Get-MailboxDatabase | where {($_.MasterServerOrAvailabilityGroup -eq "CADAG01") -and ($_.IsExcludedFromProvisioning -eq $False) -and ($_.IsSuspendedFromProvisioning -eq $False) }|Get-Random).Name
		}
		Write-Host "Enabling mailbox on $($newuser.samAccountName) with mail $id on $exchdb"
		Enable-Mailbox -DomainController $dc -Identity $newuser.samAccountName -displayName $newuser.displayname -alias $newuser.samAccountName -primarySMTPAddress $id -Database $exchdb | Out-Null
		# TODO Enable mailbox culture when we know language
	}
	
    # Enable Lync
    # No other check then if already Lync enabled - should we check for mailbox?
    [bool]$lyncenabled = (Get-ADUser -Server $dc -Identity $newuser.samAccountName -Properties msRTCSIP-UserEnabled)."msRTCSIP-UserEnabled"
	if ( $lyncenabled -ne $True ) {
        # TODO: handle US East/West
        if ( $newuser.c -eq "CA" ) { $registrarPool=$LyncPool[$EastWest[$sitesProvince[$siteCode]]] }
        else { $registrarPool=$LyncPool[$newuser.c] }
        if ( $registrarPool ) {
            Write-Host "Enabling Lync on $($newuser.samAccountName) with sip: sip:$id on $registrarPool"
            Enable-CsUser -DomainController $dc –Identity $id –RegistrarPool $registrarPool –SipAddress "sip:$id" | Out-Null
        } else { Write-Host "Could not enable Lync on $($newuser.samAccountName) as because registrarPool is null." }
    } else {
        Write-Host "Account $($newuser.samAccountName) is already Lync enabled with $($newuser."msRTCSIP-PrimaryUserAddress")."
    }

    # Enable Enterprise Voice if Lync enabled
    [bool]$lyncenabled = (Get-ADUser -Server $dc -Identity $newuser.samAccountName -Properties msRTCSIP-UserEnabled)."msRTCSIP-UserEnabled"
	[bool]$evenabled = (Get-CSUser -DomainController $dc -Identity $id).EnterpriseVoiceEnabled
    if ( $lyncenabled -eq $True -and $evenabled -ne $True ) {
        [bool]$triedEV = $True
        # If phone number available...
        $sipnumber=Get-FreeSkypeNumber $newuser.extensionAttribute9
        if ( $sipnumber ) {
            Write-Host "Found free number: $sipnumber for $($newuser.extensionAttribute9)"
            # define lync number
            $lineUri = "tel:+1$sipnumber"
            # define telephoneNumber
            # TODO detect region and format appropriatly using E.164
            $phoneNum  = [String]::Format('{0:+1 (###) ###-####}',$sipnumber)
            Write-Host "Enabling Enterprise Voice on $id with lineUri: $lineUri"
            Set-CsUser -DomainController $dc -Identity $id -EnterpriseVoiceEnabled $True -LineUri $lineUri
            Write-Host "Setting phone number for $id with number: $phoneNum"
            Set-ADUser -Server $dc -Identity $newuser.samAccountName -OfficePhone $phoneNum
        } else {
            Write-Host "No phone number available for $($newuser.samAccountName) in site $sitecode. Skipping Enterprise Voice/UMMailbox enablement."
            $triedEV = $false
        }
    } elseif ( $evenabled -eq $True ) {
        Write-Host "Account $($newuser.samAccountName) in already Enterprise Voice enabled."
    } else {
        Write-Host "Can't enable Enterprise Voice on account $($newuser.samAccountName)."
    }

    # Enable UM Mailbox if Mailbox, Lync & EV enabled
    [bool]$mailboxenabled = (Get-Mailbox -DomainController $dc $id).IsMailboxEnabled
    [bool]$evenabled = (Get-CSUser -DomainController $dc -Identity $id).EnterpriseVoiceEnabled
	[bool]$lyncenabled = (Get-ADUser -Server $dc -Identity $newuser.samAccountName -Properties msRTCSIP-UserEnabled)."msRTCSIP-UserEnabled"
    [bool]$UMStatus = (Get-Mailbox -DomainController $dc -Identity $id).UMEnabled
    if ( $mailboxenabled -eq $True -and $evenabled -eq $True -and $UMStatus -ne $True ) {
        [bool]$triedUM = $True
        # TODO - use user language when available. using province for now.
        if ( $newuser.state -eq "Québec" ) { $UMpolicy = Get-UMMailboxPolicy -UMDialPlan CA-DialPlan-CA_Fr }
        else { $UMpolicy = Get-UMMailboxPolicy -UMDialPlan CA-DialPlan-CA }
        # note: $sip = $id ( = upn/mail )
        Write-Host "Enabling Unified Messaging on $($newuser.samAccountName) with lineUri: $lineUri on UM policy: $($UMpolicy.Name)"
        Enable-UMMailbox -DomainController $dc -Identity $id -UMMailboxPolicy $UMpolicy.Name | Out-Null
    } elseif ( $UMStatus -eq $True ) {
        Write-Host "Account $($newuser.samAccountName) in already UM enabled."
    } else {
        Write-Host "Account $($newuser.samAccountName) not suitable for UM."
    }	

    # Post Checks
    [bool]$mailboxenabled = (Get-Mailbox -DomainController $dc -Identity $id).IsMailboxEnabled
    [bool]$evenabled = (Get-CSUser -DomainController $dc -Identity $id).EnterpriseVoiceEnabled
	[bool]$lyncenabled = (Get-ADUser -Server $dc -Identity $newuser.samAccountName -Properties msRTCSIP-UserEnabled)."msRTCSIP-UserEnabled"
    [bool]$UMStatus = (Get-Mailbox -DomainController $dc -Identity $id).UMEnabled

    if ( $mailboxenabled -ne $True ) {
        Write-Host "Failed enabling mailbox on $($newuser.samAccountName)."
        $ok = $false
    } elseif ( $lyncenabled -ne $True ) {
        Write-Host "Failed enabling Lync on $($newuser.samAccountName)."
        $ok = $false
    } elseif ( $triedEV -eq $True -and $evenabled -ne $True ) { 
        Write-Host "Failed enabling EV on $($newuser.samAccountName)."
        $ok = $false
    } elseif ( $triedUM -eq $True -and $UMStatus -ne $True ) {
        Write-Host "Failed enabling UM on $($newuser.samAccountName)."
        $ok = $false
    } else {
        $ok = $True
		$adaccount = Get-ADUser -Identity $($newuser.samAccountName) -Properties * -Server $dc
		$ummailbox = Get-UMMailbox -Identity $($adaccount.UserPrincipalName) -DomainController $dc
    }

    # If Mailbox/Lync enabled, move account to target OU if defined
    # TODO: Define OU for non-CA users
    if ( $ok -ne $false ) {
        $targetOU = $sitesOU[$($newuser.extensionAttribute9)]
        if ( $targetOU ) { $targetOU = "$targetOU,DC=corp,DC=pbwan,DC=net" } else {
            Write-Host "Target OU is not defined for $($newuser.extensionAttribute9). Moving to temporary OU."
            $targetOU = "OU=Automated,OU=Users,OU=Canada,DC=corp,DC=pbwan,DC=net"
        }
        Write-Host "Moving account $($adaccount.samAccountName) to: $targetOU"
        # REMOVE -Whatif WHEN GOING INTO PROD
        Move-ADObject -Whatif -Server $dc -Identity $($adaccount.distinguishedName) -TargetPath $targetOU
        $body  = "This is an automatic notification to make you aware that an action is required following a new hired employee.<br>"
        $body += "Please make the necessary changes regarding the Group permissions.<br>"
        $body += "<br>"
        $body += "Employee Number: <b>$employeeNumber</b><br>"
		$body += "Given Name: <b>$($adaccount.givenName)</b><br>"
        $body += "Surname: <b>$($adaccount.sn)</b><br>"
        $body += "Display Name: <b>$($adaccount.displayName)</b><br>"
        $body += "Title: <b>$($adaccount.title)</b><br>"
        $body += "Department: <b>$($adaccount.department)</b><br>"
        if ( $adaccount.manager ) { $body += "Manager: <b>$((Get-ADUser $adaccount.manager -Properties displayName).displayName)</b><br>" }
        $body += "<br>"
        $body += "Account details:<br>"
        $body += "Legacy login name (SamAccountName:) <b>CORP\$($adaccount.samAccountName)</b><br>"
        $body += "Login name (UserPrincipalName:) <b>$($adaccount.UserPrincipalName)</b><br>"
        $body += "The temporary password assigned is: <b>WSPcube99</b><br>"
        if ( $mailboxenabled -eq $True ) { $body += "Primary mail address: <b>$($adaccount.mail)</b><br>" }
        if ( $lyncenabled -eq $True ) { $body += "Skype (SIP) address: <b>$($adaccount."msRTCSIP-PrimaryUserAddress")</b><br>" }
        if ( $evenabled -eq $True ) { $body += "Enterprise Voice Enabled: <b>$($adaccount."msRTCSIP-Line")</b> - <b>$($adaccount.telephoneNumber)</b><br>" } Else { $body += "Enterprise Voice is not Enabled.<br>" }
        if ( $UMStatus -eq $True ) { $body += "Mailbox enabled with Unified Messaging (UM) policy: <b>$($ummailbox.UMMailboxPolicy)</b><br>" }
        $body += "<br>"

        if ($adaccount.UserPrincipalName -match "[0-9]") {  $body += "[WARNING] The account name contains a number which might indicate a duplicate. Please validate if that's acceptable with the user.<br>" }

        if ( $targetOU -eq "OU=Automated,OU=Users,OU=Canada,DC=corp,DC=pbwan,DC=net" ) { $body += "[WARNING] The account has been moved to a temporary location. Site $($newuser.extensionAttribute9) does not have a default OU." }
		$body += "Account Canonical Name: <b>$($adaccount.cn)</b><br>"
		$body += "Account location: <b>$targetOU</b><br>"
        $body += "<br>"
        $body += "Automated Process<br>"
        $body += "Please do not reply to this email. This is an automatic notification and replies to this email address are unmonitored"
        Write-Host "Sending mail notification."
        # CHANGE RECIPIENT AND REMOVE COMMENT WHEN GOING INTO PROD
        #Send-MailMessage -From noreply@wsp.com -To martin.thomas@wsp.com -Encoding UTF8 -BodyAsHtml -Subject "New Arrival – User Created – $($newuser.displayName) - $employeeNumber – $($newuser.samAccountName)" -SmtpServer smtp-ca.1wsp.com -Body $body
		"$(Get-Date -UFormat "%Y-%m-%d %X") - $($adaccount.samAccountName) - $($newuser.extensionAttribute9) - Ok ($ok) - Mailbox ($mailboxenabled) - Lync ($lyncenabled) - EV ($evenabled) - UM ($UMStatus) - $targetOU" | Out-File -Append -FilePath c:\ops\logs\Enable-LyncMailbox-$(get-date -uformat %Y-%m).log
    } else {
        Write-Host "Not done processing $($newuser.samAccountName)."
        "$(Get-Date -UFormat "%Y-%m-%d %X") - $($newuser.samAccountName) - $($newuser.extensionAttribute9) - Ok ($ok) - Mailbox ($mailboxenabled) - Lync ($lyncenabled) - EV ($evenabled) - UM ($UMStatus)" | Out-File -Append -FilePath c:\ops\logs\Enable-LyncMailbox-$(get-date -uformat %Y-%m).log
		## todo email about failure
    }
	Write-Host "----------------------------------------------------"
}

#debug only!
#Disable-CsUser CAAT031335 -DomainController $dc -confirm:$false;Disable-Mailbox CAAT031335 -DomainController $dc -confirm:$false;Set-ADUser -server $dc CAAT031335 -Replace @{extensionAttribute9="CAMTR400"}
#Disable-CsUser CAMN088888 -DomainController $dc -confirm:$false;Disable-Mailbox CAMN088888 -DomainController $dc -confirm:$false;Set-ADUser -server $dc CAMN088888 -Replace @{extensionAttribute9="CAMTR100"}
#Disable-CsUser CAPD098886 -DomainController $dc -confirm:$false;Disable-Mailbox CAPD098886 -DomainController $dc -confirm:$false;Set-ADUser -server $dc CAPD098886 -Replace @{extensionAttribute9="CATHL100"}; Set-ADUser -server $dc CAPD098886 -Replace @{extensionAttribute4="098886"}
#Disable-CsUser CAPD098887 -DomainController $dc -confirm:$false;Disable-Mailbox CAPD098887 -DomainController $dc -confirm:$false;Set-ADUser -server $dc CAPD098887 -Replace @{extensionAttribute9="CAEDM600"}
#Disable-CsUser CAPD098888 -DomainController $dc -confirm:$false;Disable-Mailbox CAPD098888 -DomainController $dc -confirm:$false;Set-ADUser -server $dc CAPD098888 -Replace @{extensionAttribute9="CAQUE500"}
#$testacct="CAPD098885","CAPD098883","CAPD098882","CAMN088888","CAAT031335","CAPD098886","CAPD098887","CAPD098888","CAPD098884"
#$testacct|%{get-aduser $_|Move-ADObject -Server $dc -TargetPath "OU=Accounts - Automated,OU=Canada,DC=corp,DC=pbwan,DC=net" }

if ( $ExchSession ) { Remove-PSSession $ExchSession }
if ( $LyncSession ) { Remove-PSSession $LyncSession }

Stop-Transcript