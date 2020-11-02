if (-Not $cred ) { $cred = Get-Credential -Credential admin.martin.thomas@opus.co }
if (-Not $opuscred ) { $opuscred = Get-Credential -Credential "opus\mat181_a" }
if (-Not $corpcred ) { $corpcred = Get-Credential -Credential "corp\camt055545it9" }
if (-Not $it5creds ) { $it5creds = Get-Credential -Credential "corp\camt055545it5" }

$stalePSS = Get-PSSession | ?{$_.Availability -eq "None"}
if ( $stalePSS ) { Remove-PSSession $stalePSS }

if (-Not $Session ) {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    #Import-Module AzureAD
    #Connect-AzureAD -Credential $cred
    Import-Module MSonline
    Connect-MsolService -Credential $cred
}

if (-Not $(Get-PSSession|where {$_.ComputerName -like "*lyncpool*" -or $_.ComputerName -like "*fepool*"}) ) {
	$LyncSession = New-PSSession -ConnectionUri https://enaarlyncpool.corp.pbwan.net/OcsPowerShell -Credential $it5creds 
    #$LyncSession = New-PSSession -ConnectionUri https://glb-s4b-fepool-dccan100.corp.pbwan.net/OcsPowerShell -Credential $it5creds
    Import-PSSession $LyncSession | Out-Null
}

function prepgroups {

$opusmig=@()
#$source = "\\cacol2nas10\data$\users\Quebec\martin.thomas\opus\Exchange Batches 1-7.xlsx"
$source = "j:\opus\Batch 11.xlsx"
$groupName = "P365-CA-20180917-03"
Import-Excel $source  -WorksheetName "Batch 11"|%{
    if ($_."E-Mail Address") {
        $opusmig += $_."E-Mail Address"
    }
}

$opusmig|%{
	$opusacct = Get-ADUser -Server opus.global -Filter {mail -eq $_} -Properties * | Select Name, DistinguishedName, givenName, sn, displayName, SamAccountName, UserPrincipalName, mail, extensionAttribute12, msExchExtensionAttribute20, msExchExtensionAttribute21, msExchExtensionAttribute23 -ErrorAction SilentlyContinue
	$ea12 = $opusacct.extensionAttribute12
    if (-Not  $opusacct ) { write-host -background darkred $_ "not found in Opus."}
    if ( $ea12 ) {
	    $corpacct = Get-ADUser -Filter {extensionAttribute12 -eq $ea12} -ErrorAction SilentlyContinue
	    if (-Not $corpacct ) { write-host -background darkred $_ "($ea12) not found in CORP." }
    	return $opusacct
    } else { write-host -background darkred $_ "EA12 not defined in Opus."}
} | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "\\corp\ca\data$\users\Quebec\martin.thomas\opus\$groupName.csv"
}

function prepwaves {
Param ( [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)] [string]$groupName )
# > 48h before - run as opus EO admin
Write-Host "Adding group and adding membership"
$csv = Import-Csv \\corp\ca\data$\users\Quebec\martin.thomas\opus\$groupName.csv
if (-Not $(Get-ADGroup -Server opus.global -Identity $groupName)) { New-ADGroup -Name $groupName -GroupCategory Security -GroupScope Global -Server opus.global -Verbose -Path "OU=Groups,OU=CA,DC=opus,DC=global" -Credential $opuscred }
$csv | %{
    $sam = $_.samAccountName
    if ( (Get-MsolUser -UserPrincipalName $_.UserPrincipalName).isLicensed -eq $true ) {
        Add-ADGroupMember -Server opus.global -Verbose -Identity $groupName -Members $sam -ErrorAction SilentlyContinue -Credential $opuscred
    } else {
        Write-Host $_.UserPrincipalName "is not Office 365 Licensed."
    }
}
}

function post {
Param ( [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)] [string]$group )
$members = @()
$members = Get-ADGroupMember -Server opus.global -Identity $group
$members | % {$i = 0} { $i++
    $total =  $members.Count; $pct = $([math]::Round( ( ( $i / $total ) * 100 ), 1))
    Write-Progress -Activity "Opus Exchange Migration Pre-Flight Script" -Status "Processing $sam" -PercentComplete $pct -CurrentOperation "$i of $total, $pct%"
    $sam = $_.SamAccountName
    $corpacct = Get-ADUser -Filter {extensionAttribute12 -eq $sam} -Properties proxyAddresses, targetAddress, mail -ErrorAction SilentlyContinue
    $OPUSacct = Get-ADUser -Server opus.global -Identity $sam -ErrorAction SilentlyContinue
    #$pwdSetOpus = (Get-ADUser -Server opus.global -Identity $sam -Properties pwdLastSet | select @{name="pwdLastSet"; expression={[datetime]::fromFileTime($_.pwdlastset)}}).pwdlastset
    #$pwdSetCORP = (Get-ADUser -Filter {extensionAttribute12 -eq $sam} -Properties pwdLastSet | select @{name="pwdLastSet"; expression={[datetime]::fromFileTime($_.pwdlastset)}}).pwdlastset
    #write-host "CORP Last Pwd Set for $sam - Opus $pwdSetOpus, CORP: $pwdSetCORP"

    if ( $corpacct ) {
# pre
        Add-ADGroupMember -Verbose -Server opus.global -Identity "NZ - Outlook 3 month cache" -Members $sam -ErrorAction SilentlyContinue -Credential $opuscred
        Add-ADGroupMember -Verbose -Server corp.pbwan.net -Identity GRP-GLB-IT-EVCloudJournalO365 -Members $corpacct.SamAccountName -Credential $corpcred
# post
        if ( $( (Get-Mailbox $OPUSacct.UserPrincipalName ).ForwardingSmtpAddress ) -eq "smtp:$($corpacct.mail)" ) {
            Write-Host "[Migrated] $((Get-Mailbox $OPUSacct.UserPrincipalName ).ForwardingSmtpAddress) - $($corpacct.mail)"

            $sip = "sip:$($corpacct.mail)"; Set-CsUser -Verbose -Identity $corpacct.UserPrincipalName -SipAddress $sip
            Set-CASMailbox -Identity $OPUSacct.UserPrincipalName -ActiveSyncEnabled $false -ImapEnabled $false -PopEnabled $false
            Add-DistributionGroupMember -Verbose -Identity "BT - NDR Transport Rule" -Member $OPUSacct.UserPrincipalName -ErrorAction SilentlyContinue
            Set-Mailbox -Verbose -Identity $OPUSacct.UserPrincipalName -DeliverToMailboxAndForward:$False

            $ta = $corpacct.proxyAddresses.Where( {$_ -like "SMTP:*@wsponline.mail.onmicrosoft.com"})
            If ($ta.Count -eq 1) {$newTarget = $ta[0]} Else {$newTarget = $null}
            If ($newTarget) {
                If ($corpacct.targetAddress -ne $newTarget) {
                    Write-Host "Changing targetAddress for $sam from: $($corpacct.targetAddress) to: $newTarget"
                    Set-ADUser -Identity $corpacct.samAccountName -Replace @{targetAddress=($newTarget)} -Verbose -Server corp.pbwan.net -Credential $corpcred
                } Else {
                    Write-Host "TargetAddress for $sam is already OK: $($corpacct.targetAddress)" -ForegroundColor Green
                }
            } else { Write-Host "Failed to find new target address for account $sam" -ForegroundColor Red }
        } else { Write-Host -BackgroundColor DarkRed "[NOT MIGRATED] $((Get-Mailbox $OPUSacct.UserPrincipalName ).ForwardingSmtpAddress) - $($corpacct.mail)" }
    }
}
}
