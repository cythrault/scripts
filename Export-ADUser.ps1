Param( [Parameter(Mandatory=$true)][string]$sam )
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://dccan100exc01/PowerShell/
Import-PSSession $Session
get-aduser $sam -properties mailNickname,lastlogontimestamp | Select-Object Name,mailNickname,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}
$outBaseName="\\CaCol1Pst01.corp.pbwan.net\CORPDisabledUsers$\$($sam)-$(get-date -uformat %Y%m%d)"
csvde.exe -s dccan100dom21.corp.pbwan.net -f "$($outBaseName)-allattr.csv" -r "(sAMAccountName=$sam)"
$mailNickname=$(get-aduser -Identity $sam -Properties mailNickname).mailNickname
if ($mailNickname) {
	$mr=New-MailboxExportRequest -Mailbox $mailNickname -Name $mailNickname -BadItemLimit 999 -AcceptLargeDataLoss -FilePath "$($outBaseName).pst" -Confirm:$false -Verbose
	while ( $mr.Status -ne "Completed" ) {
		Sleep 10
		$mr=$(Get-MailboxExportRequest $mailNickname)
		$mrstats=$(Get-MailboxExportRequestStatistics $mr.name)
		Write-Progress -Activity "Exporting $mailNickname" -Status "$($mr.Status) / $($mrstats.StatusDetail.Value)" -percentComplete $mrstats.PercentComplete
		}
	}