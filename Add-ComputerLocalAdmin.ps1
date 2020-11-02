$adminCORP = "corp\camt055545it9"
$adminGCG = "gcg\mt2"
$adminMMM = "mmm\mt1"
$adminFOCUSED = "focused\admin-mt"
$adminDMZ = "dmz\mt1"

$creds = Get-Credential -Username "admin" -Message "type password"

Write-Host "Getting server list..."
$servers = Get-ADComputer -Server dmz.local -Filter {OperatingSystem -Like "Windows *Server*"} | where {$_.DistinguishedName -NotLike "*OU=Domain Controllers*"}
$servers += Get-ADComputer -Server focuscorp.ca -Filter {OperatingSystem -Like "Windows *Server*"} | where {$_.DistinguishedName -NotLike "*OU=Domain Controllers*"}
$servers += Get-ADComputer -Server gcg.local -Filter {OperatingSystem -Like "Windows *Server*"} | where {$_.DistinguishedName -NotLike "*OU=Domain Controllers*"}
$servers += Get-ADComputer -Server mmm.ca -Filter {OperatingSystem -Like "Windows *Server*"} | where {$_.DistinguishedName -NotLike "*OU=Domain Controllers*"}
$servers += Get-ADComputer -Server corp.pbwan.net -Filter {OperatingSystem -Like "Windows *Server*"} -SearchBase "OU=Canada,DC=corp,DC=pbwan,DC=net"

$GroupName = "CORP/GRP-RBC-F-GLB-ServerAdmins-ALL"
$Group = [ADSI]"WinNT://$GroupName,group"
$start = Get-Date
[nullable[double]]$secondsRemaining = $null
$servers|%{
	$Computer = $_
	#Write-Host $Computer.DistinguishedName
	switch -wildcard ($Computer.DistinguishedName.ToLower()) {
		'*dc=corp*' { $adminUsername = $adminCORP }
		'*dc=gcg*' { $adminUsername = $adminGCG }
		'*dc=focus*' { $adminUsername = $adminFOCUSED }
		'*dc=mmm*' { $adminUsername = $adminMMM }
		'*dc=dmz*' { $adminUsername = $adminDMZ }
		default { $adminUsername = $null }
		}
	Try {
		[ADSI]$AdminGroup = "WinNT://$($Computer.DNSHostName)/Administrators,group"
		$AdminGroup.PsBase.Username = $adminUsername
		$AdminGroup.PsBase.Password = $creds.GetNetworkCredential().Password
		$AdminGroup.Add($Group.Path) | Out-Null
		$Status = "Added $GroupName to local Administrators group"
		}
	Catch {
		$Status = $_.Exception.Message.Replace("`n","").Replace("`r","")
		}
	[pscustomobject]@{Name=$Computer.Name;DNSHostName=$Computer.DNSHostName;Enabled=$Computer.Enabled;DistinguishedName=$Computer.DistinguishedName;Status=$Status}
	$i++
	$pct = ($i / $servers.Count) * 100
	$secondsElapsed = (Get-Date) - $start
	$secondsRemaining = ($secondsElapsed.TotalSeconds / $i) * ($servers.Count - $i)
	Write-Progress -Activity "Applying permissions..." -CurrentOperation $Computer.DNSHostName -Status "Completed: $i of $($servers.Count) - $([math]::Round($pct))%" -PercentComplete $pct -SecondsRemaining $secondsRemaining
}
Write-Progress -Activity "Applying permissions..." -Completed -Status "All done."
Write-Host "Completed $($servers.Count) servers in $($secondsElapsed.TotalMinutes) minutes."