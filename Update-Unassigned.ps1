# -file "c:\ops\Update-Unassigned.ps1"

if (-Not $(Get-PSSession|where ComputerName -like "*lyncpool*")) {
    Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Lync Session."
    $LyncSession = New-PSSession -ConnectionUri https://enaarlyncpool.corp.pbwan.net/OcsPowerShell -Authentication NegotiateWithImplicitCredential 
    Import-PSSession $LyncSession | Out-Null
}
    
Get-CsUnassignedNumber |where {$_.identity -like "CA*"} |select Identity, NumberRangeStart, NumberRangeEnd|Export-Csv -Path "c:\ops\listUnassigned.csv" -Encoding UTF8 -NoTypeInformation

if(Compare-Object -ReferenceObject $(Get-Content -Path "c:\ops\listUnassigned.csv") -DifferenceObject $(Get-Content -Path "C:\inetpub\ADUsers\App_Data\listUnassigned.csv")) {
    Copy-Item -Path "C:\inetpub\ADUsers\App_Data\listUnassigned.csv" -Destination "C:\inetpub\ADUsers\App_Data\listUnassigned.old\listUnassigned - $(get-date -uformat %Y-%m-%d).csv"
    Copy-Item -Path "c:\ops\listUnassigned.csv" -Destination "C:\inetpub\ADUsers\App_Data\listUnassigned.csv"
	"$(Get-Date -UFormat "%Y-%m-%d %X") - Unassigned Number List was updated." | Out-File -Append -FilePath c:\ops\logs\Update-Unassigned-$(get-date -uformat %Y-%m).log
}

if ( $LyncSession ) { Remove-PSSession $LyncSession }