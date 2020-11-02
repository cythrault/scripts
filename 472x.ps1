$dcs = $([System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain()).DomainControllers.Name
if ( [System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain().Children ) {
    $dcs += [System.Directoryservices.Activedirectory.Domain]::GetCurrentDomain().Children.DomainControllers.Name
    }
$total = $dcs.Count
$filter=@{
    Logname='Security'
    ID=@(1102,4625,4720,4722,4726,4727,4728,4729,4730,4732,4733,4738,4740,4746,4747,4745,4751,4752,4756,4757,4761,4762,4767)
    #StartTime=[datetime]::Today.AddDays(-7)
}
$i = 0; $startTime = Get-Date
$dcs | Sort-Object | ForEach-Object{
	$i++
	$elapsedTime = New-TimeSpan -Start $startTime -End (Get-Date)
    $timeRemaining = New-TimeSpan -Seconds (($elapsedTime.TotalSeconds / ($i / $total)) - $elapsedTime.TotalSeconds)
    Write-Progress -Activity "Exporting" -PercentComplete $(($i / $total) * 100)  -Status "$_ - $i of $total - Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -SecondsRemaining $timeRemaining.TotalSeconds
    Write-Host "Scanning $_"
    Get-WinEvent -ComputerName $_ -FilterHashtable $filter |
        Select-Object TimeCreated, Id, @{
            n='AccountName';
            e={ ($_.message -replace '\n', ' ') -replace '.*?account name:\t+([^\s]+).*', '$1' }
        }, @{
            n='TargetAccount';
            e={ ($_.message -replace '\n', ' ') -replace '.*account name:\t+([^\s]+).*', '$1' }
        }
} | Export-Excel 472x.xlsx -AutoSize -FreezeTopRow
#Send-MailMessage -Attachments "E:\Security Logs\472x.csv" -To david.gilmour@wsp.com -From ADUSers-Admin@wsp.com -Subject 472x -SmtpServer DCCAN100EXC02.corp.pbwan.net