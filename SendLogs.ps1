# Program/script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
# Add arguments: -NoProfile -ExecutionPolicy Unrestricted -File sendlogs.ps1
# Starts in: c:\ops

function Get-LPInputFormat{
	param([String]$InputType)
	switch($InputType.ToLower()){
		"ads"{$inputobj = New-Object -comObject MSUtil.LogQuery.ADSInputFormat}
		"bin"{$inputobj = New-Object -comObject MSUtil.LogQuery.IISBINInputFormat}
		"csv"{$inputobj = New-Object -comObject MSUtil.LogQuery.CSVInputFormat}
		"etw"{$inputobj = New-Object -comObject MSUtil.LogQuery.ETWInputFormat}
		"evt"{$inputobj = New-Object -comObject MSUtil.LogQuery.EventLogInputFormat}
		"fs"{$inputobj = New-Object -comObject MSUtil.LogQuery.FileSystemInputFormat}
		"httperr"{$inputobj = New-Object -comObject MSUtil.LogQuery.HttpErrorInputFormat}
		"iis"{$inputobj = New-Object -comObject MSUtil.LogQuery.IISIISInputFormat}
		"iisodbc"{$inputobj = New-Object -comObject MSUtil.LogQuery.IISODBCInputFormat}
		"ncsa"{$inputobj = New-Object -comObject MSUtil.LogQuery.IISNCSAInputFormat}
		"netmon"{$inputobj = New-Object -comObject MSUtil.LogQuery.NetMonInputFormat}
		"reg"{$inputobj = New-Object -comObject MSUtil.LogQuery.RegistryInputFormat}
		"textline"{$inputobj = New-Object -comObject MSUtil.LogQuery.TextLineInputFormat}
		"textword"{$inputobj = New-Object -comObject MSUtil.LogQuery.TextWordInputFormat}
		"tsv"{$inputobj = New-Object -comObject MSUtil.LogQuery.TSVInputFormat}
		"urlscan"{$inputobj = New-Object -comObject MSUtil.LogQuery.URLScanLogInputFormat}
		"w3c"{$inputobj = New-Object -comObject MSUtil.LogQuery.W3CInputFormat}
		"xml"{$inputobj = New-Object -comObject MSUtil.LogQuery.XMLInputFormat}
	}
	return $inputobj
}

function Get-LPOutputFormat{
	param([String]$OutputType)
	switch($OutputType.ToLower()){
		"csv"{$outputobj = New-Object -comObject MSUtil.LogQuery.CSVOutputFormat}
		"chart"{$outputobj = New-Object -comObject MSUtil.LogQuery.ChartOutputFormat}
		"iis"{$outputobj = New-Object -comObject MSUtil.LogQuery.IISOutputFormat}
		"sql"{$outputobj = New-Object -comObject MSUtil.LogQuery.SQLOutputFormat}
		"syslog"{$outputobj = New-Object -comObject MSUtil.LogQuery.SYSLOGOutputFormat}
		"tsv"{$outputobj = New-Object -comObject MSUtil.LogQuery.TSVOutputFormat}
		"w3c"{$outputobj = New-Object -comObject MSUtil.LogQuery.W3COutputFormat}
		"tpl"{$outputobj = New-Object -comObject MSUtil.LogQuery.TemplateOutputFormat}
	}
	return $outputobj
}

function Invoke-LPExecute{
	param([string] $query, $inputtype)
    $LPQuery = new-object -com MSUtil.LogQuery
	if($inputtype){
    	$LPRecordSet = $LPQuery.Execute($query, $inputtype)	
	} else {
		$LPRecordSet = $LPQuery.Execute($query)
	}
    return $LPRecordSet
}

function Invoke-LPExecuteBatch{
	param([string]$query, $inputtype, $outputtype)
    $LPQuery = new-object -com MSUtil.LogQuery
    $result = $LPQuery.ExecuteBatch($query, $inputtype, $outputtype)
    return $result
}

 
function Get-LPRecord{
	param($LPRecordSet)
	$LPRecord = new-Object System.Management.Automation.PSObject
	if( -not $LPRecordSet.atEnd()) {
		$Record = $LPRecordSet.getRecord()
		for($i = 0; $i -lt $LPRecordSet.getColumnCount();$i++) {        
			$LPRecord | add-member NoteProperty $LPRecordSet.getColumnName($i) -value $Record.getValue($i)
		}
	}
	return $LPRecord
}

function Get-LPRecordSet{
	param([string]$query)
	$LPRecordSet = Invoke-LPExecute $query
	$LPRecords = new-object System.Management.Automation.PSObject[] 0
	for(; -not $LPRecordSet.atEnd(); $LPRecordSet.moveNext()) {
		$LPRecord = Get-LPRecord($LPRecordSet)
		$LPRecords += new-Object System.Management.Automation.PSObject	
        $RecordCount = $LPQueryResult.length-1
        $LPRecords[$RecordCount] = $LPRecord
	}
	$LPRecordSet.Close();
	return $LPRecords
}

$servername = $env:computername

$infile = (Get-ChildItem -Path C:\inetpub\logs25\SMTPSVC1 | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
$inputformat = Get-LPInputFormat "w3c"

$query = @"
SELECT Date, REVERSEDNS(c-ip) AS Client, COUNT(*) as Total
FROM $infile
WHERE sc-status<>0
GROUP BY Date, Client
ORDER BY Date, Total desc
"@

$records = Get-LPRecordSet $query $inputformat
$date = Get-Date
$attachment = "$($servername)-$($date.ToString("yyyyMMdd")).csv"
$records | Export-CSV -notype -enc utf8 $attachment

$body += "<body><table width=""560"" border=""1""><tr>"
$records[0] | ForEach-Object { 
	foreach ($property in $_.PSObject.Properties){$body += "<td>$($property.name)</td>"}
} 
$body += "</tr><tr>"
$records | ForEach-Object {
	foreach ($property in $_.PSObject.Properties){$body += "<td>$($property.value)</td>"}
	$body += "</tr><tr>"
}
$body += "</tr></table></body>"

Send-MailMessage -Body $body -BodyAsHtml -From $servername@bnquebec.ca -To martin.thomas@banq.qc.ca -SmtpServer courriel.banq.qc.ca -Subject "Log SMTP $($servername) $($date.ToString())" -Attachments $attachment