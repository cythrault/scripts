# %windir%\System32\WindowsPowerShell\v1.0\powershell.exe
# -Noninteractive -ExecutionPolicy Bypass –Noprofile -file %logonserver%\netlogon\PstUsr.ps1

$LogPath = "\\bnquebec.ca\BAnQ\logsession\log\session"
$UserName = $env:UserName
$ComputerName = $env:ComputerName
$UserLog = "$($LogPath)\NOMS\$($UserName).txt"
$MachineLog = "$($LogPath)\MACHINES\$($ComputerName).txt"
$date = $(Get-Date -UFormat %Y-%m-%d)
$time = $(Get-Date -UFormat %H:%M:%S)
$NewFileHeader = "CRÉATION DU FICHIER LE $($date) À $($time)"
$BAnQIP = (Test-Connection $ComputerName -count 1).Ipv4Address.IPAddressToString
$LServer = ($env:LOGONSERVER).Replace("\\","")
[string]$LogString = $BAnQIP + [char]9 + $date + [char]9 + $time + [char]9 + "LogonServer:" + $LServer + [char]9 + $ComputerName + [char]9 + $UserName

if ( !(Test-Path -Path $UserLog -PathType Leaf) ) { $NewFileHeader | Out-File -FilePath $UserLog -Encoding UTF8 }
if ( !(Test-Path -Path $MachineLog -PathType Leaf) ) { $NewFileHeader | Out-File -FilePath $MachineLog -Encoding UTF8 }

$LogString | Out-File -FilePath $UserLog -Encoding UTF8 -Append
$LogString | Out-File -FilePath $MachineLog -Encoding UTF8 -Append

$catalogs = Get-ChildItem -Path HKCU:\Software\Microsoft\Office -Recurse | Where-Object{$_.Property -like "*.pst"}
if ($catalogs) {
	$PSTLog = "$($LogPath)\PST\$($UserName).csv"
	$report = $catalogs.Property | Where-Object{$_ -like "*.pst"} | Sort-Object -Unique | ForEach-Object{
        #$pstprops = 
        Get-ChildItem $_ -ErrorAction SilentlyContinue
		if ($pstprops) {
			$pst = New-Object psobject -Property @{
				Username = $env:UserName;
				ComputerName = $env:ComputerName;
				FullName = $_;
				LastWriteTime = $pstprops.LastWriteTime;
				Length = $pstprops.Length
			}
		}
		$pst
	}
	if ($report) { $report | Export-Csv -Path $PSTLog -Encoding UTF8 -NoTypeInformation }
}


