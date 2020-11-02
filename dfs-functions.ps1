# $dfstargets=Get-DfsnRoot bnquebec.ca |%{Get-DfsnFolder -Path "$($_.Path)\*"|%{Get-DfsnFolderTarget $_.Path | Select Path,TargetPath,State}}
# $dfstargets|Get-DFSTargetUsage|Export-Excel dfs-banq.xlsx

function Write-Log { 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path='C:\Logs\PowerShellLog.log', 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
}

function Get-Site {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$Computer
		)
	
	Process {
		$ip = 0
		if ([net.ipaddress]::tryparse($computer,[ref]$ip)) {
			[net.ipAddress]$ip=$ip.IPAddressToString
		} else {
			trap{ $ip = $null; continue }
			[net.ipAddress]$ip = (Resolve-DNSName $Computer -ErrorAction Stop).IPAddress
		}
		if ($ip) {

			# Resolve possible NetworkNumbers
			for ($bit = 30 ; $bit -ge 1; $bit--) {
				[int]$octet = [math]::Truncate(($bit - 1 ) / 8)
				$net = [byte[]]@()
				for($o=0;$o -le 3;$o++) {
					$ba = $ip.GetAddressBytes()
					if ($o -lt $octet) {
						$net += $ba[$o]
					} elseif ($o -eq $octet) {
						$factor = 8 + $octet * 8 - $bit
						$divider = [math]::pow(2,$factor)
						$value = $divider * [math]::Truncate($ba[$o] / $divider)
						$Net += $value
					} else {
						$net += 0
					}
				}
			$NetWork = [string]::join('.',$net) + "/$bit"

			# try to find Subnet in AD
			if ($verbose.IsPresent) {write-host -fore 'yellow' "Trying : $network"}
			$de = New-Object directoryservices.directoryentry('LDAP://rootDSE')
			$Root = New-Object directoryservices.directoryentry("LDAP://$($de.configurationNamingContext)")
			$ds = New-Object directoryservices.directorySearcher($root)
			$ds.filter = "(CN=$NetWork)"
			$r = $ds.findone()
   
			# If subnet found, write info and exit script.
			if ($r) {
				[pscustomobject]@{Host=$Computer; IP=$ip; Site=($r.Properties.siteobject -split ',*..=')[1]; Subnet=$r.properties.name[0]; SiteDescription=$r.properties.description[0]}
				break
				}
			}
		}
		else {
			write-host -fore 'red' "Could not resolve $computer"
			[pscustomobject]@{Host=$Computer; IP=$null; Site=$null; Subnet=$null; SiteDescription=$null}
			}
	}
}

function Get-ComputerSite($ComputerName) {
   $site = nltest /server:$ComputerName /dsgetsite 2>$null
   if($LASTEXITCODE -eq 0){ $site[0] }
}

function Get-DirectorySizeWithRobocopy {
	Param (
		[Parameter(Mandatory=$true)]
		[string]$folder
	)

	Process {
		$fileCount = 0
		$totalBytes = 0
		robocopy /l /nfl /ndl $folder \\localhost\C$\nul /e /bytes | ?{ $_ -match "^[ \t]+(Files|Bytes) :[ ]+\d" } | %{
				$line = $_.Trim() -replace ' :',':' -replace '[ ]{1,}',','
				$value = $line.split(',')[1]
				if ( $line -match "Files:" ) { $fileCount = $value } else { $totalBytes = $value }
				}
		[pscustomobject]@{Path=$folder;Files=$fileCount;GBytes=[math]::Round(($totalBytes/1073741824),2);Bytes=$totalBytes}
		}
}
	
function Get-DFSTargetUsage {

	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$Path,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$TargetPath,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$State
		)
	Begin {
		$i=0
		$total = $dfstargets.count
		$startTime = Get-Date
		}
	Process {
		$i++
		$elapsedTime = New-TimeSpan -Start $startTime -End (Get-Date)
		$timeRemaining = New-TimeSpan -Seconds (($elapsedTime.TotalSeconds / ($i / $total)) - $elapsedTime.TotalSeconds)
		Write-Progress -Activity "Scanning DFS Targets - $i of $total - Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -PercentComplete $( $i / $total * 100 ) -SecondsRemaining $timeRemaining.TotalSeconds -Status "Path: $Path - TargetPath: $TargetPath"
		$computer=$TargetPath.Split("\")[2]
		$share=$TargetPath.Split("\")[3]
		try {
			$localsharepath=(Get-WmiObject -Class Win32_Share -ComputerName $computer -ErrorAction Stop | ? {$_.Name -eq $share}).Path
			$localpath=$TargetPath.Replace("\\$($computer)\$($share)",$localsharepath)
			$ComputerSite = Get-ComputerSite $computer
			}
		catch {
			Write-Host -ForegroundColor Red "Could not connect to $($computer) to get shares information."
			$output = [pscustomobject]@{Path=$Path;TargetPath=$TargetPath;State=$state;LocalPath=$null;Files=$null;GBytes=$null}
			}
		try {
			Write-Progress -Activity "Scanning DFS Targets - $i of $total - Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -PercentComplete $( $i / $total * 100 ) -SecondsRemaining $timeRemaining.TotalSeconds -Status "Path: $Path - TargetPath: $TargetPath" -CurrentOperation "Computer: $Computer - Share: $share - LocalPath: $localpath"
			if ( $localsharepath ) {
				$remotedata=Invoke-Command -ComputerName $computer -ScriptBlock ${Function:Get-DirectorySizeWithRobocopy} -ArgumentList $localpath -ErrorAction Stop
				$output = [pscustomobject]@{Path=$Path;TargetPath=$TargetPath;State=$state;LocalPath=$remotedata.Path;Files=$remotedata.Files;GBytes=$remotedata.GBytes;Computer=$computer;IP=(Resolve-DNSName -Type A -Name $Computer).IPAddress;Subnet=$ComputerSite.Subnet;Site=$ComputerSite.Site;SiteDescription=$ComputerSite.SiteDescription}
				}
			else {
				Write-Host -ForegroundColor Red "Share $($share) does not exist on $($computer)"
				$output = [pscustomobject]@{Path=$Path;TargetPath=$TargetPath;State=$state;LocalPath=$null;Files=$null;GBytes=$null}
				}
			}
		catch {
			Write-Host -ForegroundColor Red "Could not connect to $($computer) to get $($share) information from local path - $($localpath)"
			$output = [pscustomobject]@{Path=$Path;TargetPath=$TargetPath;State=$state;LocalPath=$null;Files=$null;GBytes=$null}
			}
		$output
		}
	End {
		Write-Host "Elapsed Time: $($elapsedTime.ToString('hh\:mm\:ss'))"
		}
	}