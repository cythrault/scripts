function SyncDFS {

	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$dossier,
		[bool]$reverse = $False
	)

	Begin {
		#$remote = "\\pbwfsp0t.bnquebec.ca\d$\"
		#$local = "D:\"
		$remote = "c:\test"
		$local = "c:\test2"
		#$options = @("/E","PURGE","/B","/COPYALL","/R:1","/W:5","/MT:8")
		$options = @("/R:1","/W:5","/MT:8")
		$logpath = "c:\ops"
	}

	Process {
		$start = (Get-Date)
		
		if ( $dossier.Contains("\") ) {
			$logfile = "$($logpath)\robocopy-$($dossier.Replace(' ','').Split("\")[1])-$($start.ToString("yyyyMMddHHMMss")).log"
		} else {
			$logfile = "$($logpath)\robocopy-$($dossier.Replace(' ',''))-$($start.ToString("yyyyMMddHHMMss")).log"
		}
		
		if ( $reverse ) { 
			Write-Host "Reverse"
			$src = "$($remote)\$($dossier)"; $dest = "$($local)\$($dossier)"
		} else {
			$src = "$($local)\$($dossier)"; $dest = "$($remote)\$($dossier)"
		}

		robocopy "$src" "$dest" @options "/LOG:$logfile"
		
		$end = (Get-Date)
		$elapsed = "Elapsed Time: $(($end-$start).totalminutes) minutes"
	}
	
}	

$dossiers = @()
$dossiers += "ANQ"
$dossiers += "Archivage"
$dossiers += "Images\Archives Audio-Visuel"
$dossiers += "Autre_en_Transit"
$dossiers += "BAnQ"
$dossiers += "Tous\Direction des communications"
$dossiers += "Data"
$dossiers += "Images\ImagesCAM"
$dossiers += "Images\ImagesCAQ"
$dossiers += "Perso"
$dossiers += "Sqla\SQLA_IA"
$dossiers += "Tous"
$dossiers += "TousGeneral"

$dossiers | SyncDFS #-Reverse $True