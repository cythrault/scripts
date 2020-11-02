#C:\tools\scripts\Check-ConflictDeleted.ps1

$FilterHashTable = @{
    LogName   = 'DFS Replication'
    ProviderName= 'DFSR' 
    ID        = 4412
    StartTime = (Get-Date).AddDays(-14)
    EndTime   = Get-Date
}

$events = Get-WinEvent -FilterHashtable $FilterHashTable | ForEach-Object {
    $Values = $_.Properties | ForEach-Object { $_.Value }
    [PSCustomObject]@{
        Time      = $_.TimeCreated
        "Original File Path"			=	$Values[1]
		"New Name in Conflict Folder"	=	$Values[8]
        "Replicated Folder Root"		=	$Values[2]
        "File ID"						=	$Values[3]
        "Replicated Folder Name"		=	$Values[4]
        "Replicated Folder ID"			=	$Values[0]
        "Replication Group Name"		=	$Values[5]
        "Replication Group ID"			=	$Values[6]
        "Member ID"						=	$Values[7]
        "Partner Member ID"				=	$Values[9]
    }
}

ForEach ( $event in $events ) {
    $file = $event."Original File Path"
    $folderRoot = $event."Replicated Folder Root"
    $folderName = $event.Replicated Folder Name"

    if ( !(Test-Path -Path $file -PathType Leaf) ) {
        Write-Host -BackgroundColor Black -ForegroundColor Red "$($file) not found."

        $event

        $volume = $file.Split("\")[0]
        $restoreFolder = "DFSRRestore\$($folderName)"
        $restorePath = $volume + [char]92 + $restoreFolder
        if ( !(Test-Path -Path $restorePath -PathType Container) ) { New-Item -Path $volume -Name $restoreFolder -ItemType Directory }

        $manifestPath = $folderRoot + "\DfsrPrivate\ConflictAndDeletedManifest.xml"
        $manifestPath
        $restorePath
        
        Restore-DfsrPreservedFiles -Path $manifestPath -RestoreToPath $restorePath -CopyFiles -Verbose -WhatIf




    }
}