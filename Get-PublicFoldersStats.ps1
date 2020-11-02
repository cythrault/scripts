
Write-Progress -Activity "Finding Public Folders" -Status "running get-publicfolders -recurse"
$folders = get-publicfolder -recurse -resultsize unlimited
$arFolderData = @()
$totalfolders = $folders.count
$i = 1

foreach ($folder in $folders) {
	$statusstring = "$i of $totalfolders"
	write-Progress -Activity "Gathering Public Folder Information" -Status $statusstring -PercentComplete ($i/$totalfolders*100)
	$folderstats = get-publicfolderstatistics $folder.identity
	$folderdata = new-object Object
	$folderdata | add-member -type NoteProperty -name FolderName $folder.name
	$folderdata | add-member -type NoteProperty -name FolderPath $folder.identity
	$folderdata | add-member -type NoteProperty -name LastAccessed $folderstats.LastAccessTime
	$folderdata | add-member -type NoteProperty -name LastModified $folderstats.LastModificationTime
	$folderdata | add-member -type NoteProperty -name Created $folderstats.CreationTime
	$folderdata | add-member -type NoteProperty -name ItemCount $folderstats.ItemCount
	$folderdata | add-member -type NoteProperty -name Size $folderstats.TotalItemSize
	$folderdata | Add-Member -type NoteProperty -Name Mailenabled $folder.mailenabled

	if ($folder.mailenabled) {
		#since there is no guarentee that a public folder has a unique name we need to compare the PF's entry ID to the recipient objects external email address
		$entryid = $folder.entryid.tostring().substring(76,12)
		# $primaryemail = (get-recipient -filter "recipienttype -eq 'PublicFolder'" -resultsize unlimited) | where {$_.externalemailaddress -like "*$entryid"}).primarysmtpaddress
		$primaryemail = (get-recipient $folder.mailrecipientguid.tostring()).primarysmtpaddress
		$folderdata | add-member -type NoteProperty -name PrimaryEmailAddress $primaryemail
		} else {
		$folderdata | add-member -type NoteProperty -name PrimaryEmailAddress "Not Mail Enabled"
		}

	if ($folderstats.ownercount -gt 0) {
		$owners = get-publicfolderclientpermission $folder.identity | where {$_.accessrights -like "*owner*"}
		$ownerstr = ""
		foreach ($owner in $owners) {
			$ownerstr += $owner.user.displayname + ","
			}
		} else {
		$ownerstr = ""
		}
		$folderdata | add-member -type NoteProperty -name Owners $ownerstr
		$arFolderData += $folderdata
		$i++
	}

$arFolderData | export-csv -path PublicFolderData.csv -enc utf8 -notypeinformation