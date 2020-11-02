$user = "carole.matte"
$user = "florian.daveau"

$user = "NATHALIE.CHAMBERLAND"
$gcserver = (Get-ADForest).DomainNamingMaster
$ADUser = Get-ADUser -Server $gcserver`:3268 -Identity $user
[string]$domain = switch -wildcard ( $ADUser.DistinguishedName ) {
    "*DC=archives,DC=bnquebec,DC=ca" { "archives.bnquebec.ca" }
    "*DC=biblio,DC=bnquebec,DC=ca" { "biblio.bnquebec.ca" }
    default { "bnquebec.ca" }
}
[string]$dc = (Get-ADDomainController -DomainName $domain -Discover).HostName

if ( !((Get-Mailbox $user -DomainController $dc).ArchiveDatabase)) {
    Enable-Mailbox $user -Archive -DomainController $dc
}

$folder = "\\bnquebec.ca\banq\Archivage\$($user)"
$PSTs = Get-ChildItem $folder -Filter *.pst
$PSTs | Select-Object -Property Name, @{Name="sizeGB";Expression={[math]::round(($_.length/1gb),1)}}, LastWriteTime | Sort-Object -Property LastWriteTime -Descending

$openPST = Invoke-Command -ComputerName pvwfsp0h -ScriptBlock {Get-SmbOpenFile | Where-Object{$_.Path -like "*.pst"}}

$PSTs.fullname.replace("\\bnquebec.ca\banq\Archivage\","")|ForEach-Object{
    if ($openPST.ShareRelativePath -contains $PSItem) {
        write-host "$($PSItem) is openned!"
    }
}

$acl = Get-Acl $folder
$identity = "BNQUEBEC\Exchange Trusted Subsystem"
$fileSystemRights = "Modify"
$type = "Allow"
$fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $type
$fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
$ACL.SetAccessRule($fileSystemAccessRule)
Set-Acl -Path $folder -AclObject $ACL

ForEach ($pst in $PSTs) {
    Write-Host "Importing $($pst.FullName)"
    New-MailboxImportRequest $user -FilePath $pst.FullName -IsArchive -DomainController $dc -BadItemLimit 100 -Confirm:$true
}
Get-MailboxImportRequest

# Get-MailboxImportRequest -Name MailboxImport|Get-MailboxImportRequestStatistics -IncludeReport|fl *