$sam = "pchabot"

$pdc = ([System.Directoryservices.Activedirectory.Forest]::GetCurrentForest()).RootDomain.PDCRoleOwner.Name
$user = Get-ADUser -Server $pdc`:3268 -Identity $folder.Name -ErrorAction SilentlyContinue

[string]$dc = switch -wildcard ( $user.DistinguishedName ) {
    "*DC=archives,DC=bnquebec,DC=ca" { (Get-ADDomain archives.bnquebec.ca).PDCEmulator }
    "*DC=biblio,DC=bnquebec,DC=ca" { (Get-ADDomain biblio.bnquebec.ca).PDCEmulator }
    default { (Get-ADDomain bnquebec.ca).PDCEmulator }
}

ADMT USER /N $sam /IF:YES /SD:$sourcedomain /TD:bnquebec.ca /TO:"1.gestion/usagers/employes" /MigrateGroups:No /UUR:Yes    

$user = Get-ADUser -Server $((Get-ADDomain bnquebec.ca).PDCEmulator) -Identity $sam -ErrorAction SilentlyContinue

Set-ADUser -Identity $user -ChangePasswordAtLogon $false
Set-ADUser -Identity $user -Clear scriptPath
$newUPN = $sam + '@banq.qc.ca'
Set-ADuser -Identity $user -UserPrincipalName $newUPN  

Invoke-Command -ComputerName pvwaad0a -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

$ename = $employee.identity
$sname = $employee."Sam Account"
$lsname = (Get-Culture).TextInfo

$query="BEGIN TRANSACTION;
    UPDATE
        aspnetUsers
    SET
        Username = 'BNQUEBEC\$($lsname.ToTitleCase($sname))',
        NormalizedUserName = 'BNQUEBEC\$($sname.ToUPPER())'
    WHERE
        UserName like '%\$sname'
    UPDATE
        Clients
    SET
        Username = 'BNQUEBEC\$($lsname.ToTitleCase($sname))'
    WHERE
        UserName like '%\$sname'
    COMMIT;"

$params = @{'server'='AGPROD01\SQLPROD01';'Database'='C2_V5'}
Invoke-Sqlcmd @params -Query $query -Username admigration -Password P@ssw0rd1
$params = @{'server'='AGDEV01\SQLDEV01';'Database'='C2_V5_ACCEPT'}
Invoke-Sqlcmd @params -Query $query -Username admigration -Password P@ssw0rd1
