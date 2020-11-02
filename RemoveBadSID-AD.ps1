param ($Action,$folder,$Opt)

$Forest=Get-ADRootDSE
$Domain=(Get-ADDomain).distinguishedname
$Conf  =$Forest.configurationNamingContext
$Schema=$Forest.SchemaNamingContext
$ForestName=$Forest.rootDomainNamingContext
$DomainDNS ="DC=DomainDnsZones,$ForestName"
$ForestDNS ="DC=ForestDnsZones,$ForestName"

$domsid=(Get-ADDomain).domainsid.tostring()

if     (($Action)-and($Action.ToUpper() -like "/LIST"))     { $Remove=$False; $OU=$False }
elseif (($Action)-and($Action.ToUpper() -like "/LISTOU"))   { $Remove=$False; $OU=$True }
elseif (($Action)-and($Action.ToUpper() -like "/REMOVE"))   { $Remove=$True; $OU=$False }
elseif (($Action)-and($Action.ToUpper() -like "/REMOVEOU")) { $Remove=$True; $OU=$True }
else {
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "SYNTAX: RemoveBadSID-AD.ps1 [/LIST|/REMOVE|/LISTOU|/REMOVEOU[/DOMAIN|/CONF|/SCHEMA|/DOMAINDNS|/FORESTDNS|dn[/RO|/SP]]]"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "PARAM1: /LISTOU List only CNs&OUs /LIST List all objects, /REMOVE Clean all objects /REMOVEOU Clean only CNs&OUs"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "PARAM2: /DOMAIN Actual domain /CONF Conf. Part./SCHEMA /DOMAINDNS /FORESTDNS or a specific DN between double-quotes"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "OPTION1: /RO lists/Removes only objects with unknown SIDs of the domain"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "OPTION2: /SP lists access permissions for all analyzed objects"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "If no DN is indicated, the current domain will be used"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' "SAMPLE1 : RemoveBadSID-AD.ps1 /REMOVEOU /DOMAIN /RO"
  Write-host -backgroundcolor 'White' -Foregroundcolor 'Blue' 'SAMPLE2 : RemoveBadSID-AD.ps1 /LIST "OU=MySite,DC=Domain,DC=local"'
  Break
  }

if     (($Folder)-and($Folder.ToUpper() -like "/CONF"))      { $Folder=$Conf }
elseif (($Folder)-and($Folder.ToUpper() -like "/SCHEMA"))    { $Folder=$Schema }
elseif (($Folder)-and($Folder.ToUpper() -like "/DOMAIN"))    { $Folder=$Domain }
elseif (($Folder)-and($Folder.ToUpper() -like "/DOMAINDNS")) { $Folder=$DomainDNS }
elseif (($Folder)-and($Folder.ToUpper() -like "/FORESTDNS")) { $Folder=$ForestDNS }
elseif (($Folder)-and($Folder.ToUpper() -match "DC=*"))      { Write-output "This DistinguishedName will be analyzed : $Folder" }
else   { $folder=$domain; Write-output "This current domain will be analyzed : $Domain"}

Write-output "Analyzing the following object : $Folder"

if (($Opt)-and($Opt.ToUpper() -like "/RO")) { $Show=$False } Else { $Show=$True }
if (($Opt)-and($Opt.ToUpper() -like "/SP")) { $ShowPerms=$True } Else { $ShowPerms=$False }

# Functions list

function RemovePerms($fold) {
$f=get-item AD:$fold
$fName=$f.distinguishedname
If ($Show) { write-output $fname }
$x=[System.DirectoryServices.ActiveDirectorySecurity](get-ACL AD:$f)
if ($ShowPerms) { write-output $x.access |sort -property IdentityReference -unique |ft -auto IdentityReference,IsInherited,AccessControlType,ActiveDirectoryRights}
$mod=$false
$OldSID=""

foreach ($i in $x.access) {
  if ($i.identityReference.value.tostring() -like "$domsid*"){
    $d=$i.identityReference.value.tostring()
    if ($OldSid -ne $d) { write-output "Unknown SID $d on $fname"; $OldSid=$d  }
    if ($Remove) { $x.RemoveAccessRuleSpecific($i) ; $Mod=$True}
    }
#   elseif ($i.identityReference.value.tostring() -like "S-1-5-21-*") { write-output "SID inconnu $d sur $fname"}
  }
#write-output $x.access
if ($mod){ Set-ACL -aclobject $x -path AD:$f; write-output "Unknown SID cleaned on $fname" }
}

Function RecurseFolder($fold)
{
  $f=$fold
#  If ($Show) { write-output $f }
  If ($OU) { $ListFold=get-childitem AD:$f -force |where { ($_.ObjectClass -like "container")-or($_.ObjectClass -like "OrganizationalUnit") }}
  Else { $ListFold=get-childitem AD:$f -force }
  foreach ($e in $ListFold){
     $FD=$e.Distinguishedname
#     write-output $FD
     RemovePerms $FD     
  }
#  $ListFold=get-childitem AD:$f -force
# |where { ($_.ObjectClass -like "container")-or($_.ObjectClass -like "OrganizationalUnit") }
  foreach ($e in $ListFold){ RecurseFolder($e.Distinguishedname) }
}

# Début du programme

RemovePerms($Folder)
RecurseFolder($Folder)
