function AddOU($DN)
{
if ( ([adsi]::Exists("LDAP://$DN")) -ne "True" ) {

# A regex to split the DN, taking escaped commas into account
$DNRegex = '(?<![\\]),'

# We'll need to traverse the path, level by level, let's figure out the number of possible levels 
$Depth = ($DN -split $DNRegex).Count
# Step through each possible parent OU
for($i = 1;$i -le $Depth;$i++)
{
    $NextOU = ($DN -split $DNRegex,$i)[-1]
    if($NextOU.IndexOf("OU=") -ne 0 -or [ADSI]::Exists("LDAP://$NextOU")) {
    }
    else
    {
        # OU does not exist, remember this for later
        [String[]]$MissingOUs += $NextOU
    }
}

# Reverse the order of missing OUs, we want to create the top-most needed level first
[array]::Reverse($MissingOUs)

# Now create the missing part of the tree, including the desired OU
foreach($OU in $MissingOUs)
{
write "for each"
    $newOUName = (($OU -split $DNRegex,2)[0] -split "=")[1]
    $newOUPath = ($OU -split $DNRegex,2)[1]
 
	New-ADOrganizationalUnit -Name $newOUName -Path $newOUPath -Verbose

}
}
}
