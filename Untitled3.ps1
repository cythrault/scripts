cls
Function Compare-ObjectProperties {
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject 
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                ODS=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                AD=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }        
    }
    if ($diffs) {return ($diffs | Select PropertyName,ODS,AD)}
}
$attr="sAMAccountName","sn","givenName","displayName","initials","mail","userPrincipalName","telephonenumber","mobile","facsimiletelephonenumber","homephone","ipphone","physicalDeliveryOfficeName","c","co","countryCode","EmployeeType","title","company","department","Manager","extensionAttribute1","extensionattribute2","extensionattribute3","extensionattribute4","extensionattribute5","extensionattribute9","extensionAttribute11","url","wWWHomePage","accountexpires"
$sattr="sAMAccountName","sn","givenName","displayName","initials","mail","userPrincipalName","telephonenumber","mobile","facsimiletelephonenumber","homephone","ipphone","physicalDeliveryOfficeName","c","co","countryCode","EmployeeType","title","company","department",@{Name="Manager"; Expression = {$(get-aduser -Identity $_.manager -properties displayname).displayname}},"extensionAttribute1","extensionattribute2","extensionattribute3","extensionattribute4","extensionattribute5","extensionattribute9","extensionAttribute11","url","wWWHomePage","accountexpires"
$ods=Import-Csv -Encoding utf8 -Delimiter "," -Path J:\ods.csv
$ods | ForEach-Object {
    Write-Progress -activity "exporting AD accounts" -percentComplete ($c++ / $ods.count*100)
    $ad = get-aduser -identity $_.sAMAccountName -Properties $attr | select $sattr
    Compare-ObjectProperties $_ $ad
} | ft -auto