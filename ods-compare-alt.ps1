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
                User= $ReferenceObject.samAccountName
                PropertyName=$objprop
                ODS=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                AD=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }        
    }
    if ($diffs) {return ($diffs | Select User,PropertyName,ODS,AD)}
}
$attr="sAMAccountName","givenName","sn","mail","title"
$sattr=$attr
$ods=Import-excel -Path "J:\ods\Employee Extract using GPD View (Feb 26, 2018).xlsx"
$output="ods-upd-alt-$(get-date -uformat %Y%m%d).csv"
$c=0
$results = $ods | ForEach-Object {
    Write-Progress -activity "Comparing ODS/AD accounts properties" -percentComplete ($c++ / $ods.count*100) -CurrentOperation $_.sAMAccountName -Status "Processing..."
    If ( $( Get-ADUser -LDAPFilter "(sAMAccountName=$($_.sAMAccountName))" ) -eq $Null ) {
        Write-Host "$($_.sAMAccountName) User does not exist in AD"
        $props = $_ | Get-Member -MemberType Property,NoteProperty | % Name
        foreach ($prop in $props) { $ad.$prop = "User Not Found" }
        Compare-ObjectProperties $_ $ad
        }
    Else {
        $ad = get-aduser -identity $_.sAMAccountName -Properties $attr | select $sattr
        Compare-ObjectProperties $_ $ad
        }
}
$results | Export-Csv -NoTypeInformation -Encoding utf8 -Path $output
$results | Export-Excel -Path "ods-upd-alt-$(get-date -uformat %Y%m%d).xlsx"