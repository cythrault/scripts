function Get-FreeSkypeNumber {
    Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("Site Code")]
		[string]
        $sitecode
    )
    $sn=import-csv -path C:\inetpub\ADUsers\App_Data\listUnassigned.csv
    $numberrange=@()
    $assignedto=$null
    ($sn -match $sitecode)|%{
        [int]$areacode = $_.NumberRangeStart.TrimStart("tel:+1").SubString(0,3)
        [int]$start = $_.NumberRangeStart.TrimStart("tel:+1").SubString(3,7)
        [int]$end = $_.NumberRangeEnd.TrimStart("tel:+1").SubString(3,7)
        $numberrange += $start..$end
    }
    foreach ($suffix in $numberrange ) {
        [int64]$number="$areacode$suffix"
        $ldapfilter = "(|(proxyAddresses=eum:$number*)(msRTCSIP-Line=tel:+1$number))"
        $assignedto = (get-aduser -LDAPFilter $ldapfilter).name
        if ( -Not $assignedto ) { break }
    }
    return $number
}

#Get-FreeSkypeNumber "CAMTR400"
$return=$null
$return=Get-FreeSkypeNumber "CAMTR400"
write-host $return
$return.GetType()