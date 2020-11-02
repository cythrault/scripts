$attr="sAMAccountName","sn","givenName","displayName","initials","mail","userPrincipalName","telephonenumber","mobile","facsimiletelephonenumber","homephone","ipphone","physicalDeliveryOfficeName","c","co","countryCode","EmployeeType","title","company","department","Manager","extensionAttribute1","extensionattribute2","extensionattribute3","extensionattribute4","extensionattribute5","extensionattribute9","extensionAttribute11","url","wWWHomePage","accountexpires"
$sattr="sAMAccountName","sn","givenName","displayName","initials","mail","userPrincipalName","telephonenumber","mobile","facsimiletelephonenumber","homephone","ipphone","physicalDeliveryOfficeName","c","co","countryCode","EmployeeType","title","company","department",@{Name="Manager"; Expression = {$(get-aduser -Identity $_.manager -properties displayname).displayname}},"extensionAttribute1","extensionattribute2","extensionattribute3","extensionattribute4","extensionattribute5","extensionattribute9","extensionAttribute11","url","wWWHomePage","accountexpires"
$ods=Import-Csv -Encoding utf8 -Delimiter "," -Path J:\ods.csv
$output="ods-upd-$(get-date -uformat %Y%m%d).csv"
$c=0
$ods | ForEach-Object {
    Write-Progress -activity "exporting AD accounts" -percentComplete ($c++ / $ods.count*100)
    If ( $( Get-ADUser -LDAPFilter "(sAMAccountName=$($_.sAMAccountName))" ) -eq $Null ) {
        Write-Host "$($_.sAMAccountName) User does not exist in AD"
        Write-Output "$($_.sAMAccountName) User does not exist in AD"
        }
    Else {
        get-aduser -identity $_.sAMAccountName -Properties $attr
        }
    } | select-object $sattr | Export-Csv -NoTypeInformation -Encoding utf8 -Path $output