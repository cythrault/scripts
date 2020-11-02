#load assembly needed for AD lookup
add-type -assemblyname system.directoryservices.accountmanagement

$users = get-aduser -filter * -Server corp.pbwan.net -properties sidhistory | ?{$_.sidhistory -ne ""} | select name, samaccountname, distinguishedname, sidhistory

#hashtable containing the pbwan trusts
$sidHash = @{}
$sidHash["S-1-5-21-527237240-1500820517-725345543"] = "corp.pbwan.net"
$sidHash["S-1-5-21-210520846-1106061574-7473742"] = "halsall.com"
$sidHash["S-1-5-21-1394335539-3719247576-210320554"] = "dmz.local"
$sidHash["S-1-5-21-1757981266-1085031214-839522115"] = "pbwan.net"
$sidHash["S-1-5-21-555392383-2828798492-3464129327"] = "WSPUSGov.com"
$sidHash["S-1-5-21-129462270-2139338679-4080918649"] = "MGWSP.CO.UK"
$sidHash["S-1-5-21-2420330522-2884553087-655016459"] = "wspgroup.com"
$sidHash["S-1-5-21-1804815775-1190227046-625696398"] = "gcg.local"
$sidHash["S-1-5-21-77810565-118882789-1848903544"] = "focuscorp.ca"
$sidHash["S-1-5-21-220523388-1214440339-725345543"] = "cra.com"
$sidHash["S-1-5-21-73586283-602162358-682003330"] = "ccrd.com"
$sidHash["S-1-5-21-1332025955-1498798356-211625990"] = "mmm.ca"
$sidHash["S-1-5-21-1111383825-1399753330-1979989523"] = "uk.wspgroup.com"
$sidHash["S-1-5-21-2996190024-3608494798-319437263"] = "WSPGROUP.NET"
$sidHash["S-1-5-21-768029164-11021426-988572150"] = "smithcarter.com"
$sidHash["S-1-5-21-4281975367-2065093054-4284754087"] = "ca.wspgroup.com"
$sidHash["S-1-5-21-2560413578-2447596656-1060672358"] = "terraingroup.com"
$sidHash["S-1-5-21-1437618765-291011396-617630493"] = "se.wspgroup.com"
$sidHash["S-1-5-21-3255023295-237464038-3251150289"] = "vaughan.splconsultants.ca"
$sidHash["S-1-5-21-910836297-523650652-2076119496"] = "richmond.levelton.com"
$sidHash["S-1-5-21-1064525262-3820111905-3418901292"] = "wspgeo.local"

$finallist = @()
foreach($user in $users){
    
    foreach($userSidhistory in $user.sidhistory){
        $obj = new-object psobject | select displayname, samaccountname, Sidhistory, LinkedDomain, linkedDomainSAN, linkedDomainDN
        $obj.displayname = $user.name
        $obj.samaccountname = $user.samaccountname    
        $obj.sidhistory = $userSidhistory.value

        #convert sidhistory to domain name
        $domainSid = $sidhash[$($usersidhistory.value.Substring(0,$usersidhistory.value.lastindexof('-')))]
        $obj.linkedDomain = $domainSid

        #if domain is not found
        if($obj.linkedDomain -eq $null){
            $obj.linkedDomain = "Domain Not found"
            $obj.linkedDomainSAN = "Domain not found"
            $obj.linkedDomainDN = "Domain not found"
        }
    }
    $finallist += $obj
}
#sort list by domain to simplify queries
$finallist = $finallist | sort linkedDomain

$cpt = 0
$tempDomain = ""
$ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain
$idtype = [System.DirectoryServices.AccountManagement.IdentityType]::sid

foreach($obj in $finallist){

        if($obj.LinkedDomain -ne "Domain Not found"){

            try{
                $domain = $obj.linkedDomain
                $pc = New-Object System.DirectoryServices.accountManagement.principalContext($ct,$domain)
                $aduser = [System.DirectoryServices.accountmanagement.userprincipal]::findbyidentity($pc,$idtype,$obj.sidhistory)
    
    
                $obj.linkedDomainSAN = $aduser.SamAccountName
                $obj.linkedDomainDN = $aduser.distinguishedname
            }
            catch [Exception]{
                $obj.linkedDomainSAN = "Error during Query"
                $obj.linkedDomainDN = "Error during Query"
            }
            $finallist[$cpt] = $obj
        }
        $cpt++
}

#we then scan which linked Sidhistories could create problems by having different SamAccountNames
$finalFiltered = $finallist | ?{$_.linkedDomain -ne "Domain Not found"}
$odditieslist = @()
foreach($obj in $finalFiltered){
    if($obj.samaccountname -ne $obj.linkedDomainSan){
        $odditieslist += $obj
    }
}
$finallist | Export-Csv "C:\temp\list.csv" -Delimiter ';' -NoTypeInformation
$csvFile = gci "C:\temp\list.csv"
$recipients = @("marc.plamondon@wspgroup.com","martin.thomas@wspgroup.com","mario.cardinal@wspgroup.com")
$smtpServer = "smtp-ca.wspgroup.com"

#HTML STyle for formatting
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"

if($odditieslist -ne $null){

$htmlbody = $odditieslist | ConvertTo-Html -Head $style
$htmlbody = "<font color=`"red`"><br>Here is the list of unmatching SamAccountNames</font><br><br>" + $htmlbodyfirst
}
else{

$htmlbody = "<font color=`"green`"><br>There are no unmatching SamAccountNames</font><br><br>"
}
Send-MailMessage -BodyAsHtml -Body $htmlbody -From "no-reply@wspgroup.com" -To $recipients -Attachments $csvFile -SmtpServer  -Subject "sidhistory oddities list"

