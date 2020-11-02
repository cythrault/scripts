$Location = "CA"
$Environment = "p" 
$MSASeqNumber = "08"
$GroupDescription = "Prod SCOM"
$ComputerList = @("DCCAN100SQL22$", "DCCAN100SQL23$")
$ADGroupOU = "OU=Groups,OU=Enterprise,DC=corp,DC=pbwan,DC=net"
$IsGMSA = $false

####################################################

    $ADGroupName = "GRP-MSA-$Location-SQL $GroupDescription"
    $MSAServer = "g$($Location)$($Environment)SQL$($MSASeqNumber)SVC"
    $MSAAgent = "g$($Location)$($Environment)SQL$($MSASeqNumber)AGT"

    New-ADGroup –name $ADGroupName –groupscope DomainLocal –path $ADGroupOU -Description "gMSA Group for $($GroupDescription)"

    ForEach ($Computer in $ComputerList)
    {
         Add-ADGroupMember -Identity $ADGroupName -Members "$Computer"
    }

    New-ADServiceAccount -Name $MSAServer -DNSHostName "$($MSAServer).corp.pbwan.net" -PrincipalsAllowedToRetrieveManagedPassword $ADGroupName -Description $ADGroupName -Enabled $true
    New-ADServiceAccount -Name $MSAAgent -DNSHostName "$($MSAAgent).corp.pbwan.net" -PrincipalsAllowedToRetrieveManagedPassword 