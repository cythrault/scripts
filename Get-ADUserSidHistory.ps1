Get-ADObject -LDAPFilter "(sIDHistory=*)" -Property objectClass, distinguishedname, samAccountName, objectSID, sIDHistory |
Select-Object objectClass, DistinguishedName, SamAccountName, objectSID -ExpandProperty SIDHistory |
ForEach-Object {
   [PSCustomObject]@{
      ObjectClass       = $_.objectClass
      DistinguishedName = $_.DistinguishedName
      SamAccountName    = $_.SamAccountName
      SID               = $_.ObjectSID
      DomainSID         = $_.AccountDomainSID
      SIDHistory        = $_.Value
   }
} |
Out-GridView -Title "AD Objects with SIDHistory"