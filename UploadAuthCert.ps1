Connect-MsolService
$CertFile = "$env:SYSTEMDRIVE\OAuthConfig\OAuthCert.cer"
$objFSO = New-Object -ComObject Scripting.FileSystemObject
$CertFile = $objFSO.GetAbsolutePathName($CertFile)
$cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate
$cer.Import($CertFile)
$binCert = $cer.GetRawCertData()
$credValue = [System.Convert]::ToBase64String($binCert)
$ServiceName = "00000002-0000-0ff1-ce00-000000000000"
$p = Get-MsolServicePrincipal -ServicePrincipalName $ServiceName
New-MsolServicePrincipalCredential -AppPrincipalId $p.AppPrincipalId -Type asymmetric -Usage Verify -Value $credValue