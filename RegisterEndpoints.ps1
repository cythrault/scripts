$ServiceName = "00000002-0000-0ff1-ce00-000000000000";
$x = Get-MsolServicePrincipal -AppPrincipalId $ServiceName;
$x.ServicePrincipalnames.Add("https://courriel.banq.qc.ca/");
$x.ServicePrincipalnames.Add("https://autodiscover.banq.qc.ca/");
Set-MSOLServicePrincipal -AppPrincipalId $ServiceName -ServicePrincipalNames $x.ServicePrincipalNames;