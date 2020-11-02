[cmdletbinding()]
param()

$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
$obj = @()
foreach ($Site in $Sites) {

 $obj += New-Object -Type PSObject -Property (
  @{
   "SiteName"  = $site.Name
   "SubNets" = $site.Subnets -Join ";"
   "Domains" = $site.Domains -Join ";"
   "Location" = $site.Location -Join ";"
   "Servers" = $Site.Servers -Join ";"
  }
 )
}
$obj | Export-Csv 'sites.csv' -NoType