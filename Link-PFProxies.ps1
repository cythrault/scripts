# Link-PFProxies.ps1 
# 
# Change these two values to match your environment. 
# $container should point to the MESO container you want to run against. 
# $pfTreeDN should contain the distinguishedName of the public folder hierarchy object.

$container = [ADSI]("LDAP://CN=Microsoft Exchange System Objects,DC=bnquebec,DC=ca") 
$pfTreeDN = "CN=Public Folders,CN=Folder Hierarchies,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=bnquebec1,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=bnquebec,DC=ca"

#################################################

$filter = "(!(homemdB=*))" 
$propertyList = @("distinguishedName") 
$scope = [System.DirectoryServices.SearchScope]::OneLevel

$finder = new-object System.DirectoryServices.DirectorySearcher($container, $filter, $propertyList, $scope) 
$finder.PageSize = 100 
$results = $finder.FindAll()

("Found " + $results.Count + " folder proxies with no homeMDB...") 
foreach ($result in $results) 
{ 
    ("Fixing object: " + $result.Path) 
    #$entry = $result.GetDirectoryEntry() 
    #$entry.Put("homeMDB", $pfTreeDN) 
    #$entry.SetInfo() 
}