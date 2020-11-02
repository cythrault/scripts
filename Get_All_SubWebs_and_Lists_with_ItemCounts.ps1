#environment variables 
$username = "username"
$password = "password"
$url = "siteURL"

$securePassword = ConvertTo-SecureString $password -AsPlainText -Force 

#add SharePoint Online DLL - update the location if required
$programFiles = [environment]::getfolderpath("programfiles")
add-type -Path $programFiles'\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll'  

# connect/authenticate to SharePoint Online and get ClientContext object.. 
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword) 
$ctx.Credentials = $credentials 

#get all the sub webs
$Web = $ctx.Web  
$ctx.Load($web)  
$ctx.Load($web.Webs)    
$ctx.executeQuery()

Write-Host -ForegroundColor Yellow "There are:" $web.Webs.Count "sub webs in this site collection"

#get all the lists 
foreach ($subweb in $web.Webs)
{
    $lists = $subweb.Lists
    $ctx.Load($lists)
    $ctx.ExecuteQuery()
    Write-Host -ForegroundColor Yellow "The site URL is" $subweb.Url

    #output the list details
    Foreach ($list in $lists)
    {
        Write-Host "List title is: " $list.Title". This list has: " $list.ItemCount " items"
    }
}