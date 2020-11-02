# Syntax:	Get-TenantStatus.ps1 -name "microsoft"
#
# Determines if the tenant name requested is available or already taken within Office 365.
# Provide a desired tenant name without the .onmicrosoft.com suffix to check if that tenant name is available.

param( $name )

Function Usage
{
$strScriptFileName = ($MyInvocation.ScriptName).substring(($MyInvocation.ScriptName).lastindexofany("\") + 1).ToString()

@"

NAME:
$strScriptFileName

EXAMPLE:
C:\PS> .\$strScriptFileName -name `"microsoft`"

"@
}

If (-not $name) {Usage;Exit}

$uri = "https://portal.office.com/Signup/CheckDomainAvailability.ajax"
$body = "p0=" + $name + "&assembly=BOX.Admin.UI%2C+Version%3D16.0.0.0%2C+Culture%3Dneutral%2C+PublicKeyToken%3Dnull&class=Microsoft.Online.BOX.Signup.UI.SignupServerCalls"

$response = Invoke-RestMethod -Method Post -Uri $uri -Body $body

$valid = $response.Contains("SessionValid")
if ($valid -eq $false)
{
	Write-Host -ForegroundColor Red $response
	Exit
}

$available = $response.Contains("<![CDATA[1]]>")
if ($available) {
	Write-Host -ForegroundColor Green "Available"
} else {
	Write-Host -ForegroundColor Yellow "Taken"
}