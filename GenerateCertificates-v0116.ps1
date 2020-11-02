[Cmdletbinding()]
Param(
	[Parameter(Mandatory=$True)][String]$Computer = $(throw "-Computer is required."),
	[String]$Password = 'WSPgroup2016',
	[String]$WorkDir = 'c:\ops\requests',
	[String]$CAConfig = 'ennycica01.corp.pbwan.net\Parsons Brinckerhoff Production Issuing CA 1 R1',
	[String]$DNSSuffix = 'corp.pbwan.net',
	[String]$Template = 'PBWebServer',
	[String]$Town = 'Montreal',
	[String]$Province = 'Quebec',
	[String]$PFXOutFile = '$Computer.pfx',
	[String]$CertRequestINF = "$WorkDir\$Computer.inf"
)

if (-NOT([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Administrator priviliges are required. Please restart this script with elevated rights." -ForegroundColor Red
    Throw "Administrator priviliges are required. Please restart this script with elevated rights."
}

Write-Host "Work Directory: $WorkDir"
Write-Host "Certificate CA: $CAConfig"
Write-Host "Certificate Template: $Template"
Write-Host "Output File: $PFXOutFile"
Write-Host "PFX Password: $Password"
	
if(-Not(Test-Path -Path $WorkDir -PathType Container)) {New-Item -Path $WorkDir -ItemType Directory -ErrorAction Stop}
if(-Not(Test-Path -Path \\$Computer.$DNSSuffix\c$\ops -PathType Container)){New-Item -Path \\$Computer.$DNSSuffix\c$\ops -ItemType Directory -ErrorAction Stop}

Write-Host "`nCreating $WorkDir\$Computer.inf using this content:"
Remove-Item $WorkDir\$Computer.inf -ErrorAction SilentlyContinue

@"
[Version]
Signature="$Windows NT$"

[NewRequest]
Subject = "CN=$Computer.$DNSSuffix,OU=Global - IT Group Operations,O=WSP Group Limited,L=$Town,S=$Province,C=CA,E=GLOBAL-ITGroupOperationsTeam@WSPGroup.com"
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
FriendlyName = $Computer
KeySpec = 1
KeyUsage = CERT_KEY_ENCIPHERMENT_KEY_USAGE

[RequestAttributes] 
CertificateTemplate = $Template
	
[EnhancedKeyUsageExtension]
OID = 1.3.6.1.5.5.7.3.1
"@ | Out-File -FilePath "$WorkDir\$Computer.inf"

Get-Content $WorkDir\$Computer.inf

Write-Host "`nWriting request to $WorkDir\$Computer.req"
Remove-Item $WorkDir\$Computer.req -ErrorAction SilentlyContinue
(& certreq -new $CertRequestINF "$WorkDir\$Computer.req")

$title = "Confirm Certificate Creation"
$message = "Proceed with certificate request for $Computer.$DNSSuffix?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
if ($result -gt 0) {exit}

Write-Host "`nSubmitting request for $Computer to $CAConfig using $WorkDir\$Computer.req and writing to $WorkDir\$Computer.cer"
Remove-Item $WorkDir\$Computer.cer -ErrorAction SilentlyContinue
(& certreq -submit -config "$CAConfig" "$WorkDir\$Computer.req" "$WorkDir\$Computer.cer")
Write-Host "Accepting request for $WorkDir\$Computer.cer"
(& certreq -accept "$WorkDir\$Computer.cer")

Write-Host "Exporting certificate to $WorkDir\$PFXOutFile with private key."
(& certutil -p $Password -exportPFX "$Computer.$DNSSuffix" "$WorkDir\$PFXOutFile")

$title = "Copy certificate to remote host"
$message = "Should we copy the certificate for $Computer.$DNSSuffix to \\$Computer.$DNSSuffix\c$\ops?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
if ($result -eq 0) {
	Write-Host "`nCopying $WorkDir\$PFXOutFile to \\$Computer.$DNSSuffix\c$\ops"
	Copy-Item -Path $WorkDir\$PFXOutFile -Destination \\$Computer.$DNSSuffix\c$\ops
}

$title = "Import certificate to remote host computer store"
$message = "Should import the certificate for $Computer.$DNSSuffix to it's computer store?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
if ($result -eq 0) {
	Write-Host "`nImporting $WorkDir\$PFXOutFile to $Computer.$DNSSuffix computer store."
	Copy-Item -Path $WorkDir\$PFXOutFile -Destination \\$Computer.$DNSSuffix\c$\ops
	$CertPath = "c:\ops\$PFXOutFile"
	Invoke-Command -ComputerName $Computer -ScriptBlock {
		param(
			[Parameter(Mandatory=$true)][String]$Password,
			[Parameter(Mandatory=$true)][String]$CertPath
			)
		[String]$certRootStore = "localmachine"
		[String]$certStore = "My"
		$pfxPass = ConvertTo-SecureString $Password -AsPlainText -Force
		$pfx = new-object System.Security.Cryptography.X509Certificates.X509Certificate2 
		$pfx.import($CertPath,$pfxPass,"Exportable,PersistKeySet") 
		$store = new-object System.Security.Cryptography.X509Certificates.X509Store($certStore,$certRootStore) 
		$store.open("MaxAllowed") 
		$store.add($pfx) 
		$store.close() 
	} -ArgumentList $Password, $CertPath
}

$title = "Remove Certificate from local computer store"
$message = "Should we delete the certificate for $Computer.$DNSSuffix in the local computer store?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 1) 
if ($result -eq 0) {
	Write-Host "Deleting certificate from local computer store."
	(& certutil –privatekey –delstore 'MY' "$Computer.$DNSSuffix")
}

Write-Host "`nYou should consider archiving or deleting the work files $Computer.* under $WorkDir."