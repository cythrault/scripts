<#
.NOTES
	Author:	Martin Thomas
#>

param(
	[Parameter(Mandatory=$true)]
	[ValidateSet('pbenas0a','phenas0a','phenas0b')]
	[string] $nas,
	$AESKeyFilePath = "$PSScriptRoot\AESKey.key"
)

If (![System.IO.File]::Exists($AESKeyFilePath)) {
	$AESKey = New-Object Byte[] 32
	[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
	Set-Content $AESKeyFilePath $AESKey -Verbose
} Else { $AESKey = Get-Content $AESKeyFilePath }

function Get-DevicePassword([string]$CredFilePath,[string]$systemtype){
	if(![System.IO.File]::Exists($CredFilePath)){
		Set-Content $CredFilePath $($(Get-Credential -Username "root" -Message "Enter credentials for Isilon").Password | ConvertFrom-SecureString -Key $AESKey) -Verbose
	}
	return $(Get-Content $CredFilePath | ConvertTo-SecureString -Key $AESKey)
}

$currentdate=$(get-date -uformat %Y%m%d)

if (Get-Module -ListAvailable -Name Posh-SSH) {
    Import-Module -Name Posh-SSH
} else {
    if (-NOT([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Throw "Please install Posh-SSH manually as administrator or run this script with administrative privileges."
    }
    Install-Module -Name Posh-SSH;
    Import-Module -Name Posh-SSH
}

$user = "root"
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $user, $(Get-DevicePassword "$PSScriptRoot\isilon.cred" "Isilon Credentials.")
$subnet = "10.9.83."
$command = "isi_audit_viewer -t protocol"

$ips = @{
	'pbenas0a' = 11..59
	'phenas0a' = 111..159
	'phenas0b' = 161..209
}

$cpt=0
$ips.$nas | ForEach {
	$ip = $subnet + $_
	if ( Test-NetConnection -ComputerName $ip -InformationLevel Quiet -CommonTCPPort SMB ) {
		$sshsession = $(New-SSHSession -ComputerName $ip -Credential $creds -OperationTimeout 999 -AcceptKey)
		$hostname = $(Invoke-SSHCommand -Index $sshsession.SessionId -Command "hostname" -TimeOut 999).Output
		$logname = "$currentdate-$($hostname[0].ToLower()).log"
		#Write-Progress -Activity "Downloading audit logs $($hostname)" -Status "in progress" -PercentComplete ($cpt++ / $equipment.count * 100) -CurrentOperation ("Downloading audit logs from $($hostname) : $($cpt/$($ips.$nas.Count))")
		Write-Host -NoNewline "Connecting to $($hostname) ($($ip)) to run $($command)... "
		$out = $(Invoke-SSHCommand -Index $sshsession.SessionId -Command $command -TimeOut 999).Output
		#$out = $out[0..($out.count - 2)]
		#$out = $out -ne "done"
		if ( $out ) { Out-File -Encoding UTF8 -FilePath $logname -InputObject $out }
		Remove-SSHSession -SessionId $sshsession.SessionId
	}
}