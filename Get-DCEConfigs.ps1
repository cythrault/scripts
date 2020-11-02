<#
.NOTES
	Author:	Martin Thomas
#>

param(
$bladesOA = $true,
$bladesVC = $true,
$fabric = $true,
$email = $false,
$savetoSP = $true,
$AESKeyFilePath = "$PSScriptRoot\AESKey.key"
)

If (![System.IO.File]::Exists($AESKeyFilePath)) {
	$AESKey = New-Object Byte[] 32
	[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
	Set-Content $AESKeyFilePath $AESKey -Verbose
} Else { $AESKey = Get-Content $AESKeyFilePath }

function Get-DevicePassword([string]$CredFilePath,[string]$systemtype){
	if(![System.IO.File]::Exists($CredFilePath)){
		Set-Content $CredFilePath $($(Get-Credential -Username "admin" -Message "Enter credentials for $($systemtype)").Password | ConvertFrom-SecureString -Key $AESKey) -Verbose
	}
	return $(Get-Content $CredFilePath | ConvertTo-SecureString -Key $AESKey)
}

[console]::ForegroundColor = "Green"
[console]::BackgroundColor = "Black"

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

if ( $savetoSP -ne $false ) {
    $SPDir="https://sharepoint2013.wspgroup.com/sites/20IiWwB1Qpe4s8/Architecture/Configuration/"
    $webclient = New-Object System.Net.WebClient 
    $webclient.Credentials = $(Get-Credential -Username ([string]([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").UserPrincipalName) -Message "Enter your Sharepoint Credentials")
}

if ( ( $bladesOA -ne $false ) -or ( $bladesVC -ne $false) ) {
    $bladeuser="Administrator"
    $bladecreds = New-Object System.Management.Automation.PSCredential -ArgumentList $bladeuser, $(Get-DevicePassword "$PSScriptRoot\blades.cred" "Blades and Virtual Connects")
    }

if ( $bladesOA -ne $false ) {
$equipment = "cacol2bld0101", "cacol2bld0301", "cacol4bld0301", "cacol4bld0401" | ForEach { 
    [pscustomobject]@{
        HostName = $_
		EquipmentType = "Blade"
        Credentials = $bladecreds
        Command = "show all"
        SPFolder = "HPE - Compute"
        FileSuffix = "oa-showall"
	    }	
    }
}

if ( $bladesVC -ne $false ) {
$equipment += "cacol2swi0101", "cacol2swi0301", "cacol4swi0301", "cacol4swi0401" | ForEach { 
    [pscustomobject]@{
        HostName = $_
		EquipmentType = "vSwitch"
        Credentials = $bladecreds
        Command = "show all *"
        SPFolder = "HPE - Compute"
        FileSuffix = "vc-showall"
	    }	
    }
}

if ( $fabric -ne $false ) {
$brocadeuser="admin"
$brocadecreds = New-Object System.Management.Automation.PSCredential -ArgumentList $brocadeuser, $(Get-DevicePassword "$PSScriptRoot\brocades.cred" "Brocades")
$equipment += "cacol2fab0201", "cacol2fab0202", "cacol4fab0201", "cacol4fab0202" | ForEach { 
	[pscustomobject]@{
		HostName = $_
		EquipmentType = "Brocade"
		Credentials = $brocadecreds
		Command = "supportshow"
		SPFolder = "HPE - Storage"
		FileSuffix = "supportshow"
		}	
	}
}
$cpt=0
if ($equipment) { $equipment | ForEach {
	Write-Progress -Activity "Downloading config from $($equipment.count) devices: $($equipment.HostName)" -Status "in progress" -PercentComplete ($cpt++ / $equipment.count * 100) -CurrentOperation ("Downloading config from $($_.HostName) ($cpt/$($equipment.count))")
	Write-Host -NoNewLine "Connecting to $($_.HostName) to run $($_.Command)... "
    $sshsession = $(New-SSHSession -ComputerName $_.HostName -Credential $_.Credentials -OperationTimeout 999 -AcceptKey)
    $logname = "$currentdate-$($_.HostName)-$($_.FileSuffix).log"
	$(Invoke-SSHCommand -Index $sshsession.SessionId -Command $_.Command -TimeOut 999).Output | Out-File -Encoding UTF8 -FilePath $logname
	Remove-SSHSession -SessionId $sshsession.SessionId
    if ( $savetoSP -ne $false ) {
        Write-Host "Copying $logname to Sharepoint: $SPDir$($_.SPFolder)/"
        $file = Get-ChildItem $logname
        $webclient.UploadFile($SPDir + $($_.SPFolder) + "/" + $file.Name, "PUT", $file.FullName)
        }
	}

    Write-Host "Packaging logs into an archive: $currentdate-showall.zip and deleting them."
    Compress-Archive -Verbose -Force -Path "$currentdate-*.log" -DestinationPath "$currentdate-showall.zip"
    Remove-Item -Verbose -Path "$currentdate-*.log"
}

if ( $email -ne $false ) {
	$o = New-Object -com Outlook.Application
	$mail = $o.CreateItem(0)
	#$mail.importance = 2
	$mail.Subject = “Monthly OA/VC show all“
	$mail.Body = “Attached is a compressed archive which include the configuration of the HP Virtual Connect, Onboard Administrator and Brocade at the Montreal and AirDrie datacenters as of today. They are also available under $SPDir for archival purposes.“
	$mail.To = “rene.noel@hpe.com;martin.thomas@wsp.com;Mathieu.Charbonneau@wsp.com;Anthony.Daniel@wsp.com"
	$mail.Attachments.Add((Get-ChildItem $currentdate-showall.zip).FullName)
	$mail.Send()
	$o.Quit()
}

Read-Host -Prompt "Press Enter to exit"