# https://evotec.xyz/how-to-change-your-own-expired-password-when-you-cant-login-to-rdp/
function Set-PasswordRemotely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string] $UserName,
        [Parameter(Mandatory = $true)][string] $OldPassword,
        [Parameter(Mandatory = $true)][string] $NewPassword,
        [Parameter(Mandatory = $true)][alias('DC', 'Server', 'ComputerName')][string] $DomainController
    )
    $DllImport = @'
[DllImport("netapi32.dll", CharSet = CharSet.Unicode)]
public static extern bool NetUserChangePassword(string domain, string username, string oldpassword, string newpassword);
'@
    $NetApi32 = Add-Type -MemberDefinition $DllImport -Name 'NetApi32' -Namespace 'Win32' -PassThru
    if ($result = $NetApi32::NetUserChangePassword($DomainController, $UserName, $OldPassword, $NewPassword)) {
        Write-Output -InputObject 'Password change failed. Please try again.'
    } else {
        Write-Output -InputObject 'Password change succeeded.'
    }
}

function pwdreset {
	param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("SamAccountName")][string]$sam,
		[Alias("Password")][string]$pwd="P@ssw0rd!"
		)
	Begin { $sam = $sam.Trim() }
	Process {
		Unlock-ADAccount -Identity $sam -Verbose
		Set-ADUser -Identity $sam -ChangePasswordAtLogon:$true -Verbose
		Set-ADAccountPassword -Identity $sam -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $pwd -Force) -Verbose
	}
}

# credential save

function creds ([Parameter(Mandatory=$true)]$user) {
$AESKeyFilePath = "$Env:UserProfile\AESKey.key"
$credsFile = "$Env:UserProfile\creds.aes"

	If (![System.IO.File]::Exists($AESKeyFilePath)) {
		$AESKey = New-Object Byte[] 32
		[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
		Set-Content $AESKeyFilePath $AESKey -Verbose
	} Else { $AESKey = Get-Content $AESKeyFilePath }

	If (![System.IO.File]::Exists($credsFile)) {
		Set-Content $credsFile $($(Get-Credential -Username $user -Message "Message").Password | ConvertFrom-SecureString -Key $AESKey)
		}

	return (New-Object System.Management.Automation.PSCredential($user, (Get-Content $credsFile | ConvertTo-SecureString -Key $AESKey)))
}

# sharepoint stuff

function Invoke-LoadMethod() {
param(
   $ClientObject = $(throw "Please provide an Client Object instance on which to invoke the generic method")
) 
   $ctx = $ClientObject.Context
   $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
   $type = $ClientObject.GetType()
   $clientObjectLoad = $load.MakeGenericMethod($type) 
   $clientObjectLoad.Invoke($ctx,@($ClientObject,$null))
}

function PrintWebProperties()
{
 param(
        [Parameter(Mandatory=$true)][string]$url,
        [Parameter(Mandatory=$false)][System.Net.NetworkCredential]$credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials
    )
   $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($url)
   $ctx.Credentials = $credentials
   $web = $ctx.Web 
   Invoke-LoadMethod -ClientObject $web
   $ctx.ExecuteQuery()
   
   Write-Host "Web Properties:"
   Write-Host "Title: $($web.Title)"
   Write-Host "Url: $($web.ServerRelativeUrl)"
}

function processWeb($web)
{
    $lists = $web.Lists
    $ctx.Load($web)
    $ctx.ExecuteQuery()
    Write-Host "Web URL is" $web.Url
}

# GPU Functions: http://www.virtu-al.net/2015/10/26/adding-a-vgpu-for-a-vsphere-6-0-vm-via-powercli/

Function Get-GPUProfile {
    Param ($VMHost)
    $VMhost = Get-VMhost $VMhost
    $VMHost.ExtensionData.Config.SharedPassthruGpuTypes
}
  
Function Get-vGPUDevice {
    Param ($vm)
    $VM = Get-VM $VM
    $vGPUDevice = $VM.ExtensionData.Config.hardware.Device | Where { $_.backing.vgpu}
    $vGPUDevice | Select Key, ControllerKey, Unitnumber, @{Name="Device";Expression={$_.DeviceInfo.Label}}, @{Name="Summary";Expression={$_.DeviceInfo.Summary}}
}
  
Function Remove-vGPU {
    Param (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,Position=0)] $VM,
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,Position=1)] $vGPUDevice
    )
  
    $ControllerKey = $vGPUDevice.controllerKey
    $key = $vGPUDevice.Key
    $UnitNumber = $vGPUDevice.UnitNumber
    $device = $vGPUDevice.device
    $Summary = $vGPUDevice.Summary
  
    $VM = Get-VM $VM
  
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec[] (1)
    $spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
    $spec.deviceChange[0].operation = 'remove'
    $spec.deviceChange[0].device = New-Object VMware.Vim.VirtualPCIPassthrough
    $spec.deviceChange[0].device.controllerKey = $controllerkey
    $spec.deviceChange[0].device.unitNumber = $unitnumber
    $spec.deviceChange[0].device.deviceInfo = New-Object VMware.Vim.Description
    $spec.deviceChange[0].device.deviceInfo.summary = $summary
    $spec.deviceChange[0].device.deviceInfo.label = $device
    $spec.deviceChange[0].device.key = $key
    $_this = $VM  | Get-View
    $nulloutput = $_this.ReconfigVM_Task($spec)
}
  
Function New-vGPU {
    Param ($VM, $vGPUProfile)
    $VM = Get-VM $VM
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec[] (1)
    $spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
    $spec.deviceChange[0].operation = 'add'
    $spec.deviceChange[0].device = New-Object VMware.Vim.VirtualPCIPassthrough
    $spec.deviceChange[0].device.deviceInfo = New-Object VMware.Vim.Description
    $spec.deviceChange[0].device.deviceInfo.summary = ''
    $spec.deviceChange[0].device.deviceInfo.label = 'New PCI device'
    $spec.deviceChange[0].device.backing = New-Object VMware.Vim.VirtualPCIPassthroughVmiopBackingInfo
    $spec.deviceChange[0].device.backing.vgpu = "$vGPUProfile"
    $vmobj = $VM | Get-View
    $reconfig = $vmobj.ReconfigVM_Task($spec)
    if ($reconfig) {
        $ChangedVM = Get-VM $VM
        $vGPUDevice = $ChangedVM.ExtensionData.Config.hardware.Device | Where { $_.backing.vgpu}
        $vGPUDevice | Select Key, ControllerKey, Unitnumber, @{Name="Device";Expression={$_.DeviceInfo.Label}}, @{Name="Summary";Expression={$_.DeviceInfo.Summary}}
  
    }   
}

function Change-TitleBar() {
    if ($global:DefaultVIServers.Count -gt 0) {
        $strWindowTitle = "[PowerCLI] Connected to {0} server{1}:  {2}" -f $global:DefaultVIServers.Count, $(if ($global:DefaultVIServers.Count -gt 1) {"s"}), (($global:DefaultVIServers | %{$_.Name}) -Join ", ")
    }
    else {
        $strWindowTitle = "[PowerCLI] Not Connected"
    }
    $host.ui.RawUI.WindowTitle = $strWindowTitle
}

Function NewPools {

param(
	$VMHost = $(Get-VMHost $(menu "dccan100vmw0605.corp.pbwan.net","dccan100vmw0606.corp.pbwan.net")),
	[array]$VMs = $(menu -menuItems $(Get-VM | Select Name,VMHost | where { $_.VMHost.Name -eq $vmhost }).Name -Multiselect),
	[string]$vGPUprof = $(menu -menuItems $($($VMHost.ExtensionData.Config.SharedPassthruGpuTypes) + "None")),
	[array]$adusers = "corp\camt055545"
)

	while ( $aduser = Read-Host -Prompt 'sAMAccountName of users to add entitlement?' ) {
	try { 
		Get-ADUser -Identity $aduser
		$adusers += "corp\" + $aduser
		}
	catch { $_.Exception.Message }
	}

	Switch -regex ( $vGPUprof ) {
		'm10' { $PoolGPUType = "vGPU-M10" }
		'm60' { $PoolGPUType = "vGPU-M60" }
		default { $PoolGPUType = "noGPU" }
	}

	Write-Host "Creating pool(s) with the following VM(s) using the $vGPUprof GPU profile: $VMs for $adusers"
	if ( (menu -menuItems "Yes", "No") -ne "Yes" ) { throw "stopping processing." }
	
	ForEach ( $NewVM in $VMs ) {
		write-host "get vm for $NewVM"
		$vm = Get-VM -Name $NewVM
		$PoolName = "$PoolGPUType-$($vm.Name.Substring($vm.Name.Length - 3))"
		write-host "poweroff if on"
		if (($vm | select PowerState).PowerState -ne "PoweredOff") {
			write-host "powering off guest"
			Shutdown-VMGuest -VM $vm -Confirm:$false
			write-host -nonewline "wait for poweroff"
			do {sleep 9; write-host -nonewline "."} until ($(Get-VM -Name $vm).PowerState -eq "PoweredOff"); write-host "done"
			}
		write-host "remove gpu"
		if (Get-vGPUDevice -vm $vm) { Remove-vGPU -VM $vm -vGPUDevice (Get-vGPUDevice -vm $vm) }
		write-host "create pool"
		New-HVPool -Manual -PoolName $PoolName -UserAssignment FLOATING -Vcenter cathl100vce01 -VM $vm.Name -Source VIRTUAL_CENTER
		write-host "add gpu"
		New-vGPU -VM $vm -vGPUProfile $vGPUprof
		write-host "start vm"
		Start-VM -VM $vm
		write-host -nonewline "wait for availability"
		do {sleep 9; write-host -nonewline "."} until ( (Get-HVMachine -MachineName $vm.Name).Base.BasicState -eq 'AVAILABLE' ); write-host "done"
		write-host "add entitlement"
		$adusers | % { Get-HVPool -PoolName $PoolName | New-HVEntitlement -User $_ -Confirm:$false }
	}
}

Function ConnectVCHV {
	Connect-VIServer -Server cathl100vce01
	Import-Module VMware.VimAutomation.HorizonView
	Import-Module VMware.Hv.Helper
	$hvServer = Connect-HVServer -server view-ca.wsp.com -Credential $(creds corp\camt055545it9)
	$hvServices = $hvserver.ExtensionData
	$csList = $hvServices.ConnectionServer.ConnectionServer_List()
}

Function EnableCopyPaste {

$VMHost = $(Get-VMHost $(menu -menuItems "dccan100vmw0605.corp.pbwan.net","dccan100vmw0606.corp.pbwan.net"))
[array]$VMs = $(menu -menuItems $(Get-VM | Select Name,VMHost | where { $_.VMHost.Name -eq $vmhost }).Name -Multiselect)

$VMs | %{
	$vm
    New-AdvancedSetting $_ -Name isolation.tools.copy.disable -Value false -Confirm:$false -Force:$true
    New-AdvancedSetting $_ -Name isolation.tools.paste.disable -Value false -Confirm:$false -Force:$true
	}
}

function addIP {
	[CmdletBinding(DefaultParameterSetName='IP')]
	param (
		[Parameter(Mandatory = $true, ParameterSetName = 'IP')]
        [IPAddress]
        $IP,
        [Parameter(Mandatory = $true, ParameterSetName = 'HostName')]
        [String]
        $HostName
	)
	if ($HostName) {
		Write-Host "Resolving $HostName"
		[IPAddress]$IP = (Resolve-DnsName $HostName).IPAddress
	}
	$IPlist = (Get-ReceiveConnector |?{$_.Identity -like "*Anonymous Relay*"}).RemoteIPRanges | Sort-Object -Unique
	
	if ($IPlist -match $IP.IPAddressToString) {
		Write-Warning "$IP is already white-listed"
	} else {
		try {
			$NameHost = (Resolve-DnsName $IP -ErrorAction Stop).NameHost
			Write-Host "Adding $IP ($NameHost)"
		}
		catch {
			Write-Warning "Unable to resolve $IP"
			Write-Host "Adding $IP"
		}
		$IPlist += $IP.IPAddressToString
		$IPlist = $IPlist | Sort-Object -Unique
	}
	#Set-ReceiveConnector "PVWEXC0F\Anonymous Relay PVWEXC0F" -RemoteIPRanges 255.255.255.255
	Set-ReceiveConnector "PVWEXC0F\Anonymous Relay PVWEXC0F" -RemoteIPRanges $IPlist
	#Set-ReceiveConnector "PVWEXC0G\Anonymous Relay PVWEXC0G" -RemoteIPRanges 255.255.255.255
	Set-ReceiveConnector "PVWEXC0G\Anonymous Relay PVWEXC0G" -RemoteIPRanges $IPlist
}

function Get-LocalAdmins {
	param ([Parameter(Mandatory = $true)]$strcomputer)
	$admins = Gwmi win32_groupuser -computer $strcomputer
	$admins = $admins |? {$_.groupcomponent -like '*"Administrat*"'}
	$admins | %{
		$_.partcomponent -match ".+Domain\=(.+)\,Name\=(.+)$" | Out-Null
		$matches[1].trim('"') + "\" + $matches[2].trim('"')
	}
}


Function Get-CachedCredential 
 { 
    <# 
        .SYNOPSIS 
            Return a list of cached credentials 
        .DESCRIPTION 
            This function wraps cmdkey /list and returns an object that contains 
            the Targetname, user and type of cached credentials. 
        .PARAMETER Type 
            To filter the list provide one of the basic types of  
            cached credentials 
                Domain 
                Generic 
        .EXAMPLE 
            Get-CachedCredential 

            Target                         Type            User                  
            ------                         ----            ----                  
            Domain:target=server-02        Domain Password COMPANY\Administrator 
            Domain:target=server-01        Domain Password COMPANY\Administrator 
            LegacyGeneric:target=server-04 Generic         COMPANY\Administrator 
            LegacyGeneric:target=server-03 Generic         COMPANY\Administrator 

            Description 
            ----------- 
            This example shows using the syntax without passing a type parameter, which is 
            the same as passing -Type All 
        .EXAMPLE 
            Get-CachedCredential -Type Domain 

            Target                         Type            User                  
            ------                         ----            ----                  
            Domain:target=server-02        Domain Password COMPANY\Administrator 
            Domain:target=server-01        Domain Password COMPANY\Administrator 

            Description 
            ----------- 
            This example shows using type with one of the valid types available 
        .NOTES 
            FunctionName : Get-CachedCredential 
            Created by   : jspatton 
            Date Coded   : 06/23/2014 10:11:42 

            ** 
            This function does not return a cached credential that doesn't hold 
            a value for user 
            ** 
        .LINK 
            https://code.google.com/p/mod-posh/wiki/CachedCredentialManagement#Get-CachedCredential 
        .LINK 
            http://technet.microsoft.com/en-us/library/cc754243.aspx 
        .LINK 
            http://www.powershellmagazine.com/2014/04/18/automatic-remote-desktop-connection/ 
    #> 
    [CmdletBinding()] 
    Param 
        ( 
        [ValidateSet("Generic","Domain","Certificate","All")] 
        [string]$Type 
        ) 
    Begin 
    { 
        $Result = cmdkey /list 
        } 
    Process 
    { 
        $Return = @() 
        $Temp = New-Object -TypeName psobject 
        foreach ($Entry in $Result) 
        { 
            if ($Entry) 
            { 
                $Line = $Entry.Trim(); 
                if ($Line.Contains('Target: ')) 
                { 
                    Write-Verbose $Line 
                    $Target = $Line.Replace('Target: ',''); 
                    } 
                if ($Line.Contains('Type: ')) 
                { 
                    Write-Verbose $Line 
                    $TargetType = $Line.Replace('Type: ',''); 
                    } 
                if ($Line.Contains('User: ')) 
                { 
                    Write-Verbose $Line 
                    $User = $Line.Replace('User: ',''); 
                    Add-Member -InputObject $Temp -MemberType NoteProperty -Name Target -Value $Target 
                    Add-Member -InputObject $Temp -MemberType NoteProperty -Name Type -Value $TargetType 
                    Add-Member -InputObject $Temp -MemberType NoteProperty -Name User -Value $User 
                    $Return += $Temp; 
                    Write-Verbose $Temp; 
                    $Temp = New-Object -TypeName psobject 
                    } 
                } 
            } 
        } 
    End 
    { 
        if ($Type -eq "All" -or $Type -eq "") 
        { 
            Write-Verbose "ALL" 
            return $Return; 
            } 
        else 
        { 
            Write-Verbose "FILTERED" 
            if ($Type -eq "Domain") 
            { 
                $myType = "Domain Password" 
                } 
            if ($Type -eq "Certificate") 
            { 
                $myType = "Generic Certificate" 
                } 
            return $Return |Where-Object -Property Type -eq $myType 
            } 
        } 
    } 
Function Add-CachedCredential 
{ 
    <# 
        .SYNOPSIS 
            Add a cached credential to the vault 
        .DESCRIPTION 
            This function wraps cmdkey /add and stores a TargetName and 
            user/pass combination in the vault 
        .PARAMETER TargetName 
            The name of the object to store credentials for, typically 
            this would be a computer name 
        .PARAMETER Type 
            Add credentials in one of the few valid types of 
            cached credentials 
                Domain 
                Generic 
        .PARAMETER Credential 
            A PSCredential object used to securely store user and  
            password information 
        .EXAMPLE 
            Add-CachedCredential -TartName server-01 -Type Domain -Credential (Get-Credential) 

            CMDKEY: Credential added successfully. 

            Description 
            ----------- 
            The basic syntax of the command 
        .EXAMPLE 
            "server-04","server-05" |Add-CachedCredential -Type Domain -Credential $Credential 

            CMDKEY: Credential added successfully. 

            CMDKEY: Credential added successfully. 

            Description 
            ----------- 
            This example shows passing in Targetnames on the pipeline 
        .NOTES 
            FunctionName : Add-CachedCredential 
            Created by   : jspatton 
            Date Coded   : 06/23/2014 12:13:21 
        .LINK 
            https://code.google.com/p/mod-posh/wiki/CachedCredentialManagement#Add-CachedCredential 
        .LINK 
            http://technet.microsoft.com/en-us/library/cc754243.aspx 
        .LINK 
            http://www.powershellmagazine.com/2014/04/18/automatic-remote-desktop-connection/ 
    #> 
    [CmdletBinding()] 
    Param 
        ( 
        [Parameter(Mandatory=$true,ValueFromPipeline=$True)] 
        [string]$TargetName, 
        [ValidateSet("Generic","Domain")] 
        [string]$Type, 
        [Parameter(Mandatory=$true)] 
        [System.Management.Automation.PSCredential]$Credential 
        ) 
    Begin 
    { 
        $Username = $Credential.UserName; 
        $Password = $Credential.GetNetworkCredential().Password; 
        } 
    Process 
    { 
        foreach ($Target in $TargetName) 
        { 
            switch ($Type) 
            { 
                "Generic" 
                { 
                    $Result = cmdkey /generic:$Target /user:$Username /pass:$Password 
                    if ($LASTEXITCODE -eq 0) 
                    { 
                        Return $Result; 
                        } 
                    { 
                        Write-Error $Result 
                        Write-Error $LASTEXITCODE 
                        } 
                    } 
                "Domain" 
                { 
                    $Result = cmdkey /add:$Target /user:$Username /pass:$Password 
                    if ($LASTEXITCODE -eq 0) 
                    { 
                        Return $Result; 
                        } 
                    { 
                        Write-Error $Result 
                        Write-Error $LASTEXITCODE 
                        } 
                    } 
                } 
            } 
        } 
    End 
    { 
        } 
    } 
Function Remove-CachedCredential 
{ 
    <# 
        .SYNOPSIS 
            Remove a target from the vault 
        .DESCRIPTION 
            This function wraps cmdkey /delete to remove a specific 
            target from the vault 
        .PARAMETER TargetName 
            The target to remove 
        .EXAMPLE 
            Remove-CachedCredential -TargetName server-04 

            CMDKEY: Credential deleted successfully. 

            Description 
            ----------- 
            This example shows the only usage for this command 
        .NOTES 
            FunctionName : Remove-CachedCredential 
            Created by   : jspatton 
            Date Coded   : 06/23/2014 12:27:18 
        .LINK 
            https://code.google.com/p/mod-posh/wiki/CachedCredentialManagement#Remove-CachedCredential 
        .LINK 
            http://technet.microsoft.com/en-us/library/cc754243.aspx 
        .LINK 
            http://www.powershellmagazine.com/2014/04/18/automatic-remote-desktop-connection/ 
    #> 
    [CmdletBinding()] 
    Param 
        ( 
        [Parameter(Mandatory=$true)] 
        [string]$TargetName 
        ) 
    Begin 
    { 
        } 
    Process 
    { 
        $Result = cmdkey /delete:$TargetName 
        } 
    End 
    { 
        if ($LASTEXITCODE -eq 0) 
        { 
            Return $Result; 
            } 
        { 
            Write-Error $Result 
            Write-Error $LASTEXITCODE 
            } 
        } 

    }
	
function monitor-moverequest {
param([string]$Identity = "")
	$Stats = Get-MoveRequestStatistics -Identity $Identity
	do {
		$Stats = Get-MoveRequestStatistics -Identity $Identity
		Write-Progress -Activity "Moving $Identity" -Status ([string]$Stats.BytesTransferred+" transferred of "+[string]$Stats.TotalMailboxSize+" - "+[int]$Stats.PercentComplete+"%"+" - "+[string]$Stats.StatusDetail) -PercentComplete $Stats.PercentComplete
		Start-Sleep -Seconds 5
	} while (($Stats.PercentComplete -le 99) -and ($Stats.Status.Value -eq "InProgress"))

	do {
		$Stats = Get-MoveRequestStatistics -Identity $Identity
		Write-Progress -Activity "Moving $Identity" -Status $Stats.Status -PercentComplete 100
		Start-Sleep -Seconds 5
	} while ($Stats.StatusDetail -ne "Completed")

	$Stats.StatusDetail
}

function move-mailbox {
	param([string]$Identity = "")
	if (get-mailbox -identity $identity) {
	
	
	
	}
}

function seeforest {
	Set-ADServerSettings -ViewEntireForest:$true
}

function Copy-ADGroupMember {
[CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact='High')]
Param(
	[Parameter(ValueFromPipeline=$True, Position=0)]
	[PSObject]$Identity,
	[PSObject]$Destination,
	[Switch]$Mirror
)
	$SourceMembers = Get-ADGroupMember -Identity $Identity | Select-Object -ExpandProperty distinguishedName
	$TargetMembers = Get-ADGroupMember -Identity $Destination | Select-Object -ExpandProperty distinguishedName
	$DifferenceMembers = Compare-Object -ReferenceObject $SourceMembers -DifferenceObject $TargetMembers
	$ProcessMembers = $DifferenceMembers | Where-Object {$_.SideIndicator -eq '<='} | Select-Object -ExpandProperty InputObject
	If ($PSCmdlet.ShouldProcess($Destination, "Add objects:`r`n`t$($ProcessMembers -join "`r`n`t")`r`n")) {
		Add-ADGroupMember -Identity $Destination -Members $ProcessMembers
	}
	If ($Mirror) {
		$ProcessMembers = $DifferenceMembers | Where-Object {$_.SideIndicator -eq '=>'} | Select-Object -ExpandProperty InputObject
		If ($PSCmdlet.ShouldProcess($Destination, "Remove objects:`r`n`t$($ProcessMembers -join "`r`n`t")`r`n")) {
			Remove-ADGroupMember -Identity $Destination -Members $ProcessMembers
		}
	}
}