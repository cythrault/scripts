<#	
	.NOTES
	===========================================================================
	Created on:   	2017/11/03 
	Created by:   	Drago Petrovic | Dominic Manning
	Organization: 	MSB365.blog
	Filename:     	Enable-UMM.ps1	

	Find us on:
		* Website:  https://www.msb365.blog
		* Technet:	https://social.technet.microsoft.com/Profile/MSB365
		* LinkedIn:	https://www.linkedin.com/in/drago-petrovic/
		* Xing:     https://www.xing.com/profile/Drago_Petrovic
	===========================================================================
	.DESCRIPTION
		The Script connect to the exchange online and AzureAD or your on-premises Environment. It collects the Mailbox information (UserPrincipalName) from the exchange online or
		on-premise Environment and the LineURI form AzureAD. The phonenumber will be modified to be usable as Extention to enable the UM Role on Exchange for all Users.
		The Script Lists all existing UM-Policies and shows it a chooseble options during the Script.
	
	.NOTES
		Requires PowerShell 5.0, or the AzureAD PowerShell module (both  for Office365/Exchange online)

	.COPYRIGHT
	Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
	to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
	and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
	WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
	===========================================================================
	.CHANGE LOG
		V1.00, 2017/11/03 - Initial version - (Enable UM in exchange online)
		V2.00, 2017/11/14 - Added on-premise as a options | choose between Office365 and on-premise environment after starting Script to enable UM
		V2.10, 2017/11/18 - Added the option to choose the Extension number length during the script execution, Added AzureAD module check (if O365 is selected)
		V2.50, 2017/11/18 - Added support for MSonline module, added comments, added error count for summary
		V2.52, 2017/12/04 - Added selection menu for recipient types, added check if assigned plans for user includes services 'MicrosoftCommunicationsOnline' and 'exchange'  


--- keep it simple, but significant ---
#>

<#
TODOs:
- Check if extension number is used by multiple users (AD / LDAP query)
- Disable UM
- Prevent time-outs in O365 connection?
#>


#region log file
$date = (Get-Date -Format yyyyMMdd_HHmm) #create time stamp
$log = "$PSScriptRoot\$date-EnableUM.Log" #define path and name, incl. time stamp for log file
#endregion

Write-Host '--- keep it simple, but significant ---' -ForegroundColor magenta

#region environment selection, modules and credentials
#show options for environment selection
$ExcOpt = Read-Host "Choose environment to connect to. 
[1] O365 
[2] On-Premises
Your option"

#get credentials according to the selected environment
switch ($ExcOpt)
{
	1  {
		#If 1 is selected
		
			try
		{
			# Check if AzureAD module is available and import it. 
			if (!(Get-Module AzureAD) -or !(Get-Module AzureADPreview))
			{
				Import-Module AzureAD -ErrorAction Stop
				$AADModule = 'AAD'
				(Get-Date -Format G) + " " + "Azure AD module loaded" | Tee-Object -FilePath $log -Append
				
			}
			
		}
		
		catch
		{
			
			cls
			try
			{
					#Try to load MSonline module, if Azure AD module is not available
					Import-Module MSOnline -ErrorAction Stop
					$AADModule = 'MSOnline'
					(Get-Date -Format G) + " " + "MSOnline module loaded" | Tee-Object -FilePath $log -Append
				}
			catch
			{
				#If no module is available, show option to open download page for MSonline module
					Write-Host "For O365 environments you first need to install MSOnline, AzureAD, or AzureADPreview module!" -ForegroundColor Red
					Write-Host "Please install one of the modules and restart the script." -ForegroundColor Cyan
					""
					
					$red = Read-Host "Do you want to be redirected to the MS download page for the MSOnline module? [Y] Yes, [N] No. Default is No."
					switch ($red)
						{
								Y { [system.Diagnostics.Process]::Start('http://connect.microsoft.com/site1164/Downloads/DownloadDetails.aspx?DownloadID=59185')}
								N {"Script will end now."}
							default { "Script will end now."}
						}
					return
				}
			
			
			
		}
		#Ask for O365 credentials
		"O365 selected"; $O365Creds = Get-Credential -Message 'Enter your O365 credentials'
		
	}
	2  {
		#If 2 is selected, ask for On-Prem credentials
		"On-Prem selected"; $OnPremCreds = Get-Credential -Message 'Enter your Exchange On-Prem credentials'
	} 
	default { Write-Host "Please enter 1, or 2" -ForegroundColor Red; return}
}
#endregion

#region select recipient type

$patwrong = $false
$YN = $null
do
{
	do
	{ #show options for recipient type detail selection
		$RecType = Read-Host "Please select the recipient type(s) you want to include. Separate multiple values by comma (1,2,...).

        [1] User Mailbox 
        [2] Shared Mailbox
        [3] Room Mailbox
        [4] Team Mailbox
        [5] Group Mailbox
        [C] Cancel

        Your selection"
		""
		
		if ($RecType -eq 'c')
			{
				"Exiting..."
				return
			}
		#Verify entered value
		$pattern = '^(?!.*?([1-5]).*?\1)[1-5](?:,[1-5])*$'
		
		if ($RecType.Length -gt 9 -or $RecType -notmatch $pattern)
			{
				'Incorrect format!'
				sleep -Seconds 1
				$patwrong = $true
				
				#return
			}
		else
			{
				$patwrong = $false
			}
	}
	until ($patwrong -eq $false)
	#Create string for get-mailbox -recipienttypedetails parameter according to user selection
	$RecType = $RecType.Replace('1', 'UserMailbox').Replace('2', 'SharedMailbox').Replace('3', 'RoomMailbox').Replace('4', 'TeamMailbox').Replace('5', 'GroupMailbox')
	#Ask if selected types are correct
	"Following recipient type(s) will be included:"
	""
	$($RecType -split ',')
	""
	
	$YN = read-host "Correct? [Y/N]"
}
until ($yn -eq 'y')

#endregion

#region set extension length
cls
#Ask for length of extension number
$ExtLen = Read-Host "Please enter the length of the extension number in your environment for UM"
#Check if a valid digit was entered
if ($ExtLen -notmatch "\d" -or $ExtLen -eq 0)
{
	Write-Host "Unsupported format. Only digits greater then 0 are supported." -ForegroundColor Red
	return
}
#endregion

#region O365 | Connect
if ($ExcOpt -eq 1)
{
	#Connect to Exchange Online remotely
	try
	{
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop
		Write-Host "Connecting to Exchange Online..." -ForegroundColor Green
		Import-PSSession $Session -ErrorAction Stop | Out-Null
		(Get-Date -Format G) + " " + "Exchange Online connected" | Tee-Object -FilePath $log -Append
	}
	catch
	{
		(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
		return
	}
	#Connect to AzureAD
	try
	{
		Write-Host "Connecting to Azure AD..." -ForegroundColor Green
		
		if ($AADModule -eq 'AAD')
		{
			#Use AzureAD module
			$aad = Connect-AzureAD -Credential $O365Creds -ErrorAction Stop
			
		}
		else
		{
			#Use MSonline module
			connect-MsolService -credential $O365Creds -ErrorAction Stop
			
		}
		
		(Get-Date -Format G) + " " + "Azure AD $($aad.TenantDomain) connected" | Tee-Object -FilePath $log -Append
		
	}
	catch
	{
		(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
		return
	}
	
}
#endregion

#region On-Premises | connect to Exchange
if ($ExcOpt -eq 2)
{
	#Ask for Exchange server name
	$Exchange = Read-Host "Enter FQDN, or short name of on-premises Exchange server. E.g. ""EXCSRV01.contoso.com, or EXCSRV01"
		
	try
	{
		#Remote connect to Exchange On-Prem 
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchange/PowerShell/ -Authentication Kerberos -Credential $OnPremCreds
		Import-PSSession $Session -ErrorAction Stop | Out-Null
		(Get-Date -Format G) + " " + "Exchange connected" | Tee-Object -FilePath $log -Append
	}
	catch
	{
		(Get-Date -Format G) + " " + "ERROR: " + $_.exception.message | Tee-Object -FilePath $log -Append
		Return
	}
	
}
#endregion

#region Select UMPolicy
cls
#Get all UM policies
$UMPolicies = Get-UMMailboxPolicy

#Count and list UM policies 
$count = 0
foreach ($policy in $UMPolicies)
	{
		$count++
		Write-Host "[$count] - $($policy.Name)" -ForegroundColor Cyan
		
	}
#Ask for UM policy to choose. (Enter number)
[INT]$Idx = Read-Host "Enter number of UM policy to choose"

#Check if entered number is valid
if ($Idx -eq 0 -or $Idx -gt $count)
	{
		#If entered number is not valid, end script
		cls
		Write-Host "Please select a number between 1 and $count. Script ends now." -ForegroundColor Red
		return
	}
else
	{
		#Select UM policy on base of entered number
		$UMPolicy = $UMPolicies[$idx - 1].Name
		cls
		Write-Host "You have selected the following policy: $UMPolicy" -BackgroundColor Blue
	}
#endregion

#region Enable UM users
Write-Host "Fetching mailboxes..." -ForegroundColor Green
#Get all mailboxes where UM is not enabled
$mbxs = get-mailbox -RecipientTypeDetails $RecType -ResultSize unlimited | where {$_.UMEnabled -eq $false}
$mcount = 0
$successcount = 0
$errorcount = 0

#Go through all found mail boxes
foreach ($mbx in $mbxs)
{
	#Create progress bar
	$mcount++
	$percent = "{0:N1}" -f ($mcount / $mbxs.count * 100)
	Write-Progress -Activity "Enabling UM" -status "Enabling Service for $($mbx.PrimarySMTPAddress)" -percentComplete $percent -CurrentOperation "Percent completed: $percent% (no. $mcount) of $($mbxs.count) mailboxes"
	
	#Get phone number of user
	try
	{
		switch ($ExcOpt )
		{
			1 {
				#If O365 selected, get phone number from Azure AD 
				if ($AADModule -eq 'AAD')
				{
					#Use AzureAD module
					$aadUser = get-azureADUser -SearchString $mbx.UserPrincipalName -erroraction Stop
					$phone = $aadUser.TelephoneNumber
					#Throw an error if no phone number was found for the user
					if ($phone -eq "" -or $phone -eq $null)
					{
						throw "$($mbx.UserPrincipalName) - No phone number found"
						return
					}
					if ($aadUser.AssignedPlans.service -notcontains 'MicrosoftCommunicationsOnline')
					{
						throw "Error: $($mbx.userprincipalname) has no S4B Online (Plan 2) plan assigned."
						return
					}
					if ($aadUser.AssignedPlans.service -notcontains 'exchange')
					{
						throw "Error: $($mbx.userprincipalname) has no Exchange Online (E1, or E2) plan assigned."
						return
					}
				}
				else
				{
					#Use MSOnline module
					$aadUser = get-MsolUser -SearchString $mbx.UserPrincipalName -erroraction Stop
					$phone = $aadUser.PhoneNumber
					#Throw an error if no phone number was found for the user
					if ($phone -eq "" -or $phone -eq $null)
					{
						throw "$($mbx.UserPrincipalName) - No phone number found"
						return
					}
				}
			}
			2 {
				#If On-Prem is selected, use ADSI searcher to get the phone number
				$b = [adsisearcher]::new("userprincipalname=$($mbx.UserPrincipalName)")
				$result = $b.FindOne()
				$phone = $result.Properties.telephonenumber
				#Throw an error if no phone number was found for the user
				if ($phone -eq "" -or $phone -eq $null)
				{
					throw "$($mbx.UserPrincipalName) - No phone number found"
					return
				}
				
			}
		}
		
		
		#LineURI string modifiy for extension number (get only the last digits that were defined in the beginning)
		$str = $phone.TrimStart("tel:+").replace(" ","") #Trim all spaces
		$length = $str.Length #Get length of the string
		$URI =$str.Substring(($length - $ExtLen)) #Select only substring starting from string length minus defined length
		
		#Create extension mapping (maybe used for future versions)
		$ExtensionMap = @{
			User = $mbx.PrimarySMTPAddress
			Extension = $URI
			}
		#Enable UM for the mailbox
		Enable-UMMailbox -Identity $mbx.PrimarySMTPAddress -UMMailboxPolicy $UMPolicy -SIPResourceIdentifier $mbx.PrimarySMTPAddress`
						 -Extensions $ExtensionMap.Extension -PinExpired $false -ErrorAction Stop #-WhatIf
		
		#Log
		$datetime = (Get-Date -Format G)
		"$datetime SUCCESS: $($mbx.UserPrincipalName) has been enabled for UM" | Tee-Object $log -Append
		
		#Count successfully enabled mailboxes
		$successcount++
		
	}
	catch
	{
		#Log error
		$datetime = (Get-Date -Format G)
		"$datetime ERROR: $($mbx.UserPrincipalName)  $($_.Exception.Message)" | Tee-Object $log -Append
		#Count errors
		$errorcount++
	}
	
	
}
#End progress bar
Write-Progress -Activity "Enabling UM" -Completed

#endregion

#region show summary
#Number of successes
Write-Host "$successcount of $($mbxs.count) mailboxes have been successfully enabled for UM! " -ForegroundColor Green
#If errors occurred show number of errors
if ($errorcount -gt 0)
	{
		Write-Host "Number of errors during execution: $errorcount. Please check the log ""$log"" for details." -ForegroundColor Green
}
"Press any key to exit"
cmd /c pause | Out-Null
#endregion