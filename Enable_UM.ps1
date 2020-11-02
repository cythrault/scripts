$File = "C:\Scripts\Password_O365.txt"
$Key = (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32)
$AdminUser = "admin-user@company.onmicrosoft.com"
$Password = Get-Content $File | ConvertTo-SecureString -Key $Key
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminUser,$Password
 
$emailfile = "C:\Scripts\Password.txt"
$emailuser = "Email-User@domain.com"
$emailpassword = Get-Content $emailfile | ConvertTo-SecureString -Key $key
$EmailCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $emailuser,$emailpassword
$To = 'User1@domain.com','User2@domain.com'
$From = 'Email-User@domain.com'
 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session
 
$VoiceEnabledUsers = (get-csuser | where {($_.EnterpriseVoiceEnabled -eq "True") -and ($_.RegistrarPool.FriendlyName -eq "Lync-Server.domain.com")}).SamAccountName
 
foreach ($VoiceEnabledUser in $VoiceEnabledUsers){
   $UPNUser = (Get-ADUser -Identity $VoiceEnabledUser).UserPrincipalName
   $MailValue = (Get-ADUser -Identity $VoiceEnabledUser -Properties *).mail
 
   if ($MailValue -ne $null){
      $UMStatus = (Get-Mailbox $UPNUser).UMEnabled
 
      if ($UMStatus -ne "True"){
         $SIPAddress = (Get-CsUser -Identity $UPNUser).SipAddress
         $SIP = $SIPAddress.Substring(4)
         $Line = (Get-CsUser -Identity $UPNUser).LineURI
         $Extension = $Line.Substring(($Line.Length)-3)
         $Policy = Get-UMMailboxPolicy "SfB Policy"
         Enable-UMMailbox -Identity $UPNUser -UMMailboxPolicy $Policy.Name -SIPResourceIdentifier $SIP -Extensions $Extension -PinExpired $true
         $UMStatusNew = (Get-Mailbox $UPNUser).UMEnabled
         $NormalEmailTemp = @"
<tr>
   <td class="colorm">$UPNUser</td>
   <td>$UMStatus</td>
   <td>$UMStatusNew</td>
</tr>
"@
 
         $NormalEmailResult = $NormalEmailResult + "`r`n" + $NormalEmailTemp
 
         $NormalEmailUp = @"
<style>
body {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; color:#434242;}
TABLE {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TR {border-width: 1px;padding: 10px;border-style: solid;border-color: white; }
TD {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;padding: 10px;border-style: solid;border-color: white; background-color:#C3DDDB;}
.colorm {background-color:#58A09E; color:white;}
h3 {color:#BD3337 !important;}
</style>
 
<body>
<h3> Unified Messaging Activation</h3>
 
<p>Unified Messaging has been activated on the below users:</p>
 
<table>
<tr>
   <td class="colort">User</td>
   <td class="colort">UM Old</td>
   <td class="colort">UM New</td>
</tr>
 
"@
 
         $NormalEmailDown = @"
</table>
</body>
"@
 
         $NormalEmail = $NormalEmailUp + $NormalEmailResult + $NormalEmailDown
 
         send-mailmessage `
            -To $To `
            -Subject "UM Activation Report $(Get-Date -format dd/MM/yyyy)" `
            -Body $NormalEmail `
            -BodyAsHtml `
            -Priority high `
            -UseSsl `
            -Port 587 `
            -SmtpServer 'smtp.office365.com' `
            -From $From `
            -Credential $EmailCredentials
      }
   }
}
 
if ($error -ne $null){
   foreach ($value in $error){
      $ErrorEmailTemp = @"
<tr>
   <td class="colorm">$value</td>
</tr>
"@
 
      $ErrorEmailResult = $ErrorEmailResult + "`r`n" + $ErrorEmailTemp
   }
 
   $ErrorEmailUp = @"
<style>
body {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; color:#434242;}
TABLE {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TR {border-width: 1px;padding: 10px;border-style: solid;border-color: white; }
TD {font-family:Segoe, "Segoe UI", "DejaVu Sans", "Trebuchet MS", Verdana, sans-serif !important; border-width: 1px;padding: 10px;border-style: solid;border-color: white; background-color:#C3DDDB;}
.colorm {background-color:#58A09E; color:white;}
h3 {color:#BD3337 !important;}
</style>
 
<body>
<h3 style="color:#BD3337 !important;"> WARNING!!!</h3>
 
<p>There were errors during actrivation of Unified Messaging</p>
 
<p>Please check the errors and act accordingly</p>
 
<table>
 
"@
 
   $ErrorEmailDown = @"
</table>
</body>
"@
 
   $ErrorEmail = $ErrorEmailUp + $ErrorEmailResult + $ErrorEmailDown
 
   send-mailmessage `
      -To $To `
      -Subject "UM Activation Error Report $(Get-Date -format dd/MM/yyyy) - WARNING" `
      -Body $ErrorEmail `
      -BodyAsHtml `
      -Priority high `
      -UseSsl `
      -Port 587 `
      -SmtpServer 'smtp.office365.com' `
      -From $From `
      -Credential $EmailCredentials
}