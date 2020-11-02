#list of all available domains and their exchange references
$domains = @{
"gcg.local" = "cacol1exc03.gcg.local";
"focuscorp.ca"="cacol2exc01.focuscorp.ca";
"mmm.ca"="thes10-um01.mmm.ca";
"corp.pbwan.net"="dccan100exc03.corp.pbwan.net"}
#2010CORPUM
#2016CORPUM

#sequence (new user)
#1. create user
#2. enable skype (voice activated)
#3. enable UM

#sequence (skype migration)
#2. enable skype (voice activated)
#3. enable UM 

#todolist
#1. add the corp to UM
#2. enable skype (enaarcom10.corp.pbwan.net)
#3. conference rooms in skype (still enaarcom10)
#4. tryos (still enaarcom10)


#information needed from CSV
#PrimaryUserSMTPAddress
#PhoneNumber

#detect if activeDirectory module is present on pc, will exit if not available
if (Get-Module -ListAvailable -Name ActiveDirectory) {
        import-module activeDirectory
    } 
    else {
        Write-Host "ActiveDirectory module not installed, please make sure you are running powershellv2 or higher"
        write-host "press any key to exit"
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

#function to open a dialog box, allows user to pick csv file
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $openFileDialog.Title = "Please select the csv file for the activation"
    $file =  $OpenFileDialog.ShowDialog()
    return $openfiledialog.filename
}
Function write-log([string]$message,[string]$logpath){
    $message | Out-File -FilePath $logpath -Append
}

do{
    write-host "Please select a function" -ForegroundColor White
    write-host "1: Enable UM Users"
    write-host "2: Modify UM Users"
    write-host "3: Disable UM Users"
    $method = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $method = $method.character
}
while(@("1","2","3") -notcontains $method)

$csvfile = get-filename $([System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition))
$origcsv = import-csv $csvfile -Delimiter ';'
$users = @()

if($origcsv -eq $null){
    Write-Host "couldn't find any users in the CSV file"
        write-host "press any key to exit"
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
}

$date = get-date -Format "yyyy-MM-dd_HH-mm-ss" 
$outcsvName = $date + " usersList.csv"
$logfile = $date + " log.txt"
$logpath = "$([System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition))\$logfile"

##2. query ad, find org, create sublist per domain
foreach($user in $origcsv){
    $gooduser = $null
    $email = $user.PrimaryUserSMTPAddress
    foreach($key in $domains.Keys){
        $aduser = get-aduser -filter {emailaddress -eq $email} -Server $key -Properties canonicalname,msExchUserCulture,emailaddress
        if ($aduser.enabled -eq $true){
            $gooduser = $aduser
        } 
    }
    if($gooduser -ne $null){
        $objMigration = New-Object psobject | select PrimarySMTPaddress, UMMailboxPolicy, PhoneNumber, Domain, SIPAddress
        #set primarysmtpaddress based on the  CSV
        $objMigration.PrimarySMTPaddress = $email

        #set PhoneNumber based on the CSV
        $objMigration.PhoneNumber = $user.phonenumber

        #set domain based on the user's CanonicalName
        $objMigration.Domain = $gooduser.canonicalname.split("/")[0]

        #set SIPAddress based on the ad CORP account
        $erroractionpreference = "Stop"
        try{
        $objMigration.SIPAddress = (get-aduser $($gooduser.samaccountname) -Server "corp.pbwan.net" -properties "msRTCSIP-PrimaryUserAddress")."msRTCSIP-PrimaryUserAddress".split(":")[1] 
        }
        catch [Exception]{
            $message = "user $($gooduser.samaccountname) not replicated in CORP, skipping"
            write-host $message -ForegroundColor Red
            write-log $message -logpath $logpath
            Continue
        }
        $erroractionpreference = "Continue"
        #add user to list, will be done in batches per domain
        $users += $objMigration
        $message = "user $($user.primaryusersmtpaddress) found active in domain $($objMigration.domain)"
        write-host $message -ForegroundColor green
        write-log -message $message -logpath $logpath

        #set UMMailboxPolicy based on prefered language
        switch($objMigration.Domain){
            "gcg.local" { 
                    switch($gooduser.msExchUserCulture){
                        "fr-CA" {$objMigration.UMMailboxPolicy = "Canada-Fr-Ca Default Policy"}
                        Default {$objMigration.UMMailboxPolicy = "Canada Default Policy"}
                    }
                }
            "focuscorp.ca"{$objMigration.UMMailboxPolicy = "Canada-En-Us Default Policy"}
            "mmm.ca"{$objMigration.UMMailboxPolicy = "MMM-Canada-En Default Policy"}
            "corp.pbwan.net"{
                    switch($gooduser.msExchUserCulture){
                    "fr-CA" {$objMigration.UMMailboxPolicy = "CA-DialPlan-CA_Fr Default Policy"}
                        Default {$objMigration.UMMailboxPolicy = "CA-DialPlan-CA Default Policy"}
                    }
            }

        }

    }
    
    else{
        $message = "user $($user.primaryusersmtpaddress) not active on any of the $($domains.count) domains, please validate"
        write-host $message -ForegroundColor Red
        write-log -message $message -logpath $logpath
    }

}
    $users | Export-Csv -NoTypeInformation -Delimiter ';' -Path "$([System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition))\$outcsvName"

##3. remote session towards exchange server
foreach($key in $domains.keys){
    $sublist = $null
    $sublist = $users | ?{$_.domain -eq $key}   
    
    $erroractionpreference = "Stop"
    if($sublist -ne $null){
        $session = $null

        ##4. Invoke-command using session
        try{ 
        $session = New-PSSession -ComputerName $($domains[$key]) -Credential $(get-credential -Message "Please enter your $key credentials") -ConfigurationName UMConfiguration
        }
        catch [Exception]{
            $Error
            $message = "couldn't connect to the $key exchange server: $($domains[$key]), please validate your account/connectivity"
            write-host $message -ForegroundColor Red
            write-log -message $message -logpath $logpath
        }
        $erroractionpreference = "silentlycontinue"

        if($session -ne $null){

            switch($Method){
                ##6. invoke Method
                "1" { #Enable
                            $r = Invoke-Command -Session $session -ScriptBlock { Enable-UMUsers -users $using:sublist }
                            $r | %{write-log -message $_ -logpath $logpath}
                    }
                "2" { #Modify
                            $r = Invoke-Command -Session $session -ScriptBlock { Modify-UMUsers -users $using:sublist }
                            $r | %{write-log -message $_ -logpath $logpath}
                    }
                "3" { #Disable
                            $r = Invoke-Command -Session $session -ScriptBlock { Disable-UMUsers -users $using:sublist }
                            $r | %{write-log -message $_ -logpath $logpath}
                    }
            }

       }
    }
}

##todo: logging in exchange
##todo: use VSL/Calgary, depending on location (OU)

write-host "execution done"
write-host "press any key to exit"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")