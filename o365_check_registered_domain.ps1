<#  
.SYNOPSIS  
    Check domain registration in O365/AAD tenants  
.DESCRIPTION  
    This script checks if a (SMTP) domain is already registered in an O365/AAD tenant
.PARAMS
    -domain: This parameter is the domain, which should be checked against O365/AAD (e.g.: bisser.at)
    -useAnotherAccount (OPTIONAL): If your IE already has cached some O365 credentials you can use this param to use the option "Use another acocunt"
.NOTES
    File Name  : o365_check_registered_domain.ps1  
    Author     : Stephan Bisser - stephan@bisser.at - blog.bisser.at
    Requires   : Internet Explorer
    Date       : 15.03.2017
    Version    : 1.0
.LINK  
#>

param(
[string]$domain,
[bool]$useAnotherAccount
)

$password = "password"
$url = "https://portal.office.com"
$username = "dummy@"+$domain
$errorMsg = $null
$er = $null
$ie = $null

Function Main
{
    switch ($useAnotherAccount) 
    { 
        $true {. BrowseURLWithAnotherAccount} 
        $false {. BrowseURL} 
        default {. BrowseURL}
    }
}

Function BrowseURL
{
    $ie = New-Object -com InternetExplorer.Application
    $ie.visible=$false
    $ie.navigate($url) 
    while($ie.ReadyState -ne 4) {
        start-sleep -m 1000
    }    

    $ie.document.IHTMLDocument3_getElementById("cred_userid_inputtext").value= $username
    $ie.document.IHTMLDocument3_getElementById("cred_password_inputtext").click()
    $ie.document.IHTMLDocument3_getElementById("cred_password_inputtext").value = $password
    start-sleep -m 2000
    $ie.document.IHTMLDocument3_getElementById("cred_sign_in_button").click()
    start-sleep -m 2000
    $er = $ie.Document.IHTMLDocument3_getElementById("cta_error_message_text")
    $errorMsg = $er.textContent
    . CheckErrorMsg
}


Function BrowseURLWithAnotherAccount
{
    $ie = New-Object -com InternetExplorer.Application
    $ie.visible=$false
    $ie.navigate($url) 
    while($ie.ReadyState -ne 4) {
        start-sleep -m 1000
    }    

    $ie.document.IHTMLDocument3_getElementById("use_another_account").click()
    $ie.document.IHTMLDocument3_getElementById("use_another_account_link").click()

    $ie.document.IHTMLDocument3_getElementById("cred_userid_inputtext").value= $username
    $ie.document.IHTMLDocument3_getElementById("cred_password_inputtext").click()
    $ie.document.IHTMLDocument3_getElementById("cred_password_inputtext").value = $password
    start-sleep -m 2000
    $ie.document.IHTMLDocument3_getElementById("cred_sign_in_button").click()
    start-sleep -m 2000
    $er = $ie.Document.IHTMLDocument3_getElementById("cta_error_message_text")
    $errorMsg = $er.textContent
    . CheckErrorMsg
}

Function CheckErrorMsg
{
    if ($errorMsg -like "*We don't recognize this domain name*"){
        Write-Host "Domain is not yet registered" -ForegroundColor Green
    } if ($errorMsg -like "*We don't recognize this user ID or password*"){
        Write-Host "Domain is already registered" -ForegroundColor Yellow
    }
    $ie.quit()
}

. Main