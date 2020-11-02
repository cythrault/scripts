param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string] $Prenom,
	[Parameter(Mandatory=$true,ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
	[string] $Nom
)

begin {
	$Skus = Get-AzureADSubscribedSku | Sort SkuPartNumber | Out-GridView -Title "Choisir le(s) plan(s)" -Passthru
	if ( !$Skus ) { Write-Host -ForegroundColor Black -BackgroundColor Red "Le nom du plan est obligatoire."; break }
}

process {
	[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
	$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
	$PasswordProfile.Password = [system.web.security.membership]::GeneratePassword(10,2)
	$PasswordProfile.EnforceChangePasswordPolicy = $true
	$PasswordProfile.ForceChangePasswordNextLogin = $true

 	$id = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes("$($prenom.ToLower()).$($nom.ToLower())"))
	$upn = "$($id)@banqtest.onmicrosoft.com"
	$displayname = "$($nom), $($prenom)"

	try {
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Création du compte $($upn)"
		New-AzureADUser -GivenName $nom -Surname $prenom -DisplayName $displayname -UserPrincipalName $upn -MailNickName $id -AccountEnabled $true -PasswordProfile $PasswordProfile -UsageLocation CA -PreferredLanguage fr-CA
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Ajout de(s) license(s)."
		$i=0
		$Skus | %{
			$license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
			$licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
			$license.SkuId = $Skus[$i].SkuId
			$licenses.AddLicenses = $license
			Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $licenses
			$i++
			Remove-Variable -Name license
			Remove-Variable -Name licenses
		}
	}

	catch {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Échec dans la création du compte ou l'assignation de la license pour $($upn)."
		Write-Host -ForegroundColor Red $_.Exception.Message
		return
	}

	$body  = "Ceci est une notification afin de vous aviser suite à la création d'un nouveau compte dans le portail de test de BAnQ.<br>"
	$body += "<br>"
	$body += "Prénom: <b>$($prenom)</b><br>"
	$body += "Nom: <b>$($nom)</b><br>"
	$body += "<br>"
	$body += "Identifiant du compte: $($upn)<br>"
	$body += "Mot de passe temporaire: <b>$($PasswordProfile.Password)</b><br>"
	$body += "License assignée au compte: <b>$($planName)</b><br>"
	$body += "<br>"
	$body += "<br>"

	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Envoi d'un courriel de notification."
	Send-MailMessage -From nepasrepondre@banq.qc.ca -To $id@banq.qc.ca -Cc martin.thomas@banq.qc.ca,Edgar.Delgado@banq.qc.ca -Encoding UTF8 -BodyAsHtml -Subject "Création compte portail BAnQ Test - $($prenom) $($nom)" -SmtpServer courriel.banq.qc.ca -Body $body
}

end { }