function get-something {
	param(
		[Parameter(Mandatory=$true)]
		[string] $prenom,
		[Parameter(Mandatory=$true)]
		[string] $nom,
		[Parameter(Mandatory=$true)]
		[string] $domaine,
		[Parameter(Mandatory=$true)]
		[string] $id,
		[Parameter(Mandatory=$true)]
		[string] $site,
		[Parameter(Mandatory=$true)]
		[string] $ticket
	)

	new-object psobject -property @{
		prenom=$prenom
		nom=$nom
		domaine=$domaine
		id=$id
		site=$site
		ticket=$ticket
	}
}

function Get-FileName($InitialDirectory) {
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.InitialDirectory = $InitialDirectory
	$OpenFileDialog.filter = "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.Filename
}

function Nouvel-Utilisateur {
	param(
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$prenom,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$nom,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$domaine,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$id,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$site,
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$ticket
	)

Begin {
	# load Exchange and Lync PSSessions
	try {
		if (-Not $(Get-PSSession|where ConfigurationName -eq "Microsoft.Exchange")) {
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Exchange Session."
			$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://pvwexc0a.bnquebec.ca/PowerShell/ -Authentication Kerberos
			Import-PSSession $ExchSession | Out-Null
			Set-AdServerSettings -ViewEntireForest $True
		}

	# TODO add exchange online session

		if (-Not $(Get-PSSession|where {$_.ComputerName -like "*skypepool*"} ) ) {
			Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Loading Skype Session."
			$LyncSession = New-PSSession -ConnectionUri https://skypepoolfe01.banq.qc.ca/OcsPowerShell -Authentication NegotiateWithImplicitCredential
			Import-PSSession $LyncSession | Out-Null
		}
	}

	catch {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Error loading PS remote sessions - Can't proceed without valid Exchange and Skype sessions."
		break
	}

	# Check if commands are available
	if (-Not (Get-Command Get-CSUser) -or -Not (Get-Command Get-Mailbox)) {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Commands are not availables - Can't proceed without valid Exchange and Skype sessions."
		break
	}
}
	
Process {
	
	switch($domaine) { 
		"archives" { $logonscript = "LSanq.cmd" } 
		"biblio"   { $logonscript = "GBQ.cmd" } 
		default    { $logonscript = "" }
	}

	$nom = (Get-Culture).TextInfo.ToTitleCase($nom)
	$prenom = (Get-Culture).TextInfo.ToTitleCase($prenom)
	$domaine = $domaine.ToLower()
	$id = $id.ToLower()
	$site = $site.ToLower()

	#$pwd = "FrancoisLegault2018"
	$password = "!@#$%^&*0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz".tochararray() 
	$pwd = ($password | Get-Random -count 4) -join ''
	$upn = "$id@bnquebec.ca"

	if ( [bool](Get-ADUser -Filter { SamAccountName -eq $id } -Server bnquebec.ca:3268 ) -eq $True ) {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Le compte $($id) existe déjà!"
		return
	}
	
	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Création du compte $($id).."
	New-ADUser -SamAccountName "$id" -GivenName "$prenom" -Surname "$nom" -DisplayName "$nom $prenom" `
		-Name "$nom $prenom" –UserPrincipalName "$id@bnquebec.ca" -ScriptPath "$logonscript" `
		-AccountPassword $(ConvertTo-SecureString $pwd -AsPlainText -force) `
		-ChangePasswordAtLogon $true -Company "BAnQ" -Enabled $true -Server "$domaine.bnquebec.ca" `
		-Path "OU=$($site),OU=Employes,OU=Usagers,OU=1.Gestion,DC=$($domaine),DC=bnquebec,DC=ca" –Verbose

	$check = New-Object System.Collections.ArrayList
	do {
		$check = New-Object System.Collections.ArrayList
		[System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GlobalCatalogs.Name | Sort | %{
			$validation = [bool](Get-ADUser -Filter { SamAccountName -eq $id } -Server "$($psitem):3268" )
			$check.add($validation) | Out-Null
			Write-Host $psitem - $validation
		}
		$secs = 5
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Attente ($($secs) secondes) pour la réplication Active Directory du compte $($id).."
		Start-Sleep -Seconds $secs
	} while ( $check -contains $False )
	
	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Validation du compte pour $($id).."
	try { $nvuser = Get-ADUser -Identity $id -Server "$domaine.bnquebec.ca" -Properties * }
	catch { Write-Host -ForegroundColor Black -BackgroundColor Red "Le nouvel utilisateur ne semble pas avoir été crée"; break }
	
	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Ajout des groupes de base à $($id).."
	$groups = @()
	switch($site) { 
		"berri" { $groups += Get-ADGroup -Identity "Grande Bibliothèque" -Server biblio.bnquebec.ca }
		"holt"  { $groups += Get-ADGroup -Identity "Holt" -Server biblio.bnquebec.ca }
	}
	$groups += Get-ADGroup -Identity SSL.usagers -Server "$domaine.bnquebec.ca"
	$groups += Get-ADGroup -Identity SSL.Cadres -Server "$domaine.bnquebec.ca"
	$groups += Get-ADGroup -Identity SSL.Informatiques -Server "$domaine.bnquebec.ca"
	$groups += Get-ADGroup -Identity Tous -Server biblio.bnquebec.ca
	$groups += Get-ADGroup -Identity AvanProd -Server "bnquebec.ca"
	
	$groups | ForEach {
		Write-Host "Ajout du groupe $_ à $($id).."
		Add-ADGroupMember -Identity $_ -Members $nvuser
		}
		
	switch($domaine) {
		"archives" {
			switch($site) {
				"Quebec"     { $region = "CAC" }
				"Gatineau"   { $region = "CAG" }
				"Montreal"   { $region = "CAM" }
				"Quebec"     { $region = "CAQ" }
				"Rimouski"   { $region = "CAR" }
				"Abitibi"    { $region = "CARN" }
				"Sherbrooke" { $region = "CAS" }
				"7-Iles"     { $region = "CASI" }
				"3-Rivieres" { $region = "CATR" }
				default      { $region = "biblio" }
			}
		}
		"biblio" { $region = "biblio" }
		default  { $region = "biblio" }
	}

	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Création des répertoire P et V pour $($id).."
	$lecteurs = @()
	$lecteurs += "\\bnquebec.ca\Perso\$region\$id"
	$lecteurs += "\\bnquebec.ca\BAnQ\Archivage\$id"
	$lecteurs | ForEach {
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Création de $($_) pour $($id).."
		New-Item -Path $_ -ItemType Directory | Out-Null
		$acl = Get-Acl $_
		$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($id, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
		$acl.SetAccessRule($rule)
		Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Modification de la sécutité sur $($_) pour $($id).."
		Set-Acl $_ $acl
	}
 
	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Activation de la boîte aux lettre du compte $($id).."
	Set-AdServerSettings -ViewEntireForest $True
	$exchdb = (Get-MailboxDatabase | where {($_.IsExcludedFromProvisioning -eq $False) -and ($_.IsSuspendedFromProvisioning -eq $False ) }| Get-Random).Name
	$em = Enable-Mailbox -Identity $nvuser.DistinguishedName -Database $exchdb
    if ( (Get-Mailbox -Identity $nvuser.DistinguishedName -DomainController $em.OriginatingServer).IsMailboxEnabled -ne $True ) {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Échec de l'activation de la boîte au lettre sur le compte $($id).."
	}

	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Activation de Skype du compte $($id).."
	$es = Enable-CsUser -Identity $nvuser.UserPrincipalName -SipAddressType SamAccountName -SipDomain banq.qc.ca -RegistrarPool skypepoolfe01.banq.qc.ca -DomainController $em.OriginatingServer
    if ( (Get-ADUser -Identity $nvuser.samAccountName -Properties msRTCSIP-UserEnabled -Server $em.OriginatingServer)."msRTCSIP-UserEnabled" -ne $True ) {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Échec de l'activation de Skype sur le compte $($id).."
	}
		
	$nvuser = Get-ADUser -Identity $id -Server "bnquebec.ca:3268" -Properties *
	
	Write-Host -ForegroundColor Black -BackgroundColor DarkGreen "Envoi du courriel de confirmation pour $($id).."
	$body  = "Ceci est une notification afin de vous aviser suite à la création d'un nouveau compte.<br>"
	$body += "<br>"
	$body += "Billet C2 #<b>$($ticket)</b><br>"
	$body += "<br>"
	$body += "Prénom: <b>$($prenom)</b><br>"
	$body += "Nom: <b>$($nom)</b><br>"
	$body += "Affichage dans carnets: <b>$($nvuser.displayname)</b><br>"
	$body += "<br>"
	$body += "Détails du compte:<br>"
	$body += "Emplacement: <b>$($nvuser.DistinguishedName)</b><br>"
	$body += "Identifiant du compte (SamAccountName:) <b>$($domaine)\$($id)</b><br>"
	$body += "Identifiant du compte (UserPrincipalName:) $($nvuser.UserPrincipalName)<br>"
	$body += "Addresses de courriel et Skype: $($nvuser.mail)<br>"
	$body += "Mot de passe temporaire: <b>$($pwd)</b><br>"
	$body += "<br>"
	$body += "<br>"
	Send-MailMessage -From nepasrepondre@banq.qc.ca -To martin.thomas@banq.qc.ca -Encoding UTF8 -BodyAsHtml -Subject "#$($ticket) - Nouveau compte - $($prenom) $($nom) - $($id)" -SmtpServer courriel.banq.qc.ca -Body $body

}
}

#Start-Transcript -Path "New-BAnQ.log"

$importfile = Get-FileName

if ( $importfile ) {

	switch -wildcard ( [System.IO.Path]::GetExtension($importfile) ) {
		".csv" { $users = Import-Csv -Path $importfile }
		".xls*" { $users = Import-Excel -Path $importfile }
	}

	if ( $users ) { $users | Nouvel-Utilisateur }
	
} else {
	try { $user = iex (show-command get-something -passthru) }
	catch {
		Write-Host -ForegroundColor Black -BackgroundColor Red "Impossible d'obtenir l'information de(s) compte(s)."
		break
	}

	Write-Host "Information sur le compte"
	Write-Host "Prénom: $($user.prenom)"
	Write-Host "Nom: $($user.nom)"
	Write-Host "Domaine: $($user.domaine)"
	Write-Host "Identifiant: $($user.id)"
	Write-Host "Site: $($user.site)"
	Write-Host "Billet C2: $($user.ticket)"

	$options = [System.Management.Automation.Host.ChoiceDescription[]] @("Oui", "Non", "Quitter")
	$opt = $host.UI.PromptForChoice("Continuer", "Faites votre choix", $options, "0")
	if ($opt -eq 0 ) { $user | Nouvel-Utilisateur }

}

#Stop-Transcript

#  "david.gilmour", "richard.wright" | %{ Remove-Item "\\bnquebec.ca\Perso\biblio\$_"; Remove-Item "\\bnquebec.ca\BAnQ\Archivage\$_"; Remove-ADUser $_}