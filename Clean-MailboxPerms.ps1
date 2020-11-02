function Clean-MailboxPerms {
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[string]$id
	)

	Begin {
		$ErrorActionPreference = "Stop"
	}

	Process {
		
		# Mailbox Permissions
		$permissions = Get-MailboxPermission -Identity $id | ?{($_.User -like "S-*") -and ($_.IsInherited -eq $False)}
		$permissions | %{
			Write-Host "Removing Mailbox Permissions $($_.AccessRights) for $($_.User) from $($_.Identity) ($($_.IsInherited))"
			Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights $_.AccessRights -Deny:$($_.Deny) -Confirm:$True
		}
	
		# Calendar Folder Permissions
		try { $permissions = Get-MailboxFolderPermission -Identity "$($id):\Calendar" | ?{$_.User -like "Utilisateur NT*"} }
		catch { $permissions = Get-MailboxFolderPermission -Identity "$($id):\Calendrier" | ?{$_.User -like "Utilisateur NT*"} }
		$permissions | %{
			Write-Host -BackgroundColor DarkGreen "Removing Folder Permissions for $id"
			Remove-MailboxFolderPermission -Identity "$($id):\$($_.FolderName)" -User $_.User -Confirm:$True
		}
	}

	End {}

}
