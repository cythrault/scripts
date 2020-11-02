cls
$user = Get-ADUser camt055545 -Properties *
$props = $user.psobject.properties.name
write-host "getting list of users"
$users = Get-ADUser -Filter *
write-host "testing users"
foreach ($user in $users){
    try { $user | Get-ADUser -Properties * | Out-Null }
    catch {
        write-host "issue with user: $user.distinguishedName"
        foreach ($prop in $props){
			if (($prop -ne 'PropertyNames') -and ($prop -ne 'PropertyCount') -and ($prop -ne 'ModifiedProperties') -and ($prop -ne 'RemovedProperties') -and ($prop -ne 'AddedProperties')) {
				try { $user | Get-ADUser -Properties $prop | Out-Null }
				catch{
					write-host "issue with property: $prop"
					$user | Add-Member -MemberType NoteProperty -Name AttributeInError -Value $prop -Force
					}
				}
			}
        Write-Output $user
        }
    }
