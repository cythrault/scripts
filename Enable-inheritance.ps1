function Reset-FolderACL {
param ([Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]$folder)
    begin { $pdc = ([System.Directoryservices.Activedirectory.Forest]::GetCurrentForest()).RootDomain.PDCRoleOwner.Name }
    process {
        $folder = Get-Item -Path $folder
        try { $user = Get-ADUser -Server $pdc`:3268 -Identity $folder.Name -ErrorAction SilentlyContinue }
        catch { Write-Host -ForegroundColor Red "AD account not found for $($folder.FullName)" }
        if ($user) {
            Write-Host -ForegroundColor Green "AD account found for $($folder.FullName)"
            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $permission = $($user.SamAccountName), 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow'
            $rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $permission
            $acl.SetAccessRule($rule)
            #$acl.SetOwner([System.Security.Principal.NTAccount] $($user.SamAccountName))
            $acl.SetOwner([System.Security.Principal.NTAccount] "BUILTIN\Administrators")
            $acl.SetAccessRuleProtection($false, $true)
            #$acl | fl *
            #$acl.Access
            Set-Acl -Path $folder.Fullname -AclObject $acl -Verbose -WhatIf
        }
    }
    end {
        if ($user) {Remove-Variable user, permission, rule, acl}
    }
}