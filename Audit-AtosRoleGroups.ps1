$RBACRoleGroups = Get-ADGroup -Filter { description -like "RBAC Role Group*" } -Properties description, memberOf, members
$RBACFuncGroups = Get-ADGroup -Filter { description -like "RBAC Function Group*" } -Properties description, memberOf, members
$start = Get-Date; $i=0
$RBACRoleGroups | sort | %{
    $groupDN = $_.DistinguishedName
    $groupName = $_.Name
    
    $i++ ; $pct = ($i / $RBACRoleGroups.Count) * 100
    $secondsElapsed = (Get-Date) - $start
    $secondsRemaining = ($secondsElapsed.TotalSeconds / $i) * ($RBACRoleGroups.Count - $i)
    Write-Progress -Activity "Validating $groupName membership" -CurrentOperation "Working..." -Status "Completed: $i of $($RBACRoleGroups.Count) - $([math]::Round($pct))%" -PercentComplete $pct -SecondsRemaining $secondsRemaining

    $_.memberOf | %{
        try {
            $group = Get-ADObject $_ -Properties description
            if ($group.description -notlike "RBAC Function Group*" ) { [pscustomobject]@{"Group Name" = $groupName; "Unauthorized Group" = $group.Name; "Unauthorized Group Description" = $group.description } }
        }
        catch { Write-Host -Background DarkRed $_ }
    }
}
$start = Get-Date; $i=0
$RBACFuncGroups | sort | %{
    $groupDN = $_.DistinguishedName
    $groupName = $_.Name

    $i++ ; $pct = ($i / $RBACFuncGroups.Count) * 100
    $secondsElapsed = (Get-Date) - $start
    $secondsRemaining = ($secondsElapsed.TotalSeconds / $i) * ($RBACFuncGroups.Count - $i)
    Write-Progress -Activity "Validating $groupName membership" -CurrentOperation "Working..." -Status "Completed: $i of $($RBACFuncGroups.Count) - $([math]::Round($pct))%" -PercentComplete $pct -SecondsRemaining $secondsRemaining

    $_.members | %{
        try {
            $group = Get-ADObject $_ -Properties description
            if ($group.description -notlike "RBAC Role Group*" ) { [pscustomobject]@{"Group Name" = $groupName; "Unauthorized Group" = $group.Name; "Unauthorized Group Description" = $group.description } }
        }
        catch { Write-Host -Background DarkRed $_ }
    }
}