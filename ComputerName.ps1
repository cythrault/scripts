<#
.SYNOPSIS
Chech valid computer name, and change it if required
.DESCRIPTION
This script checks that the local computer name is less than 16 characters, and allows changing it accordingly
.EXAMPLE
.\ComputerName.ps1
.NOTES
Credentials required if the computer is joined to a domain
#>
cls
$computerName = [system.environment]::MachineName
$computerNameLength = $computerName.length

# Valid computer name
if ($computerNameLength -lt 16) {
    Write-Host("The computer name is $computerName, which is $computerNameLength characters.") -fore Green
}

# Invalid computer name
else {
    Write-Host("The computer name is $computerName, which is $computerNameLength characters. It must be changed!") -fore Red

    $newComputerName = Read-Host("`nPlease enter a new computer name with maximum length of 15 characters")

    # New computer name is valid
    if (($newComputerName.Length -lt 16) -and ($newComputerName.Length -gt 0)) {
        try {
            # Domain computer
            if ((gwmi win32_computersystem).partofdomain -eq $true) {
                $domainName = $env:userdomain + "\"
                Rename-Computer -NewName $newComputerName -DomainCredential $domainName -Force

            # Workgroup computer
            } else {
                Rename-Computer -NewName $newComputerName -Force
            }
        }
        catch {
            Write-Host("`nRename failed. Please update computer name manually.") -fore Red
            exit
        }
    }
    # New computer name is invalid
    else {
        Write-Host("`nComputer name must be maximum 15 characters. Exiting!") -fore Red
    }
}