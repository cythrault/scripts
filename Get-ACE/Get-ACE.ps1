<#-----------------------------------------------------------------------------
For a text file list of paths recursively get NTFS ACEs
from all folder ACLs where not inherited

Ashley McGlone (@GoateePFE)
http://aka.ms/GoateePFE
Microsoft Premier Field Engineer
March, 2014

LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not
intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
nonexclusive, royalty-free right to use and modify the Sample Code and to
reproduce and distribute the object code form of the Sample Code, provided
that You agree: (i) to not use Our name, logo, or trademarks to market Your
software product in which the Sample Code is embedded; (ii) to include a valid
copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
against any claims or lawsuits, including attorneys’ fees, that arise or result
from the use or distribution of the Sample Code.
-------------------------------------------------------------------------sdg-#>

#Requires -Version 3.0

Function Get-ACE {
Param (
        [parameter(Mandatory=$true)]
        [string]
        [ValidateScript({Test-Path -Path $_})]
        $Path
)

    $ErrorLog = @()

    Write-Progress -Activity "Collecting folders" -Status $Path `
        -PercentComplete 0
    $folders = @()
    $folders += Get-Item $Path | Select-Object -ExpandProperty FullName
	$subfolders = Get-Childitem $Path -Recurse -ErrorVariable +ErrorLog `
        -ErrorAction SilentlyContinue | 
        Where-Object {$_.PSIsContainer -eq $true} | 
        Select-Object -ExpandProperty FullName
    Write-Progress -Activity "Collecting folders" -Status $Path `
        -PercentComplete 100

    # We don't want to add a null object to the list if there are no subfolders
    If ($subfolders) {$folders += $subfolders}
    $i = 0
    $FolderCount = $folders.count

    ForEach ($folder in $folders) {

        Write-Progress -Activity "Scanning folders" -CurrentOperation $folder `
            -Status $Path -PercentComplete ($i/$FolderCount*100)
        $i++

        # Get-ACL cannot report some errors out to the ErrorVariable.
        # Therefore we have to capture this error using other means.
        Try {
            $acl = Get-ACL -LiteralPath $folder -ErrorAction Continue
        }
        Catch {
            $ErrorLog += New-Object PSObject `
                -Property @{CategoryInfo=$_.CategoryInfo;TargetObject=$folder}
        }

        $acl.access | 
            Where-Object {$_.IsInherited -eq $false} |
            Select-Object `
                @{name='Root';expression={$path}}, `
                @{name='Path';expression={$folder}}, `
                IdentityReference, FileSystemRights, IsInherited, `
                InheritanceFlags, PropagationFlags

    }
    
    $ErrorLog |
        Select-Object CategoryInfo, TargetObject |
        Export-Excel -Path ".\Errors_$($Path.Replace('\','_').Replace(':','_')).xlsx" `
            -FreezeTopRow -AutoSize

}

# Call the function for each path in the text file
Get-Content .\paths.txt | 
    ForEach-Object {
        If (Test-Path -Path $_) {
            Get-ACE -Path $_ |
                Export-Excel `
                    -Path ".\ACEs_$($_.Replace('\','_').Replace(':','_')).xlsx" `
                    -FreezeTopRow -AutoSize
        } Else {
            Write-Warning "Invalid path: $_"
        }
    }
