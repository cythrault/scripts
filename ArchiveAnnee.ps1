#Param (
#	[Parameter(Mandatory=$True)]
#	[string]$annee,
#	[string]$dossier = ".\"
#)

function Compress-Item
{
    [OutputType([IO.FileInfo])]
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('FullName')]
        [string[]]
        # The path to the files/directories to compress.
        $Path,

        [string]
        # Path to destination ZIP file. If not provided, a ZIP file will be created in the current user's TEMP directory.
        $OutFile
    )

    begin {
        Set-StrictMode -Version 'Latest'
        #Use-CallerPreference -Cmdlet $PSCmdlet -Session $ExecutionContext.SessionState

        $zipFile = $null

        [byte[]]$data = New-Object byte[] 22
        $data[0] = 80
        $data[1] = 75
        $data[2] = 5
        $data[3] = 6
        [IO.File]::WriteAllBytes($OutFile, $data)

        $shellApp = New-Object -ComObject "Shell.Application"
        $copyHereFlags = (
                            # 0x4 = No dialog
                            # 0x10 = Responde "Yes to All" to any prompts
                            # 0x400 = Do not display a user interface if an error occurs
                           0x4 -bor 0x10 -bor 0x400        
                        )
        $zipFile = $shellApp.NameSpace($(split-path($OutFile)))
        $zipItemCount = 0

    }

    process {
        if( -not $zipFile ) { return }

        $Path | Resolve-Path | Select-Object -ExpandProperty 'ProviderPath' | ForEach-Object { 
            $zipEntryName = Split-Path -Leaf -Path $_
            [void]$zipFile.CopyHere($_, $copyHereFlags)
            while($zipfile.Items().Item($zipEntryName) -eq $null) { Start-sleep -seconds 1 }
            $entryCount = Get-ChildItem $_ -Recurse | Measure-Object | Select-Object -ExpandProperty 'Count'
            $zipItemCount += $entryCount
        }

    }
}

#Write-Host "Archivage des LOG et XML sous $($dossier) de l'année $($annee)"
#$fichiers = Get-ChildItem -Path $dossier | Where-Object { (($_.Name -like "*.xml") -or ($_.Name -like "*.log")) -and ($_.LastWriteTime.Year -eq $annee)}
#$fichiers | Compress-Archive -DestinationPath Archive-$annee.zip -CompressionLevel Optimal -Verbose
#if ($?) { $fichiers | Remove-Item -Verbose } else { Write-Host -ForegroundColor Black -BackgroundColor DarkRed "Archivage en erreur, les fichiers n'ont pas été supprimés." }