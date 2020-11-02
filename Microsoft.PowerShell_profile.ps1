Write-Host -Fore Black -Back Green "Loading functions."
Get-ChildItem -Path C:\Users\martin.thomas\OneDrive\scripts\autoload\*.ps1 | Foreach-Object{ Write-Host -Fore Black -Back Green "Loading $_"; . $_.FullName } 
Write-Host -Fore Black -Back Green "Done."
Start-Transcript C:\Users\martin.thomas\OneDrive\logs\$(Get-Date -UFormat %Y%m%d-%H%M%S)-ps-$($env:USERNAME).log