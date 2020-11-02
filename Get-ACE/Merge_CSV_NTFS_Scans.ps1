<#-----------------------------------------------------------------------------
Merge all ACE*.* CSV files into a single CSV, noting server and share
**Run this script from the folder where the output files are located.

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

Remove-Item -Path .\NTFSScan.csv -Force -Confirm:$false

$CSVs = Get-ChildItem ACE*.* | Select-Object Fullname, DirectoryName, Name

ForEach ($CSV in $CSVs) {
    Import-CSV $CSV.fullname | Export-Csv .\NTFSScan.csv -Append -NoTypeInformation
}

Get-Item .\NTFSScan.csv
