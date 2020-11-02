<#
=====================================
Script Created by - Binu Balan      
Script Created on - 11/8/2015
Version - V 1.1
Requirement *
PowerShell = 2.0 or above
This script is used to perform query
Huge IIS Log Files
	 .__.
     (oo)____
     (__)    )\
        ll--ll '
=====================================
#>
cls
Write-Host " "
Write-Host " "
Write-host " AAAAAAAA     PPPPPPPPP    PPPPPPPPP    UUU    UUU" -ForegroundColor Green
Write-host "AAAAAAAAAA    PPP   PPPP   PPP   PPPP   UUU    UUU" -ForegroundColor Green
Write-host "AAA    AAA    PPP    PPP   PPP    PPP   UUU    UUU" -ForegroundColor Green
Write-host "AAAAAAAAAA    PPPPPPPP     PPPPPPPP     UUU    UUU" -ForegroundColor Green
Write-host "AAA    AAA    PPP          PPP          UUU    UUU" -ForegroundColor Green
Write-host "AAA    AAA    PPP          PPP           UUUUUUUU" -ForegroundColor Green
Write-Host " "
Write-Host " " 
Write-host "	           .__." -ForegroundColor Green
Write-host "                   (oo)____" -ForegroundColor Green
Write-host "                   (__)    )\" -ForegroundColor Green
Write-host "                      ll--ll '" -ForegroundColor Green
Write-Host "               SCRIPT BY BINU BALAN               " -ForegroundColor DarkYellow -BackgroundColor DarkBlue 
Write-Host " "
Write-Host " "


$i = 1

# Getting Input from User
# =======================
Write-Host " "
Write-Host " "
Write-Host "Pre-Requisite Check for the Logparser.exe on local path" -NoNewline

Start-Sleep -Seconds 2
If(Test-Path -Path Logparser.exe){
Write-Host "                 [   OK   ]" -ForegroundColor Green
} Else {
Write-Host "                 [ Failed ]" -ForegroundColor Red
Write-Host " "
Write-Host " "
Write-Warning "Either Logparser is not installed or you are running this script on a different folder where Logparser.exe file is unavailable."
Write-Host " "
Write-Host "To download logparser follow this link : " -NoNewline -BackgroundColor Yellow -ForegroundColor Black
Write-Host "http://www.microsoft.com/en-in/download/details.aspx?id=24659" -ForegroundColor Blue -BackgroundColor Yellow
Write-Host " "
exit
}



Write-Host " "
Write-Host " "
$ReportPath = Read-Host "Enter Report Folder Path [Ex: c:\reports] "
$ReportName = Read-Host "Enter the report file name [Ex: LogReport.csv] "
Write-Host " "
Write-Host " "

Write-Host "Select the query type that you want to perform against the log?" -ForegroundColor Yellow
Write-Host "<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> " -ForegroundColor Blue
Write-Host "1. Date             Example -2014-12-02" -ForegroundColor Green
Write-Host "2. Time             Example -00:00:12" -ForegroundColor Green
Write-Host "3. s-ip             Example -192.168.1.20" -ForegroundColor Green
Write-Host "4. cs-method        Example - GET/POST" -ForegroundColor Green
Write-Host "5. cs-uri-query     Example - /EWS/Exchange.asmx" -ForegroundColor Green
Write-Host "6. s-port           Example -443" -ForegroundColor Green
Write-Host "7. cs-username      Example -appudomain\binu.balan" -ForegroundColor Green
Write-Host "8. c-ip             Example -106.33.98.222" -ForegroundColor Green
Write-Host "9. cs(user-Agent)   Example -Microsoft+office" -ForegroundColor Green
Write-Host "10. sc-status       Example -401" -ForegroundColor Green
Write-Host "11. sc-substatus    Example -1" -ForegroundColor Green
Write-Host "12. sc-win32-status " -ForegroundColor Green
Write-Host "13. time-taken" -ForegroundColor Green
Write-Host " "
Write-Host " "

$GetInput = Read-host "Enter the query number "

switch ($GetInput) 
    { 
        1 {$WhereVal = "Date"} 
        2 {$WhereVal = "Time"} 
        3 {$WhereVal = "s-ip"} 
        4 {$WhereVal = "cs-method"} 
        5 {$WhereVal = "cs-uri-query"} 
        6 {$WhereVal = "s-port"} 
        7 {$WhereVal = "cs-username"} 
        8 {$WhereVal = "c-ip"} 
        9 {$WhereVal = "cs(user-Agent)"} 
        10 {$WhereVal = "sc-status"} 
        11 {$WhereVal = "sc-substatus"} 
        12 {$WhereVal = "sc-win32-status"} 
        13 {$WhereVal = "time-taken"} 

        default {"You have input invalid data !!"}
    }

if ($WhereVal -eq $null) {

Write-Host "You have entered invalid data. Exiting the Script"

exit

}

Write-host "Enter Log folder path. For multiple folders use Comma separated value [Example [C:\Log1,C:\Log2"
$ORRFolderpath = Read-Host "Enter here "

$WhereQuery = Read-Host ("Enter the Query for $WhereVal ")
$EachFolder = $ORRFolderpath.Split(",")
$EachIP = $WhereQuery.Split(",")


#$RName = Read-Host ("Enter Report Name with CSV Extension - (Result.csv)")
#$RPath = Read-Host ("Enter the path where you want to store the report - (C:\Report)")



ForEach ($S_Folder in $EachFolder) {

    ForEach ($IP in $EachIP) {

    $LogPath = $S_Folder
    $FileNames = Get-childItem -Path $LogPath


            ForEach ($File in $FileNames) {

            Write-Host "Script Line $i - .\LogParser.exe SELECT * INTO $ReportPath\$ReportName FROM $LogPath\$File WHERE $WhereVal LIKE '%$IP%' -filemode:0" -ForegroundColor Yellow


            $i = $i + 1

            .\LogParser.exe "SELECT * INTO $ReportPath\$ReportName FROM $LogPath\$File WHERE $WhereVal LIKE '%$IP%'" -filemode:0

            Write-Host "[ Completed ]" -ForegroundColor Green

            Write-Host "    "
            Write-Host "    "

            }


    }

}