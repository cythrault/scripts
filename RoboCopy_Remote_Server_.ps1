

 
<#     
.NOTES 
#=========================================================================== 
# Script: Robocoy_Remote_Server_.ps1  
# Created With:ISE 3.0  
# Author: Casey Dedeal 
# Date: 05/15/2016 20:51:50  
# Organization:  ETC Solutions
# File Name: Robocopy Files to Remote Server
# Comments:  
#=========================================================================== 
.DESCRIPTION 
        make sure to change these variables to make it fit into your scenario
        
        Change $Dest_PC = "PC001"
        Change $source  = "E:\Deployment\HyperV_"
        
#> 

$Space       = Write-host ""
$Dest_PC     = "PC001"
$source      = "E:\Deployment\HyperV_"
$dest        = "\\$Dest_PC\c$\ISO"
$Logfile     = "c:\temp\Robocopy1-$date.txt"
$date        = Get-Date -UFormat "%Y%m%d"
$cmdArgs     = @("$source","$dest",$what,$options) 
$what        = @("/COPYALL","/B","/SEC","/MIR")
$options     = @("/R:0","/W:0","/NFL","/NDL","/LOG:$logfile")

## Get Start Time
$startDTM = (Get-Date)

## Create ISO Folder on the Target Server if it does not exist

$Dest_PC     = "PC001"
$TARGETDIR   = "\\$Dest_PC\c$\ISO"
$Space 
$Space 

# Creatre ISO Folder ! does not exist on the Destination PC 
Write-host "Creating Folder....." -fore Green -back black
if(!(Test-Path -Path $TARGETDIR )){
    New-Item -ItemType directory -Path $TARGETDIR
}
Write-Host "........................................." -Fore Blue

## Provide Information
Write-host "Copying ISO into $PC....................." -fore Green -back black
Write-Host "........................................." -Fore Blue

## Kick off the copy with options defined 
robocopy @cmdArgs

## Get End Time
$endDTM = (Get-Date)

## Echo Time elapsed
$Time = "Elapsed Time: $(($endDTM-$startDTM).totalminutes) minutes"

## Provide time it took
Write-host ""
Write-host " Copy ISO to $PC has been completed......" -fore Green -back black
Write-host " Copy ISO to $PC took $Time        ......" -fore Blue
