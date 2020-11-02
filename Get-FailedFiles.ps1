#Requires -Version 4

function Log {
<# 
 .Synopsis
  Function to log input string to file and display it to screen

 .Description
  Function to log input string to file and display it to screen. Log entries in the log file are time stamped. Function allows for displaying text to screen in different colors.

 .Parameter String
  The string to be displayed to the screen and saved to the log file

 .Parameter Color
  The color in which to display the input string on the screen
  Default is White
  Valid options are
    Black
    Blue
    Cyan
    DarkBlue
    DarkCyan
    DarkGray
    DarkGreen
    DarkMagenta
    DarkRed
    DarkYellow
    Gray
    Green
    Magenta
    Red
    White
    Yellow

 .Parameter LogFile
  Path to the file where the input string should be saved.
  Example: c:\log.txt
  If absent, the input string will be displayed to the screen only and not saved to log file

 .Example
  Log -String "Hello World" -Color Yellow -LogFile c:\log.txt
  This example displays the "Hello World" string to the console in yellow, and adds it as a new line to the file c:\log.txt
  If c:\log.txt does not exist it will be created.
  Log entries in the log file are time stamped. Sample output:
    2014.08.06 06:52:17 AM: Hello World

 .Example
  Log "$((Get-Location).Path)" Cyan
  This example displays current path in Cyan, and does not log the displayed text to log file.

 .Example 
  "$((Get-Process | select -First 1).name) process ID is $((Get-Process | select -First 1).id)" | log -color DarkYellow
  Sample output of this example:
    "MDM process ID is 4492" in dark yellow

 .Example
  log "Found",(Get-ChildItem -Path .\ -File).Count,"files in folder",(Get-Item .\).FullName Green,Yellow,Green,Cyan .\mylog.txt
  Sample output will look like:
    Found 520 files in folder D:\Sandbox - and will have the listed foreground colors

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/06/2014
  v1.1 - 12/01/2014 - added multi-color display in the same line

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [String[]]$String, 
        [Parameter(Mandatory=$false,
                   Position=1)]
            [ValidateSet("Black","Blue","Cyan","DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","Red","White","Yellow")]
            [String[]]$Color = "Green", 
        [Parameter(Mandatory=$false,
                   Position=2)]
            [String]$LogFile,
        [Parameter(Mandatory=$false,
                   Position=3)]
            [Switch]$NoNewLine
    )

    if ($String.Count -gt 1) {
        $i=0
        foreach ($item in $String) {
            if ($Color[$i]) { $col = $Color[$i] } else { $col = "White" }
            Write-Host "$item " -ForegroundColor $col -NoNewline
            $i++
        }
        if (-not ($NoNewLine)) { Write-Host " " }
    } else { 
        if ($NoNewLine) { Write-Host $String -ForegroundColor $Color[0] -NoNewline }
            else { Write-Host $String -ForegroundColor $Color[0] }
    }

    if ($LogFile.Length -gt 2) {
        "$(Get-Date -format "yyyy.MM.dd hh:mm:ss tt"): $($String -join " ")" | Out-File -Filepath $Logfile -Append 
    } else {
        Write-Verbose "Log: Missing -LogFile parameter. Will not save input string to log file.."
    }
}

function Get-FailedFiles {
<# 
 .SYNOPSIS
  Script to parse Robocopy log file(s) and return files that failed to copy.

 .DESCRIPTION
  This script will parse Robocopy log file(s) and return a PS object that has 2 properties: FilePath and ErrorCode fro each file that failed to copy.

 .PARAMETER RoboLog
  Path to Robocopy log file

 .PARAMETER FailedFilesCSV
  Path to the output file that lists the files that failed to copy along with their error codes.

 .PARAMETER CopystatsCSV
  Path to the output file that lists summaries of files and folders copied/failed and copy duration.

 .PARAMETER Sum
  Optional switch that adds a line at the end of the CopyStatsCSV tallying all prior lines.

 .EXAMPLE
  Get-FailedFiles -RoboLog .\logs\Robo-Migrate-FileShares_NYFILSRV01P-2015-03-11_05-27-26AM.txt
  The script will return list of files that failed to copy and the corresponding error codes (in decimal)

 .EXAMPLE
  $FailedFiles = Get-FailedFiles (Get-ChildItem -Path .\logs | where { $_.Name.StartsWith('Robo') }).FullName
  This command will parse all the files under .\logs subfolder that start with 'Robo' and return the files that
  failed to copy and their error codes. This can be presented:
  $FailedFiles | FT -Auto                             # in tabular format
  $FailedFiles | Out-Gridview                         # in Powershell ISE Gridview
  $FailedFiles | Export-Csv .\FailedFiles.CSV -NoType # or exported to CSV

 .OUTPUTS 
  This function returns a PS object that has 2 properties: FilePath and ErrorCode.

 .LINK
  https://superwidgets.wordpress.com/category/powershell/

 .NOTES
  Function by Sam Boutros
  v1.0 - 03/11/2015
  v1.1 - 03/17/2015 - Added functionality to gather file and folder copy summaries and export to CSV

#>
	
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		    [ValidateScript({Test-Path -Path $_})][String[]]$RoboLog,
		[Parameter(Mandatory = $false, Position = 1)]
            [String]$FailedFilesCSV = "$PSScriptRoot\FailedFiles-$(Get-Date -format yyyy-MM-dd_hh-mm-sstt).CSV",
		[Parameter(Mandatory = $false, Position = 2)]
            [String]$CopyStatsCSV = "$PSScriptRoot\CopyStats-$(Get-Date -format yyyy-MM-dd_hh-mm-sstt).CSV",
        [Parameter(Mandatory = $false, Position = 3)] 
            [Switch]$Sum = $true
	)
	
    $FailedFiles = @()
    foreach ($LogFile in $RoboLog) {
        log 'Processing file',$LogFile,'...' Green,Cyan -NoNewLine
        $LogLines = Get-Content $LogFile
        log $LogLines.Count,'lines loaded' Cyan,Green

        log 'Detecting files that failed to copy, compiling list... ' -NoNewLine
        $LogLines -match ' ERROR ' -match 'Copying File' | % {
            $Props = [ordered]@{
                FilePath  = $_.Substring($_.IndexOf('Copying File') + 12 , $_.Length - $_.IndexOf('Copying File') - 12).Trim()
                ErrorCode = $_.Substring($_.IndexOf(' ERROR ')+7,2).Trim()
            }
            $FailedFiles += New-Object -TypeName PSObject -Property $Props
        }
        log 'detected',$Failedfiles.Count,'files that failed to copy' Green,Yellow,Green

        log 'Compiling file/folder copy statistics..' -NoNewLine
        $CopyStats = @()
        $LogLines | % {
            if ($_ -match '   Source : ') { $Source      = $_.Substring(12,$_.Length - 12 ) }
            if ($_ -match '     Dest : ') { $Destination = $_.Substring(12,$_.Length - 12 ) }
            if ($_ -match '    Dirs :')   { 
                $DirTotal   = [Long]$_.Substring(12,10).Trim() 
                $DirCopied  = [Long]$_.Substring(22,10).Trim() 
                $DirFailed  = [Long]$_.Substring(52,10).Trim() 
            }
            if ($_ -match '   Files : ' -and $_.Length -gt 60) { 
                $FileTotal  = [Long]$_.Substring(12,10).Trim() 
                $FileCopied = [Long]$_.Substring(22,10).Trim() 
                $FileFailed = [Long]$_.Substring(52,10).Trim() 
            }
            if ($_ -match '   Times :')  { $Duration  = [TimeSpan]$_.Substring(12,10).Trim() }
            if ($_ -match '   Speed : ') { $Speed = $_.Substring(12,$_.Length - 12 ).Trim() }
            if ($_ -match '   Ended : ') { 
                $Ended = $_.Substring(11,$_.Length - 11 ).Trim() 
                $Props = [ordered]@{
                    Source         = $Source
                    Destination    = $Destination
                    DirTotal       = $DirTotal
                    NewDirCopied   = $DirCopied
                    DirFailed      = $DirFailed
                    Filetotal      = $FileTotal
                    NewFileCopied  = $FileCopied
                    FileFailed     = $FileFailed
                    Duration       = $Duration
                    Speed          = $Speed
                    Ended          = $Ended
                }
                $CopyStats += New-Object -TypeName PSObject -Property $Props
            }    
        } # foreach $LogLines
        if ($Sum) {
            $Props = [ordered]@{
                Source         = 'Total'
                Destination    = ''
                DirTotal       = ($CopyStats | Measure-Object -Property DirTotal -Sum).Sum    
                NewDirCopied   = ($CopyStats | Measure-Object -Property NewDirCopied -Sum).Sum 
                DirFailed      = ($CopyStats | Measure-Object -Property DirFailed -Sum).Sum 
                Filetotal      = ($CopyStats | Measure-Object -Property Filetotal -Sum).Sum 
                NewFileCopied  = ($CopyStats | Measure-Object -Property NewFileCopied -Sum).Sum 
                FileFailed     = ($CopyStats | Measure-Object -Property FileFailed -Sum).Sum 
                Duration       = ''
                Speed          = ''
                Ended          = ''
            }
            $CopyStats += New-Object -TypeName PSObject -Property $Props
        }
        log 'done'
        log ($CopyStats | FT -Auto | Out-String)

    } # foreach $LogFile

    $FailedFiles | Export-Csv $FailedFilesCSV -NoType
    $CopyStats   | Export-Csv $CopyStatsCSV -NoType

} 