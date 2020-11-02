# ***************************************************************************************
# ***
# *** Scriptname: AnalyzeRobocopyLogs.ps1
# ***
# *** Usage: Get-Help .\AnalyzeRobocopyLogs.ps1
# ***
# *** Author: Powershell Administrator a.k.a. OcinO a.k.a. Nico Buma
# ***
# *** Version: 2.1
# ***
# *** History: 05-06-14 Fixed a Powershell v2 issue, where no error and warning lines were added to the summary
# ***          03-05-14 Totally rewrote script to increase performance and make it more diverse (2.0)
# ***                   - Changed: Added more input parameters and switches to make the script more universal
# ***                   - Added MaxLinesPerFile parameter, to increase speed even more and reduce memory usage
# ***                   - New: Error-handling instead of just breaking out of the script
# ***                   - New: changed complete reading of files to select-string
# ***                   - Script speed increased over 500% and CPU usage decreased to a minimum
# ***          30.04.14 Rewrote VBS Script in Powershell (1.0)
# ***          21.10.08 Created script in VBS (1.0)
# ***
# *** 2014 - Powershell Administrator. Use at your own risk.
# ***
# *** Speed measuring data:
# *** Reading of 110 files (95,8MB), which consist of 7 errors in 4 files and 402924 warnings in 4 files;
# *** 2m8s
# *** Tested with option SummaryAll (maximum output) and a MaxLinesPerFile of 500
# ***
# *** Same conditions as above, but tested with SummaryAllWithoutWarnings
# *** 20s
# ***
# ***************************************************************************************

<#
.SYNOPSIS 
Analyze Robocopy logs and mail the result in a summary file 

.DESCRIPTION
This Powershell script will analyze a directory (can be recursive) that contains Robocopy log files.
It will generate one output file with a summary for each log file, in the same folder as the script.
All logs that contain errors will be placed in top of this summary. All error information is included.
Once the summary is created an e-mail is sent to the specified e-mail address(es), with the summary file attached.
    
.PARAMETER Path
This option is mandatory. Specifies the path to the Robocopy log files.

.PARAMETER Recurse
Set this switch to include subfolders.

.PARAMETER MailSender
Specifies the from e-mail address. The default value is Norepl@Noreply.com

.PARAMETER MailRecipient
This option is mandatory. Specifies the recipient of the e-mail message.
If multiple recipients need to be specified, use a semicolon ( ; ) as a seperator.

.PARAMETER MailServer
This option is mandatory. Specifies the name of the SMTP server.

.PARAMETER MailSubject
Specifies the non-dynamic text part of the subject name for the e-mail message. The default text is "Analyze Robocopy Backup Logs"

.PARAMETER MailPort
Specifies the port to use for the e-mail message. The default value is 25. This option is ignored with Powershell v2.0

.PARAMETER UseSsl
Set this switch to use SSL for e-mail communication.

.PARAMETER Summary
Use one of the following parameters:
SummaryAll = creates a summary file of all Robocopy logs, includes detailed warning and error information
 in the report
SummaryAllWithoutWarnings = creates a summary file of all Robocopy logs, includes detailed error information
 in the report
ErrorsAndWarnings = creates a summary file of all Robocopy logs that contain errors and warnings,
 includes detailed warning and errorinformation. If no warnings and/or errors are found, the summary
 will only contain the start time and end time of the script.
ErrorsOnly = creates a summary file of all Robocopy logs that contain errors, includes detailed
 error information in the report. If no errors are found, the summary will only contain the start time
 and end time of the script.
(When AnalyzeRobocopyLogs encounters a 'same' file in a Robocopy log file, It will mark this as a warning)
The default option is SummaryAll

.PARAMETER MailPolicy
Use one of the following parameters:
Always = Always send an e-mail report with the summary file attached
ErrorsAndWarnings = Only send an e-mail report if the summary file contains Errors and/or Warnings
ErrorsOnly = Only send an e-mail report if the summary file contains Errors
(When AnalyzeRobocopyLogs encounters a 'same' file in a Robocopy log file, It will mark this as a warning)
The default option is Always

.PARAMETER MaxAmountOfSummaryLines
Use this parameter to set the maximum amount of detailed summary lines to include in the summary repory. The default value is 500.

.INPUTS
None. You cannot pipe objects.

.OUTPUTS
None. No object outputs.

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com;Suzie@q.com" -MailServer mail.server.com
This will analyze the "C:\Windows\Logs" folder for all Robocopy backup logs and sends the result by e-mail to
to John@doe.com and Suzie@q.com by e-mail.
It will use the default summary, mail policy, subject and mail port options.
In this example, it will NOT use SSL or recurively check the path for logs.

C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -Recurse -MailRecipient "John@doe.com" -MailServer mail.server.com
This will analyze the "C:\Windows\Logs" folder and subfolders for all Robocopy backup logs and sends the result
by e-mail to John@doe.com.

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -Recurse -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary SummaryAll -MailPolicy ErrorsOnly
This will analyze the "C:\Windows\Logs" folder (and subfolders) and check for all Errors and Warnings
 and creates a summary file of EACH log file, including a more detailed information for each error and/or
 warning, and will mail if Errors are found.
(so warnings are ignored when checking if the summary has to be mailed)

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary SummaryAllWithoutWarnings -MailPolicy Always
This will analyze the "C:\Windows\Logs" folder and check for all Errors and creates a summary file of
 each log file and ignores warnings. It will always send a mail report.

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary SummaryAllWithoutWarnings -MailPolicy ErrorsAndWarnings
This will analyze the "C:\Windows\Logs" folder and check for all Errors and creates a summary file of
 each log file and ignores warnings. It will send a mail report only if Errors are found.
(if you disable warnings in -summary the ErrorsAndWarnings mailpolicy will automatically change its
 behavior to ErrorsOnly)

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary SummaryAllWithoutWarnings -MailPolicy ErrorsOnly
This will analyze the "C:\Windows\Logs" folder and check for all Errors and creates a summary file of
 each log file and ignores warnings. It will send a mail report only if Errors are found.

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary ErrorsAndWarnings -MailPolicy ErrorsOnly
This will analyze the "C:\Windows\Logs" folder and check for all Errors and creates a detailed summary
 file of each log file that contains Errors and/or Warnings (skips the non-error and/or non-warning files).
It will send a mail report only if Errors are found.

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary ErrorsOnly -MailPolicy Always
This will analyze the "C:\Windows\Logs" folder and check for all Errors and creates a detailed summary
 file of each log file that contains Errors (skips the non-error files).
It will always send a mail report.

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -MailRecipient "John@doe.com" -MailServer mail.server.com -Summary ErrorsOnly -MailPolicy ErrorsAndWarnings
This will analyze the "C:\Windows\Logs" folder and check for all Errors and creates a detailed summary
 file of each log file that contains Errors (skips the non-error files).
 It will send a mail report only if Errors are found.
(if you disable warnings in -summary the ErrorsAndWarnings mailpolicy will automatically change its behavior to ErrorsOnly)

.EXAMPLE
C:\PS> .\AnalyzeRobocopyLogs.ps1 -Path "C:\Windows\Logs" -Recurse -Mailsender "NoReply@mydomain.com" -MailRecipient "John@doe.com" -MailServer "mail.server.com" -MailSubject "PowershellAdministrator - Removing the human error step by step" -MailPort 465 -UseSsl -Summary ErrorsOnly -MailPolicy ErrorsOnly -MaxAmountOfSummaryLines 1000
This will recursively analyze the "C:\Windows\Logs" folder and check for all Errors and creates a detailed
 summary file of each log file that contains Errors (skips the non-error files) and will use a maximum of
 1000 lines of detailed information.
It will send a mail report if Errors are found ans use NoReply@mydomain.com as from address
It will use SSL when sending mail and uses port 465 (the default SSL SMTP port) in this example.
It will also use "PowershellAdministrator - Removing the human error step by step" as non-dynamic part
 of the e-mail subject string.

.LINK
http://powershelladministrator.wordpress.com/2014/05/03/analyze-robocopy-log-files-and-mail-the-result/
http://PowershellAdministrator.wordpress.com
#>

Param(
[Parameter(Mandatory=$true)][string]$Path,
[switch]$Recurse,
[string]$MailSender = "Noreply@Noreply.com",
[Parameter(Mandatory=$true)][string]$MailRecipient,
[Parameter(Mandatory=$true)][string]$MailServer,
[string]$MailSubject = "Analyze Robocopy Backup Logs",
[int]$MailPort = 25,
[switch]$UseSsl,
[ValidateSet("SummaryAll","SummaryAllWithoutWarnings","ErrorsAndWarnings", "ErrorsOnly")][String[]]$Summary = "SummaryAll",
[ValidateSet("Always", "ErrorsAndWarnings", "ErrorsOnly")][String[]]$MailPolicy = "Always",
[int]$MaxAmountOfSummaryLines = 500
)

#Requires –Version 2.0

#Get start time
$StartTime = (Get-Date).ToString()

#Setting default variables
If($PSVersionTable.PSVersion.Major -eq 2) { $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent }
$LogFile = "$PSScriptRoot\AnalyzeRobocopyLogs.log"

#Summary variables
$SummaryAll = $false
$SummaryAllWithoutWarnings = $false
$ErrorsAndWarnings = $false
$ErrorsOnly = $false

#MailPolicy variables
$MailAlways = $false
$MailErrorsAndWarnings = $false
$MailErrorsOnly = $false

Switch($Summary)
{
    "SummaryAll" { $SummaryAll = $true }
    "SummaryAllWithoutWarnings" { $SummaryAllWithoutWarnings = $true }
    "ErrorsAndWarnings" { $ErrorsAndWarnings = $true }
    "ErrorsOnly" { $ErrorsOnly = $true }
    Default { Write-Error -Message "You entered something other than one of the 4 possible options" -Category ParserError -RecommendedAction "Use one of the 4 options as provided by the script." -CategoryReason "Incorrect option for Summary" -ErrorAction Stop -ErrorVariable $Summary }
}

Switch($MailPolicy)
{
    "Always" { $MailAlways = $true }
    "ErrorsAndWarnings" { $MailErrorsAndWarnings = $true }
    "ErrorsOnly" { $MailErrorsOnly = $true }
    Default { Write-Error -Message "You entered something other than one of the 3 possible options" -Category ParserError -RecommendedAction "Use one of the 3 options as provided by the script." -CategoryReason "Incorrect option for MailPolicy" -ErrorAction Stop -ErrorVariable $MailPolicy }
}

#Check if the given mail addresses are in the correct mail address format (contains text before and after the [at]. After the [at] and the text, it expects a [dot] with a 2 to 4 letter TLD after that)
$Regex = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$"
$MailAddresses = @()
If($MailRecipient.Contains(";")) { $MailAddresses = $MailRecipient.Split(";") }
Else { $MailAddresses = $MailRecipient }
Foreach($MailAddress in $MailAddresses)
{
    If($MailAddress -notmatch $Regex) { Write-Error -Message "Incorrect recipient e-mail address entered. Please verify that $MailAddress has been entered correctly" -Category ParserError -RecommendedAction "Correct the e-mail address or use the -MailRecipient switch to specify one." -CategoryReason "Incorrect E-mail address" -ErrorAction Stop -ErrorVariable $MailAddress }
}

# Retrieve list of files from a folder
#*************************************************
If(Test-Path($Path))
{
    #Get a list of all files if the path exists
    $FileList = Get-ChildItem -Path $Path -Recurse:($Recurse.IsPresent)
    
    $ErrorMessages = $FileList | Select-String -Pattern '(ERROR: )|( ERROR )' -CaseSensitive
    If($SummaryAll -or $ErrorsAndWarnings) { $WarningMessages = $FileList | Select-String -SimpleMatch "    same" -CaseSensitive }
    #Get all lines which robocopy uses for its summary text
    $LineNumbers = $FileList | Select-String -SimpleMatch "------------------------------------------------------------------------------"
    
    $ErrorFiles = ""
    $WarningFiles = ""
    $SummaryText = ""
    $Newline = "`r`n"
    $StrTextDivider = ""
    #71 is the size of the maximum width of the robocopy text line
    For($i = 0; $i -lt 71; $i++) { $StrTextDivider += "-" }
    $FileWithErrors = $false
    
    If($PSVersionTable.PSVersion.Major -eq 2)
    {
        $TempErrorFiles = @()
        Foreach($ErrorMessage in $ErrorMessages)
        {
            $TempErrorFiles += $ErrorMessage.Path
        }
        $ErrorFiles = $TempErrorFiles | Select-Object -Unique
    }
    Else
    {
        $ErrorFiles = $ErrorMessages.Path | Select-Object -Unique
    }
    If($SummaryAll -or $ErrorsAndWarnings)
    {
        If($PSVersionTable.PSVersion.Major -eq 2)
        {
            $TempWarningFiles = @()
            Foreach($WarningMessage in $WarningMessages)
            {
                $TempWarningFiles += $WarningMessage.Path
            }
            $WarningFiles = $TempWarningFiles | Select-Object -Unique
        }
        Else
        {
            $WarningFiles = $WarningMessages.Path | Select-Object -Unique
        }
    }

    If($ErrorsAndWarnings -or $ErrorsOnly)
    {
        #If only Errors and/or Warnings are enabled, reduce the filelist to only include these files
        $ListOfFiles = $ErrorFiles
        If($ErrorsAndWarnings)
        {
            $ListOfFiles += $WarningFiles
            $ListOfFiles = $ListOfFiles | Select-Object -Unique
        }
        $FileList = Get-ChildItem -Path $ListOfFiles
    }

    Foreach($File in $FileList)
    {
        $BottomText = ""
        $FileLineNumbers = $LineNumbers | Where-Object { $_.Path -eq $File.FullName }
        
        If($FileLineNumbers)
        {
            #Add all Robocopy summary text to the summary file
            $BeginStartLine = $FileLineNumbers[1].LineNumber
            $BeginEndLine = $FileLineNumbers[2].LineNumber - 3
            $EndStartLine = $FileLineNumbers[3].LineNumber
            $EndEndLine = $FileLineNumbers[3].LineNumber + 20
            $FileLinesToRead = $BeginStartLine..$BeginEndLine + $EndStartLine..$EndEndLine

            #Check if the file is in the list which contains the errors
            #*************************************************
            $IsError = $ErrorFiles | Select-String -SimpleMatch $File.FullName
            If($IsError)
            {
                $ErrorLines = $ErrorMessages | Where-Object { $_.Path -eq $File.FullName }
                [int[]]$ErrorLineNumbers = $ErrorLines | foreach { ($_.LineNumber)-2; ($_.LineNumber)-1 }
                $ErrorLineNumbers = $ErrorLineNumbers | Sort-Object -Unique
                If($ErrorLineNumbers.Count -gt $MaxAmountOfSummaryLines) { $ErrorLineNumbers = $ErrorLineNumbers[0..$MaxAmountOfSummaryLines] }
                $FileLinesToRead = $FileLinesToRead + $ErrorLineNumbers
                $FileWithErrors = $true
            }
            #Check if the file is in the list which contains the warnings (if SummaryAll is selected and no errors were found)
            #*************************************************
            $IsWarning = $WarningFiles | Select-String -SimpleMatch $File.FullName
            If($IsWarning -and ($SummaryAll -or $ErrorsAndWarnings))
            {
                $WarningLines = $WarningMessages | Where-Object { $_.Path -eq $File.FullName }
                [int[]]$WarningLineNumbers = $WarningLines | foreach { ($_.LineNumber)-2 }
                $WarningLineNumbers = $WarningLineNumbers | Sort-Object -Unique
                If($WarningLineNumbers.Count -gt $MaxAmountOfSummaryLines) { $WarningLineNumbers = $WarningLineNumbers[0..$MaxAmountOfSummaryLines] }
                $FileLinesToRead = $FileLinesToRead + $WarningLineNumbers
            }
            If($FileWithErrors -and ($SummaryAll -or $SummaryAllWithoutWarnings))
            {
                $BottomText += $SummaryText
                $SummaryText = ""
            }
            $SummaryText += $StrTextDivider + $Newline
            $SummaryText += "`t`t`t" + $File.Name + $Newline
            $SummaryText += $StrTextDivider + $Newline

            $TextInFile = [io.file]::ReadAllLines($File.FullName)

            #Create a list with unique file line numbers
            $UniqueFileLinesToRead = $FileLinesToRead | Sort-Object -Unique

            Foreach($Line in $UniqueFileLinesToRead)
            {
                If($Line -lt $TextInFile.Count -1) { $SummaryText += $TextInFile[$Line + 1] + $Newline }
            }
            If($FileWithErrors -and ($SummaryAll -or $SummaryAllWithoutWarnings))
            {
                $SummaryText += $BottomText
                $FileWithErrors = $false
            }
        }
    }

    #Get end time after reading all info and saving the start and end time of the analyzer to the log file
    $EndTime = (Get-Date).ToString()
    $SummaryEndLogText = "AnalyzeRobocopyLogs completed succesfully"
    $StrTextDivider = ""
    For($i = 0; $i -lt $SummaryEndLogText.Length; $i++) { $StrTextDivider += "-" }
    $SummaryText += $StrTextDivider + $Newline
    $SummaryText += $SummaryEndLogText
    #Save the analyzed summary to the logfile
    [io.file]::WriteAllText($LogFile,$StartTime + $Newline + $SummaryText + $Newline + $EndTime)

    #Get amount of errors
    [int]$AmountOfErrors = $ErrorMessages.Count
    [int]$AmountOfErrorFiles = $ErrorFiles.Count
    #If enabled; get amount of warnings
    If($SummaryAll -or $ErrorsAndWarnings)
    {
        [int]$AmountOfWarnings = $WarningMessages.Count
        [int]$AmountOfWarningFiles = $WarningFiles.Count
    }

    #Defining the conclusion text and creating the summary text based on analyzing the logs
    #*************************************************    
    If($AmountOfErrors -eq 0)
    {
        #Check if warning reporting is enabled
        $BgColor = "bgcolor=""lime"""
        If($SummaryAll -or $ErrorsAndWarnings)
        {
            #Create subject with warnings enabled and based on findings
            If($AmountOfWarnings -eq 0) { $SummaryConclusion = "JOB SUCCESS: $MailSubject Completed Successfully; $AmountOfErrors errors and $AmountOfWarnings warnings" }
	        Else
            {
     		    $BgColor = "bgcolor=""yellow"""
                $SummaryConclusion = "JOB SUCCESS, BUT WITH WARNINGS: $MailSubject Completed with Warnings; $AmountOfErrors errors and $AmountOfWarnings warnings"
            }
        }
        #Create subject without warnings enabled
        Else { $SummaryConclusion = "JOB SUCCESS: $MailSubject Completed Successfully; $AmountOfErrors errors" }
    }
    #The job failed; errors have been found
    Else
    {
    	$BgColor = "bgcolor=""red"""
	    #Check if warning reporting is enabled
        If($SummaryAll -or $ErrorsAndWarnings) { $SummaryConclusion = "JOB FAILED: $MailSubject Failed; $AmountOfErrors errors and $AmountOfWarnings warnings" }
        #Create subject without warnings enabled
        Else { $SummaryConclusion = "JOB FAILED: $MailSubject Failed; $AmountOfErrors errors" }
    }

    $SummaryText = "<TABLE border=""2"" $BgColor><TR><TD align=""center"" colspan=""4""><B><FONT size=""+2"">Results</FONT></B></TD></TR>" + $Newline
    $SummaryText += "<TR><TD align=""center""> </TD>"
    $SummaryText += "<TD>No.</TD>"
    $SummaryText += "<TD>in .. Log File(s)</TD></TR>" + $Newline
    If($AmountOfErrors -eq 0) { $BgColor = "bgcolor=""lime""" }
    Else { $BgColor = "bgcolor=""red""" }
    $SummaryText += "<TR><TD align=""center"" $BgColor><B>Errors:</B></TD>"
    $SummaryText += "<TD align=""center"" $BgColor><B>$AmountOfErrors</B></TD>"
    $SummaryText += "<TD align=""center"" $BgColor>$AmountOfErrorFiles</TD></TR>" + $Newline
    
    If($SummaryAll -or $ErrorsAndWarnings)
    {
        If($AmountOfWarnings -eq 0) { $BgColor = "bgcolor=""lime""" }
        Else { $BgColor = "bgcolor=""yellow""" }
        $SummaryText += "<TR><TD align=""center"" $BgColor><B>Warnings:</B></TD>"
        $SummaryText += "<TD align=""center"" $BgColor><B>$AmountOfWarnings</B></TD>"
        $SummaryText += "<TD align=""center"" $BgColor>$AmountOfWarningFiles</TD></TR>" + $Newline
    }

    Function Send-Mail
    {
	If($PSVersionTable.PSVersion.Major -eq 2) { Send-MailMessage -From $MailSender -To $MailAddresses -Subject $SummaryConclusion -BodyAsHtml $SummaryText -Attachments $LogFile -SmtpServer $MailServer -UseSsl:$UseSsl.IsPresent }
        Else { Send-MailMessage -From $MailSender -To $MailAddresses -Subject $SummaryConclusion -BodyAsHtml $SummaryText -Attachments $LogFile -SmtpServer $MailServer -Port $MailPort -UseSsl:$UseSsl.IsPresent }
    }
    #Send Analyze Summary Mail message if enabled by settings
    If($MailAlways -or $MailErrorsAndWarnings -or -$MailErrorsOnly)
    {
        If($MailAlways)
        {
            Send-Mail
        }
        If($MailErrorsAndWarnings)
        {
            #If errors or warnings have been found, send message
            If(($ErrorsAndWarnings -or $SummaryAll) -and (($AmountOfErrors -gt 0) -or ($AmountOfWarnings -gt 0)))
            {
                Send-Mail
            }
            #If errors have been found, send message
            ElseIf(($ErrorsOnly -or $SummaryAllWithoutWarnings) -and ($AmountOfErrors -gt 0))
            {
                Send-Mail
            }
        }
        If($MailErrorsOnly)
        {
            #If errors have been found, send message
            If($AmountOfErrors -gt 0)
            {
                Send-Mail
            }
        }
    }
}
Else { Write-Error -Message "The folder does not exist. Please verify that $Path really exists and you have sufficient permissions" -Category ObjectNotFound -RecommendedAction "Check if the folder name is correct and if you have sufficient permissions to access the folder." -CategoryReason "Folder not found" -ErrorAction Stop }