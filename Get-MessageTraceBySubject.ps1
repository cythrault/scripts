######################################################################################################
#                                                                                                    #
# Name:        Get-MessageTraceBySubject.ps1                                                         #
#                                                                                                    #
# Version:     1.1                                                                                   #
#                                                                                                    #
# Description: Searches the previous X days for a message by subject and exports the results to CSV. #
#                                                                                                    #
# Limitations: Search query is limited to 5,000,000 entries.                                         #
#                                                                                                    #
# Requires:    Remote PowerShell Connection to Exchange Online                                       #
#                                                                                                    #
# Author:      Joe Palarchio                                                                         #
#                                                                                                    #
# Usage:       Additional information on the usage of this script can found at the following         #
#              blog post:  http://blogs.perficient.com/microsoft/?p=31043                            #
#                                                                                                    #
# Disclaimer:  This script is provided AS IS without any support. Please test in a lab environment   #
#              prior to production use.                                                              #
#                                                                                                    #
######################################################################################################


<#
	.PARAMETER  Days
		Number of days back to search.

	.PARAMETER  Subject
		Subject of message to search.

	.PARAMETER  OutputFile
		Name of CSV file to populate with results.
#>

Param(
    [Parameter(Mandatory=$True)]
        [int]$Days,
    [Parameter(Mandatory=$True)]
        [string]$Subject,
    [Parameter(Mandatory=$True)]
        [string]$OutputFile
    )


[DateTime]$DateEnd = Get-Date -format g
[DateTime]$DateStart = $DateEnd.AddDays($Days * -1)

$FoundCount = 0

For($i = 1; $i -le 1000; $i++)  # Maximum allowed pages is 1000
{
    $Messages = Get-MessageTrace -StartDate $DateStart -EndDate $DateEnd -PageSize 5000 -Page $i

    If($Messages.count -gt 0)
    {
        $Status = $Messages[-1].Received.ToString("MM/dd/yyyy HH:mm") + " - " + $Messages[0].Received.ToString("MM/dd/yyyy HH:mm") + "  [" + ("{0:N0}" -f ($i*5000)) + " Searched | " + $FoundCount + " Found]"

        Write-Progress -activity "Checking Messages (Up to 5 Million)..." -status $Status

        $Entries = $Messages | Where {$_.Subject -like $Subject} | Select Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, Size, MessageId
        $Entries | Export-Csv $OutputFile -NoTypeInformation -Append

        $FoundCount += $Entries.Count
    }
    Else
    {
        Break
    }
}  

Write-Host $FoundCount "Entries Found & Logged In" $OutputFile