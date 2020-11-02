# PowerShell Bits Transfer
# Ashley McGlone
# November, 2010
#
# Original PowerShell & BITS article by Jeffrey Snover here:
#   http://blogs.msdn.com/b/powershell/archive/2009/01/11/transferring-large-files-using-bits.aspx
# Mr. Snover's article was based on pre-release code.
# This version uses the updated RTM syntax.
# This version does a checking loop to see if the file has completed transfer
# yet, and automatically completes the transfer if it is done.
# It also displays progress of the download.

$ErrorActionPreference = "Stop"

Import-Module BitsTransfer
# URL to your large download file here
# EXAMPLE: Windows Server 2008 R2 SP1 Release Candidate
$src  = "\\nz-pdc-win01.opus.global\Martin$\test.fil"
# Local file path here
$dest = "c:\ops"

# We make a unique display name for the transfer so that it can be uniquely
# referenced by name and will not return an array of jobs if we have run the
# script multiple times simultaneously.
$displayName = "MyBitsTransfer " + (Get-Date)
Start-BitsTransfer `
    -Source $src `
    -Destination $dest `
    -DisplayName $displayName `
    -Asynchronous
$job = Get-BitsTransfer $displayName
$CreationTime = $job.CreationTime

# Create a holding pattern while we wait for the connection to be established
# and the transfer to actually begin.  Otherwise the next Do...While loop may
# exit before the transfer even starts.  Print the job status as it changes
# from Queued to Connecting to Transferring.
# If the download fails, remove the bad job and exit the loop.
$lastStatus = $job.JobState
Do {
    If ($lastStatus -ne $job.JobState) {
        $lastStatus = $job.JobState
        $job
    }
    If ($lastStatus -like "*Error*") {
        Remove-BitsTransfer $job
        Write-Host "Error connecting to download."
        Exit
    }
}
while ($lastStatus -ne "Transferring")
$job

$period = 10
$TotalBytes = $job.BytesTotal
$KBTotal = $TotalBytes / 1024

# Print the transfer status as we go
do {

$LastBytesDone = $BytesDone
$BytesDone = $job.BytesTransferred
$KBDone = $BytesDone / 1024
$BytesRemain = $TotalBytes - $BytesDone

$PctDone = $([math]::Round( ( ( $BytesDone / $TotalBytes ) * 100 ), 1))
$done = $([math]::Round(((get-date) - $CreationTime).TotalMinutes))

$bps = $BytesDone / ((get-date) - $CreationTime).TotalSeconds
$kbps = $([math]::Round($bps / 1024))

$OVbps = ( $BytesDone - $LastBytesDone ) / $period
$OVkbps = $([math]::Round( $OVbps / 1024 ))

$ETA = $([math]::Round(($BytesRemain / $bps) / $period))

Write-Progress -Activity "Transferring $src to $dest" `
	-Status $job.JobState `
	-PercentComplete $PctDone `
	-CurrentOperation "Downloaded $KBDone KB of $KBTotal KB - $PctDone% @ $kbps KB/s (Last $period seconds: $OVkbps KB/s) - ETA: $ETA Minutes - Running for $done Minutes (Total $($ETA+$done) Minutes)"
Start-Sleep -s $period
}
while ($BytesDone -lt $TotalBytes)

# Print the final status once complete.
Complete-BitsTransfer $job
Write-Host (Get-Date) "-" $($job.JobState) "- Total transfer time: $([math]::Round( ($job.TransferCompletionTime - $CreationTime).TotalMinutes )) Minutes @ $kbps KB/s"
