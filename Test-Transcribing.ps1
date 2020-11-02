Function Test-Transcribing
{
<#

.SYNOPSIS
	  This function will test to see if the current system is transcribing.

.DESCRIPTION
          This function will test to see if the current system is transcribing, the current transcript will be stopped and restarted with information added to the transcript to show that the log was tested, then reutrn a boolean value.
 
.INPUTS
	None

.OUTPUTS
	Boolean

.NOTES
	NAME:	Test-Transcribing.ps1
	AUTHOR:	Darryl Kegg
	DATE:	01 October, 2015
	EMAIL:	dkegg@microsoft.com

	VERSION HISTORY:
	1.0 01 October, 2015    Initial Version
	

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>

[CmdletBinding(SupportsShouldProcess=$True)]
param()

Write-Verbose "Running function TEST-TRANSCRIBING...."
Write-Verbose "Set Boolean value to false by default"
$IsTranscribing = $false

Write-Information "Testing to see if powershell is transcribing.  If so, we will stop and re-start transcription"

Write-Verbose "Now we test to see if transcribing is in progress" 
$stopTest = try {Stop-transcript -ErrorAction stop} catch {}                                                      
 
if (!$stopTest) {write-Verbose "No Transcription was started, we do nothing."}                                     

if ($stopTest -and $stoptest.Contains("not been started")) {write-Verbose "No Transcription was started, we do nothing."}                                     
 
if ($stopTest -and $stoptest.Contains("output file"))
{
	Write-Verbose "A running transcript was found, resuming..."
	Start-Transcript -path $stoptest.Split(" ")[$stoptest.Split(" ").count-1] -append  | out-null
	Write-Information "Stopped and restarted the transcription as part of the TEST-TRANSCRIBING function"                             
	$IsTranscribing = $True
}                              

Write-Verbose "Returning the value of $IsTranscribing to the calling script" 
Return $IsTranscribing
}