function Test {
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		[Alias("patate poil")]
		[string]$AtosID
		)
	Begin {	$a = $MyInvocation.MyCommand.Parameters['AtosID'].Aliases[0]; $a }
	Process {
		$AtosID = $AtosID.Trim()
		$col2=$AtosID + "ITG"
		[pscustomobject]@{$a=$AtosID;"deuxieme colonne"=$col2;"bof"="statique"}
	}
}