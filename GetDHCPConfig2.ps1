$DHCPServers = (Get-DhcpServerInDC).DnsName

foreach ($DHCPServer in $DHCPServers) {
	if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $dhcpserver -Quiet) {
		$ErrorActionPreference = "silentlycontinue"
		$Scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer

		foreach ($Scope in $Scopes) {
			$scopeoptions += Get-DHCPServerv4OptionValue -ComputerName $DHCPServer -ScopeID $Scope.ScopeId |
				select @{label="DHCPServer"; Expression= {$DHCPServer}},
				@{label="ScopeID"; Expression= {$Scope.ScopeId}},
				@{label="ScopeName"; Expression= {$Scope.Name}},
				@{label="SubnetMask"; Expression= {$Scope.SubnetMask}},
				@{label="State"; Expression= {$Scope.State}},
				@{label="StartRange"; Expression= {$Scope.StartRange}},
				@{label="EndRange"; Expression= {$Scope.EndRange}},
				@{label="LeaseDuration"; Expression= {$Scope.LeaseDuration}},
				*
		}
		$ErrorActionPreference = "continue"
	}
}

$scopeoptions