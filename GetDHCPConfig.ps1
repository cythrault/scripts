Get-DhcpServerv4Scope -ComputerName pvwpdc0d.bnquebec.ca | %{
	$scopeoptions = Get-DhcpServerv4OptionValue -ComputerName pvwpdc0d.bnquebec.ca -ScopeId $PSItem.ScopeID
	[pscustomobject]@{
		ScopeId=$PSItem.ScopeID;
		SubnetMask=$PSItem.SubnetMask;
		Name=$PSItem.Name;
		State=$PSItem.State;
		StartRange=$PSItem.StartRange;
		EndRange=$PSItem.EndRange;
		LeaseDuration=$PSItem.LeaseDuration;
		Router=$scopeoptions.Router.Value;
		"DNS Domain Name"=$scopeoptions."DNS Domain Name".Value;
		"Boot Server Host Name"=$scopeoptions."Boot Server Host Name".Value;
		"Bootfile Name"=$scopeoptions."Bootfile Name".Value;
		Lease=$scopeoptions.Lease.Value;
		"TFTP Server IP Address"=$scopeoptions."TFTP Server IP Address".Value;
		"DNS Servers"=$scopeoptions."DNS Servers".Value;
		"Mitel option"=$scopeoptions."Mitel option".Value;
	}
}