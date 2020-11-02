Get-Mailbox -resultsize unlimited | Select-Object name,@{n="Primary Size";e={(Get-MailboxStatistics $_.identity).totalItemsize}},
@{n="Primary Item Count";e={(Get-MailboxStatistics $_.identity).ItemCount}}, 
@{n="Archive Size";e={(Get-MailboxStatistics -archive $_.identity).TotalItemSize}},
@{n="Archive Item Count";e={(Get-MailboxStatistics -archive $_.identity).ItemCount}} | ft 