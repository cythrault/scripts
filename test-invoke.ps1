$PSCommandPath
(Get-Item $PSCommandPath ).Extension
(Get-Item $PSCommandPath ).Basename
(Get-Item $PSCommandPath ).Name
(Get-Item $PSCommandPath ).DirectoryName
(Get-Item $PSCommandPath ).FullName

[string]$logPath = $(Get-Date -UFormat %Y%m%d).log

$logfile = (Get-Item $PSCommandPath ).DirectoryName + "\" +  (Get-Item $PSCommandPath ).Basename + "-" +  $(Get-Date -UFormat %Y%m%d) + ".log"
$logfile