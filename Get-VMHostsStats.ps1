$allhosts = @()
$hosts = Get-VMHost

foreach($vmHost in $hosts){
  $hoststat = "" | Select HostName, MemoryInstalled, MemoryAllocated, MemoryConsumed, MemoryUsage, CPUMax, CPUAvg, CPUMin
  $hoststat.HostName = $vmHost.name
  
  $statcpu = Get-Stat -Entity ($vmHost)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat cpu.usage.average
  $statmemconsumed = Get-Stat -Entity ($vmHost)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat mem.consumed.average
  $statmemusage = Get-Stat -Entity ($vmHost)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat mem.usage.average
  $statmemallocated = Get-VMhost $vmHost.name | Select @{N="allocated";E={$_ | Get-VM | %{$_.MemoryGB} | Measure-Object -Sum | Select -ExpandProperty Sum}}
  $statmeminstalled = Get-VMHost $vmHost.name | select MemoryTotalGB
  $statmeminstalled = $statmeminstalled.MemoryTotalGB

  $cpu = $statcpu | Measure-Object -Property value -Average -Maximum -Minimum
  $memconsumed = $statmemconsumed | Measure-Object -Property value -Average
  $memusage = $statmemusage | Measure-Object -Property value -Average
  
  $CPUMax = "{0:N0}" -f ($cpu.Maximum)
  $CPUAvg = "{0:N0}" -f ($cpu.Average)
  $CPUMin = "{0:N0}" -f ($cpu.Minimum)
  $allocated = "{0:N0}" -f ($statmemallocated.allocated)
  $consumed = "{0:N0}" -f ($memconsumed.Average/1024/1024)
  $usage = "{0:P0}" -f ($memusage.Average/100)
  $installed = "{0:N0}" -f ($statmeminstalled)

  $CPUMax = $CPUMax.ToString() #+ " %"
  $CPUAvg = $CPUAvg.ToString() #+ " %"
  $CPUMin = $CPUMin.ToString() #+ " %"
  $MemoryInstalled = $installed.ToString() #+ " GB"
  $MemoryAllocated = $allocated.ToString() #+ " GB"
  $MemoryConsumed = $consumed.ToString() #+ " GB"
  $MemoryUsage = $usage.ToString()

  $hoststat.CPUMax = $CPUMax
  $hoststat.CPUAvg = $CPUAvg
  $hoststat.CPUMin = $CPUMin
  $hoststat.MemoryInstalled = $MemoryInstalled
  $hoststat.MemoryAllocated = $MemoryAllocated
  $hoststat.MemoryConsumed = $MemoryConsumed
  $hoststat.MemoryUsage = $MemoryUsage
  $allhosts += $hoststat
}
$allhosts | Select HostName, MemoryInstalled, MemoryAllocated, MemoryConsumed, MemoryUsage, CPUMax, CPUAvg, CPUMin