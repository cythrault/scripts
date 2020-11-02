function creds {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$user
        )
    $AESKeyFilePath = "$PSScriptRoot\AESKey.key"
    $credsFile = $PSScriptRoot + "\" + ($user -Replace '.*\\') + ".aes"
	If (![System.IO.File]::Exists($AESKeyFilePath)) {
		$AESKey = New-Object Byte[] 32
		[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
		Set-Content $AESKeyFilePath $AESKey -Verbose
	} Else { $AESKey = Get-Content $AESKeyFilePath }
	If (![System.IO.File]::Exists($credsFile)) {
		Set-Content $credsFile $($(Get-Credential -Username $user -Message "Messsage").Password | ConvertFrom-SecureString -Key $AESKey)
		}
	return (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, (Get-Content $credsFile | ConvertTo-SecureString -Key $AESKey))
}

function New-Chart{
 param(
   [Parameter(Mandatory=$true)][array]$List,
   [Parameter(Mandatory=$true)][array]$Data,
   [string]$Title = " ",
   [Boolean]$Explode = $true,
   [int]$Width = 1024,
   [int]$Height = 768,
   [Boolean]$Pie = $false,
   [string]$ImageFile = $PSScriptRoot + "\chart.png",
   #For use in bar charts
   [string]$AxisXTitle = " ",
   [string]$AxisYTitle = " "
 )
 
  [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
  $chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
  $chart.Width = $Width
  $chart.Height = $Height
  # $chart.BackColor = [System.Drawing.Color]::Transparent
  [void]$chart.Titles.Add($Title)
  $chart.Titles[0].Font = "arial,48pt"
  $chart.Titles[0].Alignment = "topCenter"
  $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
  $chartarea.Name = "ChartArea1"
  $chart.ChartAreas.Add($chartarea)
  [void]$chart.Series.Add("data1")
  if($Pie)
  {
    $chart.Series["data1"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
    $chart.Series["data1"].Points.DataBindXY($List, [int[]]$Data)
    $Chart.Series["data1"]["PieLabelStyle"] = "Outside"
    $Chart.Series["data1"]["PieLineColor"] = "Black"
    $Chart.Series["data1"]["PieDrawingStyle"] = "Concave"
    if($Explode)
    {
      ($Chart.Series["data1"].Points.FindMaxByValue())["Exploded"] = $true
    }
  }
  else
  {
    $chartarea.AxisX.Interval = 1
    $ChartArea.AxisX.LabelStyle.Font = "arial,10pt"
    $ChartArea.AxisY.LabelStyle.Font = "arial,10pt"
    $ChartArea.AxisX.Title = $AxisXTitle
    $ChartArea.AxisY.Title = $AxisYTitle
    $chart.Series["data1"].Points.DataBindXY($List, [int[]]$Data)
    $maxValue = $Chart.Series["data1"].Points.FindMaxByValue()
    $maxValue.Color = [System.Drawing.Color]::Red
    $minValue = $Chart.Series["data1"].Points.FindMinByValue()
    $minValue.Color = [System.Drawing.Color]::Green
    #$Chart.Series["data1"]["DrawingStyle"] = "Cylinder"
  }
  $chart.SaveImage($ImageFile,"png")
}

If (!$swis) {
    Write-Host "Establishing connection to Orion"
    $hostname = "orion.wspgroup.com"
    $swis = Connect-Swis -Credential $(creds -User "corp\camt055545it9") -Hostname $hostname
    }

<#

$nodeID = $(Get-SwisData $swis "SELECT NodeID FROM Orion.Nodes WHERE SysName like 'MTRLPQGZ-GENIVA01%'");
$ifID = $(Get-SwisData $swis "SELECT top 1 InterfaceID FROM Orion.NPM.Interfaces WHERE NodeID = $nodeID")
$PolicyID = $(Get-SwisData $swis "SELECT PolicyID FROM Orion.Netflow.CBQoSPolicy WHERE NodeID = $nodeID and interfaceid = $ifID and PolicyFullPathName like '%\LAN_EF%'")

$bitrates = @()
$bitrates = (Get-SwisData $swis "SELECT timestamp,Bitrate FROM Orion.Netflow.CBQoSStatistics WHERE policyid = $PolicyID and timestamp > '$((Get-Date).AddDays(-7).ToUniversalTime().ToString("yyyy-MM-dd h:mm:ss tt"))'")

New-Chart -Title "Bitrates" -List $bitrates.Timestamp -Data $bitrates.Bitrate -AxisXTitle "Time" -AxisYTitle "Bitrate"

#>

