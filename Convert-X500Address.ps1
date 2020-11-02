<#
Convert-X500Address.ps1

Convert IMCEAEX string from NDR message to X500 Address format. This
Script simply displays the X500 string. Copy it and make a new 
X.500 Email address to the Exchange object.

Parameter: Pass the IMCEAEX string from NDR message in double quotes

Written By: Anand, the Awesome, Venkatachalapathy

#>
param($IMCEAEXString)

((((((($IMCEAEXString.Replace("IMCEAEX-","")).Replace("_","/")).Replace("+20"," ")).Replace("+28","(")).Replace("+29",")")).Replace("+2E",".").Replace("+2C",",")).Replace("+5F","_"))

#* * * End of the Script * * *