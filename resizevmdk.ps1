Function Resize-VMDK{
<#
	.Synopsis
	This function adjusts the reduces the size of a vmdk file
	
	.Description
	In the vCenter GUI there is no way to reduce the size of a vmdk. You can increase one by just clicking an up button in the VM setting. Decreasing doesn't work that way though.
	
	This script/function, copies the vmdk header file locally, changes a parameter within vmdk header, and copies the vmdk header file back up to the Datastore.
	
	Further, before the vmdk header file is altered, a copy/backup of it is made. I have successfully used this backup to fix a vm after screwing up the first file. Mileage may vary.
	
	After the command runs a cold storage migration of the VM will still be needed before the affect takes place.
	
	
	.Parameter VM
	This is the VM whose vmdk file will be shrunk
	
	.Parameter NewSizeGB
	The new vmdk file size in GB
	
	.Parameter Local
	The directory locally where the vmdk header file will be copied
	
	
	
	.Example
	
	resize-vmdk -vm testvm01 -local C:\Powershell -newsizGB 45
	
	This will take testvm01 and resize the selected vmdk to be 45GB and will work locally in directory C:\Powershell
	
	
	.Link
	www.vnoob.com
	
	./Notes
	
	Please make sure you have already unpartitioned the space in the guest OS before trying to remove it in vmdk
	

====================================================================
	Author:
	Conrad Ramos <conrad@vnoob.com> http://www.vnoob.com/

	Date:			2012-10-23
	Revision: 		1.0
	
	Disclaimer:
	With this function you can easily corrupt and VM partition by resizing the VMDK for a greater amount than you should. Please be careful, you have been warned, and I am not responsible if you corrupt any/all of your VMs.
	
	
====================================================================


#>



param([Parameter(Mandatory=$true)]$vm, [Parameter(Mandatory=$true)]$local,[Parameter(Mandatory=$true)][int]$newsizeGB)



$disks=get-vm $vm |get-harddisk|select -expand filename
$count=$disks.count

#Convert the New size of the vmdk to the size listed in the vmdk
$size=(($newsizeGB)*([math]::pow(1024,3)) )/512



#Setting up a Multiple Choice decision for the user
#to select the vmdk they want to resize
$i=0
$choicearray=@()

$caption = "Choose Action"
$message = "Which HardDisk are you down-sizing"
IF ($count -gt 1)
{
do {
$choice=$null
$test=$null
[string]$test=$($disks[$($i)])
$choice=new-Object System.Management.Automation.Host.ChoiceDescription "&$($i) $test","Choose which disk"
$choicearray+=$choice
$i++
	

	}	
	while ($i -lt $count)
	$choices = [System.Management.Automation.Host.ChoiceDescription[]]($choicearray)
	$answer = $host.ui.PromptForChoice($caption,$message,$choices,0)
	$disk=$disks[$answer]
}

#If there is only one disk attached to the VM it 
#skips the multiple choice
Else {$disk=$disks}

#From the VMDK name in VMware, regex out the
#relevant information like datastore location etc..
$disk
$ds=($disk -split "\s+")[0]

$ds=$ds.trimend("]").trimstart("[")
$file=($disk -split " ")[1]
$vmdk=$file.split("/")[1]
$RW="RW"

#Create a PSDrive to grab the vmdk from
Remove-PSDrive -Name vDS -erroraction silentlycontinue
new-psdrive -name vDS -location (get-datastore $ds) -psprovider vimdatastore -root '/' |out-null
#Copy VMDK Locally

copy-datastoreitem -item vDS:$file -destination $local

#Create Backup of VMDK
$date=get-date -format yyyyMMddhhmmss
Copy-Item $local/$vmdk $local/"Backup"$date$vmdk -Confirm

#Find the Current size setting and replace it
$string=(get-content $local\$vmdk)
$array=$string -split "\s+"
$num=[array]::indexof($array, $RW)
$string.replace("$($array[$num+1])", "$size")|set-content $local\$vmdk

#Copy the vmdk back up to its rightfull place
Copy-DatastoreItem -item $local\$vmdk -destination vDS:$file -confirm


}

