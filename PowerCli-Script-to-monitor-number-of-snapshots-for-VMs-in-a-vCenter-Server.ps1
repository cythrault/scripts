# The MIT License (MIT)

# Copyright (c) 2015 www.vThinkBeyondVM.com

# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

# Author :- Shivaprasada.M.A
# Email Id :- shivaprasada.m.a@gmail.com, sanshivp@gmail.com

# Snapshot Monitor PowerCli Script 
#
# How to run :-
# ------------------------------
# 1. Replace <VC_Server_IP> with your VC server IP.
# 2. Replace <VC_Username> and <VC_Password> with username and password of your VC. Make sure that both these parameters are enclosed in single quotes.
# 3. Replace <From_Email_ID>, <To_Email_ID> and <CC_Email_ID> with appropriate email IDs.
# 4. Replace <Output_file_name> with the name of the Excel file you want to dump the output, along with the path.
# 5. Edit Subject and body strings as required.
# 6. Replace <SMPT_Server_IP> with SMTP server IP.
# 7. Install VMWare PowerCli software. Open PowerCli command prompt.
# 8. Go to the directory where this file is stored and run the script as ./create_snapshots_info.ps1.
# 9. Output will be stored in the excel sheet specified in step 4.
# 10. Can schedule a task to run this script everyday automatically as this is full automated.

#Connect-VIServer -server <VC_Server_IP> -Username '<VC_Username>' -Password '<VC_Password>'

$From = "vCenter@bnquebec.ca"
$To = "martin.thomas@banq.qc.ca"
$CC = "martin.thomas@banq.qc.ca"
$Attachment = "snapshots.csv"
$Subject = "Liste des snapshots"
$Body = "Liste des snapshots."
$SMTPServer = "smtp.banq.qc.ca"

Function Get-Snapshot-Count-PerVM {
    param($array, [switch]$count)
    begin {
        $hash = @{}
    }
    process {
        $array | %{ $hash[$_] = $hash[$_] + 1 }
        if($count) {
            $hash.GetEnumerator() | ?{$_.value -gt 32} | %{
                New-Object PSObject -Property @{
                    DS_and_VMName = $_.key
                    Count = $_.value
                }
            }
        }
        else {
            $hash.GetEnumerator() | ?{$_.value -gt 0} | %{$_.key}
        }    
    }
}

$vms_list = dir -Recurse -Path vmstores:\ -Include *delta* | select -expand FolderPath | sort

Get-Snapshot-Count-PerVM $vms_list -count | Export-Csv -Path $Attachment -NoTypeInformation

Send-MailMessage -From $From -to $To -cc $CC -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Attachments $Attachment