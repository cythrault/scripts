function Get-MailboxCalendarPermission {
    <#
    .SYNOPSIS
    Retrieves a list of mailbox calendar permissions
    .DESCRIPTION
    Get a list of permissions for mailbox calendars in an exchange environment. As different languages spell 
    'calendar' differently this script first pulls the actual name of the calendar by using
    get-mailboxfolderstatistics and has proven to work across multi-lingual organizations.
    .PARAMETER MailboxName
    One mailbox name in string format.
    .PARAMETER MailboxName
    Array of mailbox names in string format.    
    .PARAMETER MailboxObject
    One or more mailbox objects.
    .PARAMETER o365
    Required if running this function in office 365.
    .LINK
    http://www.the-little-things.net   
    .NOTES
    Last edit   : 11/04/2014
    Version     :
        2.0.1 11/18/2014
            - Added the o365 flag as the folderscope option is about 4x as efficient so for on prem we want to use it instead
            - Fixed onprem calendar detection to not include folders of type 'user created'
        2.0.1 11/15/2014
            - Removed the strongly typed definition for the mailbox object to increase compatibility with o365
            - Removed the folderscope option for the get-mailboxfolderstatistics call to increase o365 compat.
        2.0.0 11/04/2014
            - Several coding structure changes and clean up.
        1.2.1 09/21/2014
            - Removed log to file options and fixed a bunch of other issues.
            - Enabled input of mailbox objects as well as strings
        1.2.0 May 6 2013    :   Fixed issue where a mailbox name produces more than one mailbox
        1.1.0 April 24 2013 :   Used new script template from http://blog.bjornhouben.com
        1.0.0 March 10 2013 :   Created script
    Author      :   Zachary Loeber
    .EXAMPLE
    Get-MailboxCalendarPermission -MailboxName "Test User1" -Verbose
    Description
    -----------
    Gets the calendar permissions for "Test User1" and shows verbose information.
    .EXAMPLE
    Get-MailboxCalendarPermission -MailboxName "user1","user2" | Format-List
    Description
    -----------
    Gets the calendar permissions for "user1" and "user2" and returns the info as a format-list.
    .EXAMPLE
    (Get-Mailbox -Database "MDB1") | Get-MailboxCalendarPermission | Format-Table Mailbox,User,Permission
    Description
    -----------
    Gets all mailboxes in the MDB1 database and pipes it to Get-MailboxCalendarPermission. Get-MailboxCalendarPermission and returns the info as an autosized format-table containing the Mailbox,User, and Permission
    #>
    [CmdLetBinding(DefaultParameterSetName='AsString')]
    param(
        [Parameter(ParameterSetName='AsStringArray', Mandatory=$True, ValueFromPipeline=$True, Position=0, HelpMessage="Enter an Exchange mailbox name")]
        [string[]]$MailboxNames,
        [Parameter(ParameterSetName='AsString', Mandatory=$True, ValueFromPipeline=$True, Position=0, HelpMessage="Enter an Exchange mailbox name")]
        [string]$MailboxName,
        [Parameter(ParameterSetName='AsMailbox', Mandatory=$True, ValueFromPipeline=$True, Position=0, HelpMessage="Enter an Exchange mailbox name")]
        $MailboxObject,
        [Parameter(HelpMessage='If running in office 365 set this switch')]
        [switch]$o365
    )
    begin {
        Write-Verbose "$($MyInvocation.MyCommand): Begin"
        $Mailboxes = @()
    }
    process {
        switch ($PSCmdlet.ParameterSetName) {
            'AsStringArray' {
                try {
                    $Mailboxes = @($MailboxNames | Foreach {Get-Mailbox $_ -erroraction Stop})
                }
                catch {
                    Write-Warning = "$($MyInvocation.MyCommand): $($_.Exception.Message)"
                }
            }
            'AsString' {
                try { 
                    $Mailboxes = @(Get-Mailbox $MailboxName -erroraction Stop)
                }
                catch {
                    Write-Warning = "$($MyInvocation.MyCommand): $($_.Exception.Message)"
                }
            }
            'AsMailbox' {
               $Mailboxes = @($MailboxObject)
            }
        }

        foreach ($Mailbox in $Mailboxes) {
            Write-Verbose "$($MyInvocation.MyCommand): Checking $($Mailbox.Name)"
            
            # Construct the full path to the calendar folder regardless of the language
            if ($o365) {
                $Calfolder = $Mailbox.Name + ':\' + [string](Get-MailboxFolderStatistics $Mailbox.Identity | 
                Where-Object {$_.FolderType -eq 'calendar'}).Name
            }
            else {
                $Calfolder = $Mailbox.Name + ':\' + [string](Get-MailboxFolderStatistics $Mailbox.Identity -folderscope calendar | 
                Where-Object {$_.FolderType -ne 'User Created'}).Name
            }
            
            # Get the permissions on the calendar
            try {
                $CalPerm = Get-MailboxFolderPermission $Calfolder
            }
            catch {
                $CalPerm = $null
                Write-Warning "$($MyInvocation.MyCommand): $($_.Exception.Message)"
            }
            foreach ($Perm in $CalPerm) {
                New-Object PSObject -Property @{
                                      'Mailbox'=$Mailbox.Name
                                      'User'=$Perm.User
                                      'Calendar' = $Calfolder
                                      'Permission'=$Perm.AccessRights
                                    }
            }
        }
    }
    end {
        Write-Verbose "$($MyInvocation.MyCommand): End"
    }
}