#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-folderPermission
# List mailbox folder permissions as permissions
# See Microsoft Docs for supported params https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailboxfolderpermission?view=exchange-ps
# PowerShell V2
#################################################

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Mailbox folder permissions - Availability only"
        Identification = @{
            Id                       = '1' # Must be unique
            mailboxFolderUser        = 'Default'
            mailboxFolderAccessRight = 'AvailabilityOnly' # Options: AvailabilityOnly, LimitedDetails, Reviewer, Editor
        }
    }
)

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Mailbox folder permissions - Limited details"
        Identification = @{
            Id                       = '2' # Must be unique
            mailboxFolderUser        = 'Default'
            mailboxFolderAccessRight = 'LimitedDetails' # Options: AvailabilityOnly, LimitedDetails, Reviewer, Editor
        }
    }
)

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Mailbox folder permissions - Reviewer"
        Identification = @{
            Id                       = '3' # Must be unique
            mailboxFolderUser        = 'Default'
            mailboxFolderAccessRight = 'Reviewer' # Options: AvailabilityOnly, LimitedDetails, Reviewer, Editor
        }
    }
)

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Mailbox folder permissions - Editor"
        Identification = @{
            Id                       = '4' # Must be unique
            mailboxFolderUser        = 'Default'
            mailboxFolderAccessRight = 'Editor' # Options: AvailabilityOnly, LimitedDetails, Reviewer, Editor
        }
    }
)