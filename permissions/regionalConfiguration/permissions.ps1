#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-MailboxRegionalConfiguration
# List Mailbox Regional Configuration options as permissions
# PowerShell V2
#################################################

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Mailbox Regional Configuration - NL"
        Identification = @{
            Id = "MailboxRegionalConfiguration-NL"
        }
    }
)