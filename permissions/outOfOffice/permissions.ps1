#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-OutOfOffice
# List Out of Office as permission
# PowerShell V2
#################################################

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Out Of Office reply"
        Identification = @{
            Id = "OutOfOffice"
        }
    }
)
