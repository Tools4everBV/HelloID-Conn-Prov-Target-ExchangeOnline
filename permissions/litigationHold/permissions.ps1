#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-LitigationHold
# List Litigation Hold options as permissions
# PowerShell V2
#################################################

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Litigation Hold - 2555 days"
        Identification = @{
            Id       = "LitigationHold-2555"
            Duration = 2555
        }
    }
)

$outputContext.Permissions.Add(
    @{
        DisplayName    = "Litigation Hold - 365 days"
        Identification = @{
            Id       = "LitigationHold-365"
            Duration = 365
        }
    }
)