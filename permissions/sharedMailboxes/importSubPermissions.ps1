#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-ImportSubPermissions-SharedMailboxes
# Import sub permissions
# PowerShell V2
#################################################

# Configure, must be the same as the values used in retrieve permissions
$permissionReference = 'smb'
$permissionDisplayName = 'Shared Mailbox'

# Define source filter for Exchange Online shared mailboxes
# Option 1 (default): use custom attributes
$filterField = 'CustomAttribute2'
$filterValue = 'HelloID Dynamic Shared Mailbox'

# Option 2: use displayName strategy
# $filterField = 'DisplayName'
# $filterValue = 'smb_'

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# PowerShell commands to import
$commands = @(
    "Get-User",
    "Get-EXOMailbox",
    "Get-EXOMailboxPermission",
    "Get-EXORecipientPermission"
)

#region functions
function Resolve-ExchangeOnlineError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorObjectConverted = $httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop

            if ($null -ne $errorObjectConverted.error_description) {
                $httpErrorObj.FriendlyMessage = $errorObjectConverted.error_description
            }
            elseif ($null -ne $errorObjectConverted.error) {
                if ($null -ne $errorObjectConverted.error.message) {
                    $httpErrorObj.FriendlyMessage = $errorObjectConverted.error.message
                    if ($null -ne $errorObjectConverted.error.code) {
                        $httpErrorObj.FriendlyMessage = $httpErrorObj.FriendlyMessage + " Error code: $($errorObjectConverted.error.code)"
                    }
                }
                else {
                    $httpErrorObj.FriendlyMessage = $errorObjectConverted.error
                }
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param()
    try {
        $rawCertificate = [system.convert]::FromBase64String($actionContext.Configuration.AppCertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $actionContext.Configuration.AppCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions

try {
    Write-Information 'Starting Exchange Online permission entitlement import for shared mailboxes'

    $actionMessage = 'importing module [ExchangeOnlineManagement]'

    $importModuleSplatParams = @{
        Name        = 'ExchangeOnlineManagement'
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = 'Stop'
    }

    $null = Import-Module @importModuleSplatParams

    Write-Information "Imported module [$($importModuleSplatParams.Name)]"

    $actionMessage = 'retrieving certificate'
    $certificate = Get-MSEntraCertificate

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
    $actionMessage = 'connecting to Microsoft Exchange Online'

    $createExchangeSessionSplatParams = @{
        Organization          = $actionContext.Configuration.Organization
        AppID                 = $actionContext.Configuration.AppId
        Certificate           = $certificate
        CommandName           = $commands
        ShowBanner            = $false
        ShowProgress          = $false
        TrackPerformance      = $false
        SkipLoadingCmdletHelp = $true
        SkipLoadingFormatData = $true
        ErrorAction           = 'Stop'
    }

    $null = Connect-ExchangeOnline @createExchangeSessionSplatParams

    Write-Information 'Connected to Microsoft Exchange Online'

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailbox?view=exchange-ps
    $actionMessage = "getting user mailboxes from Microsoft Exchange Online"
    $getUserMailboxesParams = @{
        RecipientTypeDetails = 'UserMailbox'
        ResultSize           = 'Unlimited'
        Properties           = 'GrantSendOnBehalfTo'
        ErrorAction          = 'Stop'
    }

    $userMailboxes = Get-EXOMailbox @getUserMailboxesParams | Select-Object Guid, Name, UserPrincipalName, ExternalDirectoryObjectId, GrantSendOnBehalfTo

    $userMailboxesUpnGrouped = $userMailboxes | Group-Object -Property 'UserPrincipalName' -AsHashTable -AsString
    $userMailboxesGuidGrouped = $userMailboxes | Group-Object -Property 'Guid' -AsHashTable -AsString
    $userMailboxesNameGrouped = $userMailboxes | Group-Object -Property 'Name' -AsHashTable -AsString
    Write-Information "Successfully queried [$(@($userMailboxes | Where-Object { $null -ne $_ }).Count)] user mailboxes"

    # Cleanup for memory
    $userMailboxes = $null

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailbox?view=exchange-ps
    $actionMessage = "getting shared mailboxes from Microsoft Exchange Online"
    $getSharedMailboxesParams = @{
        RecipientTypeDetails = 'SharedMailbox'
        ResultSize           = 'Unlimited'
        Properties           = 'GrantSendOnBehalfTo'
        ErrorAction          = 'Stop'
    }

    $sharedMailboxes = Get-EXOMailbox @getSharedMailboxesParams | Select-Object DisplayName, Name, Guid, UserPrincipalName, GrantSendOnBehalfTo

    Write-Information "Successfully queried [$(@($sharedMailboxes | Where-Object { $null -ne $_ }).Count)] shared mailboxes"
    
    # Filter shared mailboxes with the matching filter criteria
    if ($filterField -eq 'DisplayName') {
        $sharedMailboxes = $sharedMailboxes | Where-Object { $_.DisplayName -like "$filterValue*" }
    }
    else {
        $sharedMailboxes = $sharedMailboxes | Where-Object { $_.$filterField -eq $filterValue }
    }
    
    Write-Information "Successfully queried [$(@($sharedMailboxes | Where-Object { $null -ne $_ }).Count)] shared mailboxes matching filter [$filterField = $filterValue]"



    $actionMessage = 'querying mailbox permissions'
    
    foreach ($sharedMailbox in $sharedMailboxes) {
        # Get FullAccess permissions for this shared mailbox
        $getFullAccessPermissionsParams = @{
            Identity    = $sharedMailbox.Guid
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }
        
        $fullAccessUsers = @()
        $fullAccessPermissions = Get-EXOMailboxPermission @getFullAccessPermissionsParams | 
        Where-Object { $_.AccessRights -eq 'FullAccess' -and $_.Deny -eq $false } |
        Select-Object User
        
        foreach ($record in $fullAccessPermissions) {
            $fullAccessUser = $userMailboxesUpnGrouped[$record.User].Guid
            if ($fullAccessUser) { $fullAccessUsers += $fullAccessUser }
        }
        
        $numberOfAccounts = $fullAccessUsers.Count
        $numberOfFullAccess += $numberOfAccounts
        
        $permission = @{
            PermissionReference      = @{
                Reference = $permissionReference
            }
            DisplayName              = "Permission - $permissionDisplayName"
            SubPermissionReference   = @{
                Id = "FullAccess-$($sharedMailbox.Guid)"
            }
            SubPermissionDisplayName = "FullAccess-$($sharedMailbox.DisplayName)"
        }
        
        # Batch permissions based on the amount of account references,
        # to make sure the output objects are not above the limit
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize) }
            foreach ($batch in $batches) {
                # Every chunk is its own object: the same object is not reused for the next chunk
                $chunk = $permission.Clone()
                $chunk.AccountReferences = [array]($batch.Group | ForEach-Object { @($fullAccessUsers[$_]) })
                Write-Output $chunk
            }
        }

        # Get SendAs permissions for this shared mailbox
        $sendAsUsers = @()
        # Query per shared mailbox on the Guid (not on the name): shared mailboxes with the same name must not get each other's permissions
        $actionMessage = "getting SendAs permissions for shared mailbox with id [$($sharedMailbox.Guid)] from Microsoft Exchange Online"
        $getSendAsPermissionsParams = @{
            Identity    = $sharedMailbox.Guid
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }
        $sendAsPermissions = @(Get-EXORecipientPermission @getSendAsPermissionsParams | Where-Object { $_.AccessControlType -eq 'Allow' -and $_.AccessRights -contains 'SendAs' } | Select-Object Trustee)
        if ($sendAsPermissions.Count -gt 0) {
            foreach ($record in $sendAsPermissions) {
                $sendAsUser = $userMailboxesUpnGrouped[$record.Trustee].Guid
                if ($sendAsUser) { $sendAsUsers += $sendAsUser }
            }
        }

        $numberOfAccounts = $sendAsUsers.Count
        $numberOfSendAs += $numberOfAccounts

        $permission = @{
            PermissionReference      = @{
                Reference = $permissionReference
            }
            DisplayName              = "Permission - $permissionDisplayName"
            SubPermissionReference   = @{
                Id = "SendAs-$($sharedMailbox.Guid)"
            }
            SubPermissionDisplayName = "SendAs-$($sharedMailbox.DisplayName)"
        }

        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize) }
            foreach ($batch in $batches) {
                # Every chunk is its own object: the same object is not reused for the next chunk
                $chunk = $permission.Clone()
                $chunk.AccountReferences = [array]($batch.Group | ForEach-Object { @($sendAsUsers[$_]) })
                Write-Output $chunk
            }
        }

        # Get SendOnBehalf permissions for this shared mailbox
        $sendOnBehalfUsers = @()
        if ($null -ne $sharedMailbox.GrantSendOnBehalfTo -and $sharedMailbox.GrantSendOnBehalfTo.Count -gt 0) {
            foreach ($trustee in $sharedMailbox.GrantSendOnBehalfTo) {
                $sendOnBehalfUser = $null
                $trusteeValue = [string]$trustee

                $trusteeMailbox = $userMailboxesUpnGrouped[$trusteeValue]
                if (-not $trusteeMailbox) {
                    $trusteeMailbox = $userMailboxesGuidGrouped[$trusteeValue]
                }
                if (-not $trusteeMailbox) {
                    $trusteeMailbox = $userMailboxesNameGrouped[$trusteeValue]
                }

                if ($trusteeMailbox) {
                    $sendOnBehalfUser = $trusteeMailbox.Guid
                }

                if (-not $sendOnBehalfUser) {
                    # Not silent: a trustee that cannot be resolved is logged, other errors stop the import (no incomplete list)
                    try {
                        $resolvedTrustee = Get-EXOMailbox -Identity $trusteeValue -RecipientTypeDetails UserMailbox -ErrorAction Stop | Select-Object -First 1 -Property Guid
                        if ($resolvedTrustee.Guid) {
                            $sendOnBehalfUser = $resolvedTrustee.Guid
                        }
                    }
                    catch {
                        if ($_.Exception.Message -like "*couldn't be found*" -or $_.Exception.Message -like "*wasn't found*" -or $_.CategoryInfo.Reason -eq 'ManagementObjectNotFoundException') {
                            Write-Warning "SendOnBehalf trustee [$trusteeValue] of shared mailbox [$($sharedMailbox.Guid)] could not be resolved to a user mailbox and is skipped."
                        }
                        else {
                            throw
                        }
                    }
                }

                if ($sendOnBehalfUser) { $sendOnBehalfUsers += $sendOnBehalfUser }
            }
        }

        $numberOfAccounts = $sendOnBehalfUsers.Count
        $numberOfSendOnBehalf += $numberOfAccounts

        $permission = @{
            PermissionReference      = @{
                Reference = $permissionReference
            }
            DisplayName              = "Permission - $permissionDisplayName"
            SubPermissionReference   = @{
                Id = "SendOnBehalf-$($sharedMailbox.Guid)"
            }
            SubPermissionDisplayName = "SendOnBehalf-$($sharedMailbox.DisplayName)"
        }

        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize) }
            foreach ($batch in $batches) {
                # Every chunk is its own object: the same object is not reused for the next chunk
                $chunk = $permission.Clone()
                $chunk.AccountReferences = [array]($batch.Group | ForEach-Object { @($sendOnBehalfUsers[$_]) })
                Write-Output $chunk
            }
        }
    }

    Write-Information "Target permission import for shared mailboxes completed. Full Access: [$numberOfFullAccess] | Send As: [$numberOfSendAs] | Send on Behalf: [$numberOfSendOnBehalf]"
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }

    Write-Warning $warningMessage
    Write-Error $auditMessage
}
finally {
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $actionMessage = 'disconnecting from Microsoft Exchange Online'

    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = 'Stop'
    }

    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams

    Write-Information 'Disconnected from Microsoft Exchange Online'
}
