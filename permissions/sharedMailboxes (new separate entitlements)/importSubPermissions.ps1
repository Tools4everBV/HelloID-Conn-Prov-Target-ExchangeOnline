#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-ImportSubPermissions-SharedMailboxes
# Import sub permissions
# PowerShell V2
#################################################

# Configure dynamically based on permission data returned from EXO
$permissionMetadataByType = @{
    FullAccess   = @{ Reference = 'FullAccess'; DisplayName = 'Full Access Mailbox' }
    SendAs       = @{ Reference = 'SendAs'; DisplayName = 'Send As Mailbox' }
    SendOnBehalf = @{ Reference = 'SendOnBehalf'; DisplayName = 'Send On Behalf Mailbox' }
}

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
    'Get-Mailbox'
    , 'Get-EXOMailboxPermission'
    , 'Get-EXORecipientPermission'
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
            $errorObjectConverted = $ErrorObject | ConvertFrom-Json -ErrorAction Stop

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
                $httpErrorObj.FriendlyMessage = $ErrorObject
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
    Write-Information 'Starting Exchange Online permission entitlement import for shared mailboxes (separate entitlements)'

    $actionMessage = 'importing module [ExchangeOnlineManagement]'

    $importModuleSplatParams = @{
        Name        = 'ExchangeOnlineManagement'
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = 'Stop'
    }

    $null = Import-Module @importModuleSplatParams

    Write-Information "Imported module [$($importModuleSplatParams.Name)]"

    if ($actionContext.Configuration.UseCertificate -eq $true) {
        Write-Information 'Connecting to Exchange Online with certificate'

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
    }
    else {
        Write-Information 'Connecting to Exchange Online with secret'

        $actionMessage = 'creating access token'

        $createAccessTokenBody = @{
            grant_type    = 'client_credentials'
            client_id     = $actionContext.Configuration.AppId
            client_secret = $actionContext.Configuration.AppSecret
            resource      = 'https://outlook.office365.com'
        }

        $createAccessTokenSplatParams = @{
            Uri             = "https://login.microsoftonline.com/$($actionContext.Configuration.TenantID)/oauth2/token"
            Headers         = $headers
            Method          = 'POST'
            ContentType     = 'application/x-www-form-urlencoded'
            UseBasicParsing = $true
            Body            = $createAccessTokenBody
            Verbose         = $false
            ErrorAction     = 'Stop'
        }

        $createAccessTokenResponse = Invoke-RestMethod @createAccessTokenSplatParams

        Write-Information 'Created access token.'

        # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
        $actionMessage = 'connecting to Microsoft Exchange Online'

        $createExchangeSessionSplatParams = @{
            Organization          = $actionContext.Configuration.Organization
            AppID                 = $actionContext.Configuration.AppId
            AccessToken           = $createAccessTokenResponse.access_token
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
    }

    $actionMessage = 'getting all mailboxes from Microsoft Exchange Online'
    
    $getAllMailboxesParams = @{
        ResultSize  = 'Unlimited'
        ErrorAction = 'Stop'
    }
    
    $mailboxes = Get-Mailbox @getAllMailboxesParams
    $userMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' } | Select-Object Guid, Name, UserPrincipalName, ExternalDirectoryObjectId, GrantSendOnBehalfTo
    $userMailboxesUpnGrouped = $userMailboxes | Group-Object -Property 'UserPrincipalName' -AsHashTable -AsString
    $userMailboxesGuidGrouped = $userMailboxes | Group-Object -Property 'Guid' -AsHashTable -AsString
    $userMailboxesNameGrouped = $userMailboxes | Group-Object -Property 'Name' -AsHashTable -AsString
    Write-Information "Successfully queried [$($userMailboxes.count)] user mailboxes"
    $sharedMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' } | Select-Object DisplayName, Name, Guid, UserPrincipalName, GrantSendOnBehalfTo
    
    # Filter shared mailboxes with the matching filter criteria
    if ($filterField -eq 'DisplayName') {
        $sharedMailboxes = $sharedMailboxes | Where-Object { $_.DisplayName -like "$filterValue*" }
    }
    else {
        # Get mailbox details including custom attributes
        $allSharedMailboxesDetailed = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' } | Select-Object DisplayName, Name, Guid, UserPrincipalName, GrantSendOnBehalfTo, CustomAttribute2
        $sharedMailboxes = $allSharedMailboxesDetailed | Where-Object { $_.$filterField -eq $filterValue } | Select-Object DisplayName, Name, Guid, UserPrincipalName, GrantSendOnBehalfTo
    }
    
    Write-Information "Successfully queried [$($sharedMailboxes.count)] shared mailboxes matching filter [$filterField = $filterValue]"
    # Cleanup for memory
    $userMailboxes = $null
    $mailboxes = $null

    # Query all Send As permissions once
    $actionMessage = "getting all recipient permissions from Microsoft Exchange Online"
    $getAllRecipientPermissionsParams = @{
        ResultSize   = 'Unlimited'
        AccessRights = 'SendAs'
        ErrorAction  = 'Stop'
    }
    $allSendAsPermissions = Get-EXORecipientPermission @getAllRecipientPermissionsParams | Where-Object { $_.AccessControlType -eq 'Allow' } | Select-Object Identity, Trustee
    $allSendAsPermissionsGrouped = $allSendAsPermissions | Group-Object -Property 'Identity' -AsHashTable -AsString
    Write-Information "Successfully queried [$($allSendAsPermissions.count)] recipient permissions (SendAs)"
    # Cleanup for memory
    $allSendAsPermissions = $null

    $actionMessage = 'querying mailbox permissions'
    
    foreach ($sharedMailbox in $sharedMailboxes) {
        # Make sure the displayname has a value of max 100 char
        if (-not([string]::IsNullOrEmpty($sharedMailbox.DisplayName))) {
            $displayName = $sharedMailbox.DisplayName.Substring(0, [System.Math]::Min(100, $sharedMailbox.DisplayName.Length))
        }
        else {
            $displayName = $sharedMailbox.Guid
        }

        #region Full Access Permission
        $fullAccessMetadata = $permissionMetadataByType['FullAccess']
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
            $fullAccessUser = $userMailboxesUpnGrouped[$record.User].ExternalDirectoryObjectId
            if ($fullAccessUser) { $fullAccessUsers += $fullAccessUser }
        }
        
        $numberOfAccounts = $fullAccessUsers.Count
        $permission = @{
            PermissionReference      = @{
                Reference = $fullAccessMetadata.Reference
            }
            DisplayName              = "Permission - $($fullAccessMetadata.DisplayName)"
            SubPermissionReference   = @{
                Id         = $sharedMailbox.Guid
                Permission = 'FullAccess'
            }
            SubPermissionDisplayName = "$displayName - Full Access"
        }
        
        # Batch permissions based on the amount of account references
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($fullAccessUsers[$_]) })
                Write-Output $permission
            }
        }
        #endregion Full Access Permission

        #region Send As Permission
        $sendAsMetadata = $permissionMetadataByType['SendAs']
        $sendAsUsers = @()
        $sendAsPermissions = $allSendAsPermissionsGrouped[$sharedMailbox.Name]
        if ($null -ne $sendAsPermissions) {
            foreach ($record in $sendAsPermissions) {
                $sendAsUser = $userMailboxesUpnGrouped[$record.Trustee].ExternalDirectoryObjectId
                if ($sendAsUser) { $sendAsUsers += $sendAsUser }
            }
        }
        
        $numberOfAccounts = $sendAsUsers.Count
        $permission = @{
            PermissionReference      = @{
                Reference = $sendAsMetadata.Reference
            }
            DisplayName              = "Permission - $($sendAsMetadata.DisplayName)"
            SubPermissionReference   = @{
                Id         = $sharedMailbox.Guid
                Permission = 'SendAs'
            }
            SubPermissionDisplayName = "$displayName - Send As"
        }
        
        # Batch permissions based on the amount of account references
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($sendAsUsers[$_]) })
                Write-Output $permission
            }
        }
        #endregion Send As Permission

        #region Send On Behalf Permission
        $sendOnBehalfMetadata = $permissionMetadataByType['SendOnBehalf']
        $sendOnBehalfUsers = @()
        if ($null -ne $sharedMailbox.GrantSendOnBehalfTo -and $sharedMailbox.GrantSendOnBehalfTo.Count -gt 0) {
            foreach ($trustee in $sharedMailbox.GrantSendOnBehalfTo) {
                $sendOnBehalfUser = $null
                $trusteeValue = [string]$trustee

                # GrantSendOnBehalfTo can contain different identity formats (UPN, GUID, Name)
                $trusteeMailbox = $userMailboxesUpnGrouped[$trusteeValue]
                if (-not $trusteeMailbox) {
                    $trusteeMailbox = $userMailboxesGuidGrouped[$trusteeValue]
                }
                if (-not $trusteeMailbox) {
                    $trusteeMailbox = $userMailboxesNameGrouped[$trusteeValue]
                }

                if ($trusteeMailbox) {
                    $sendOnBehalfUser = $trusteeMailbox.ExternalDirectoryObjectId
                }

                if (-not $sendOnBehalfUser) {
                    $resolvedTrustee = Get-Mailbox -Identity $trusteeValue -ErrorAction SilentlyContinue | Select-Object -First 1 -Property ExternalDirectoryObjectId
                    if ($resolvedTrustee.ExternalDirectoryObjectId) {
                        $sendOnBehalfUser = $resolvedTrustee.ExternalDirectoryObjectId
                    }
                }

                if ($sendOnBehalfUser) { $sendOnBehalfUsers += $sendOnBehalfUser }
            }
        }
        
        $numberOfAccounts = $sendOnBehalfUsers.Count
        $permission = @{
            PermissionReference      = @{
                Reference = $sendOnBehalfMetadata.Reference
            }
            DisplayName              = "Permission - $($sendOnBehalfMetadata.DisplayName)"
            SubPermissionReference   = @{
                Id         = $sharedMailbox.Guid
                Permission = 'SendOnBehalf'
            }
            SubPermissionDisplayName = "$displayName - Send On Behalf"
        }
        
        # Batch permissions based on the amount of account references
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($sendOnBehalfUsers[$_]) })
                Write-Output $permission
            }
        }
        #endregion Send On Behalf Permission
    }

    Write-Information 'Exchange Online shared mailbox permission entitlement import completed (separate entitlements)'
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
