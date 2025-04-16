#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-SharedMailboxes-Import
# Correlate to permission
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

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
        # Write-Output $httpErrorObj
        return $httpErrorObj
    }
}
#endregion functions

try {
    Write-Information 'Starting target shared mailbox permissions import'
    $actionMessage = "importing module [ExchangeOnlineManagement]"
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = 'Get-User,Get-Mailbox,Get-MailboxPermission,Get-RecipientPermission'
        Verbose     = $false
        ErrorAction = "Stop"
    }
    $null = Import-Module @importModuleSplatParams
    Write-Information "Imported module [$($importModuleSplatParams.Name)]"

    $actionMessage = "creating access token"
    $createAccessTokenBody = @{
        grant_type    = "client_credentials"
        client_id     = $actionContext.Configuration.AppId
        client_secret = $actionContext.Configuration.AppSecret
        resource      = "https://outlook.office365.com"
    }
    $createAccessTokenSplatParams = @{
        Uri             = "https://login.microsoftonline.com/$($actionContext.Configuration.TenantID)/oauth2/token"
        Headers         = $headers
        Method          = "POST"
        ContentType     = "application/x-www-form-urlencoded"
        UseBasicParsing = $true
        Body            = $createAccessTokenBody
        Verbose         = $false
        ErrorAction     = "Stop"
    }
    $createAccessTokenResonse = Invoke-RestMethod @createAccessTokenSplatParams
    Write-Information "Created access token"

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
    $actionMessage = "connecting to Microsoft Exchange Online"
    $createExchangeSessionSplatParams = @{
        Organization          = $actionContext.Configuration.Organization
        AppID                 = $actionContext.Configuration.AppId
        AccessToken           = $createAccessTokenResonse.access_token
        CommandName           = $commands
        ShowBanner            = $false
        ShowProgress          = $false
        TrackPerformance      = $false
        SkipLoadingCmdletHelp = $true
        SkipLoadingFormatData = $true
        ErrorAction           = "Stop"
    }
    $null = Connect-ExchangeOnline @createExchangeSessionSplatParams
    Write-Information "Connected to Microsoft Exchange Online"

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailbox?view=exchange-ps
    $actionMessage = "getting all mailboxes from Microsoft Exchange Online"
    $getAllMailboxesParams = @{
        ResultSize  = 'Unlimited'
        ErrorAction = 'Stop'
    }
    $mailboxes = Get-Mailbox @getAllMailboxesParams
    $userMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' } | Select-Object Guid, UserPrincipalName, ExternalDirectoryObjectId, GrantSendOnBehalfTo
    $userMailboxesUpnGrouped = $userMailboxes | Group-Object -Property 'UserPrincipalName' -AsHashTable -AsString
    $userMailboxesExtDirObIdGrouped = $userMailboxes | Group-Object -Property 'ExternalDirectoryObjectId' -AsHashTable -AsString
    Write-Information "Successfully queried [$($userMailboxes.count)] user mailboxes"
    $sharedMailboxes = $mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' } | Select-Object DisplayName, Name, Guid, UserPrincipalName, GrantSendOnBehalfTo
    Write-Information "Successfully queried [$($sharedMailboxes.count)] shared mailboxes"
    # Cleanup for memory
    $userMailboxes = $null
    $mailboxes = $null

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-recipientpermission?view=exchange-ps
    $actionMessage = "getting all recipient permissions from Microsoft Exchange Online"
    $getAllRecipientPermissionsParams = @{
        ResultSize  = 'Unlimited'
        ErrorAction = 'Stop'
    }
    $allRecipientPermissions = Get-RecipientPermission @getAllRecipientPermissionsParams
    $allSendAsPermissions = $allRecipientPermissions | Where-Object { $_.AccessRights -eq 'SendAs' -and $_.AccessControlType -eq 'Allow' } | Select-Object Identity, Trustee
    $allSendAsPermissionsGrouped = $allSendAsPermissions | Group-Object -Property 'Identity' -AsHashTable -AsString
    Write-Information "Successfully queried [$($allSendAsPermissions.count)] recipient permissions"
    # Cleanup for memory
    $allRecipientPermissions = $null
    $allSendAsPermissions = $null

    foreach ($sharedMailbox in $sharedMailboxes) {
        # Full Access
        $getFullAccessPermissionsParams = @{
            Identity    = $sharedMailbox.Guid
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }
        $fullAccessUsers = @()
        $fullAccessList = Get-MailboxPermission @getFullAccessPermissionsParams
        $fullAccessPermissions = $fullAccessList | Where-Object { $_.AccessRights -eq 'FullAccess' -and $_.Deny -eq $false } | Select-Object User
        foreach ($record in $fullAccessPermissions) {
            $fullAccessUser = $userMailboxesUpnGrouped[$record.User].guid
            if ($fullAccessUser) { $fullAccessUsers += $fullAccessUser }
        }
        $numberOfAccounts = $fullAccessUsers.Count
        $permission = @{
            PermissionReference = @{
                Id         = $sharedMailbox.Guid
                Permission = 'FullAccess'
            }       
            Description         = $sharedMailbox.UserPrincipalName
            DisplayName         = $sharedMailbox.DisplayName + ' - Full Access'
        }
        # Batch permissions based on the amount of account references, 
        # to make sure the output objects are not above the limit
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $accountsBatchSize = 500
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize ) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($fullAccessUsers[$_]) })
                Write-Output $permission
            }
        }

        # Send As
        $sendAsUsers = @()
        $sendAsPermissions = $allSendAsPermissionsGrouped[$sharedMailbox.Name]
        foreach ($record in $sendAsPermissions) {
            $sendAsUser = $userMailboxesUpnGrouped[$record.Trustee].guid
            if ($sendAsUser) { $sendAsUsers += $sendAsUser }
        }
        $numberOfAccounts = $sendAsUsers.Count
        $permission = @{
            PermissionReference = @{
                Id         = $sharedMailbox.Guid
                Permission = 'SendAs'
            }       
            Description         = $sharedMailbox.UserPrincipalName
            DisplayName         = $sharedMailbox.DisplayName + ' - Send As'
        }
        # Batch permissions based on the amount of account references, 
        # to make sure the output objects are not above the limit
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $accountsBatchSize = 500
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize ) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($sendAsUsers[$_]) })
                Write-Output $permission
            }
        }

        # Send On Behalf
        $sendOnBehalfUsers = @()
        $sendOnPermissions = $sharedMailbox.GrantSendOnBehalfTo
        foreach ($record in $sendOnPermissions) {
            $sendOnBehalfUser = $userMailboxesExtDirObIdGrouped[$record].guid
            if ($sendOnBehalfUser) { $sendOnBehalfUsers += $sendOnBehalfUser }
        }
        $numberOfAccounts = $sendOnBehalfUsers.Count
        $permission = @{
            PermissionReference = @{
                Id         = $sharedMailbox.Guid
                Permission = 'SendOnBehalf'
            }       
            Description         = $sharedMailbox.UserPrincipalName
            DisplayName         = $sharedMailbox.DisplayName + ' - Send on Behalf'
        }
        # Batch permissions based on the amount of account references, 
        # to make sure the output objects are not above the limit
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $accountsBatchSize = 500
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize ) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($sendOnBehalfUsers[$_]) })
                Write-Output $permission
            }
        }
    }
    Write-Information 'Target permission import for shared mailboxes is completed'
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
    $actionMessage = "disconnecting from Microsoft Exchange Online"
    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }
    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    Write-Information "Disconnected from Microsoft Exchange Online"
}