#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-SharedMailboxes-Import
# Correlate to permission
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Define PowerShell commands to import
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
    Write-Information 'Starting target shared mailbox permissions import'
    $actionMessage = "importing module [ExchangeOnlineManagement]"
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = "Stop"
    }
    $null = Import-Module @importModuleSplatParams
    Write-Information "Imported module [$($importModuleSplatParams.Name)]"

    $actionMessage = "retrieving certificate"
    $certificate = Get-MSEntraCertificate

    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
    $actionMessage = "connecting to Microsoft Exchange Online"

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
        ErrorAction           = "Stop"
    }

    $null = Connect-ExchangeOnline @createExchangeSessionSplatParams
        
    Write-Information "Connected to Microsoft Exchange Online"

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

    Write-Information "Successfully queried [$($userMailboxes.count)] user mailboxes"

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

    foreach ($sharedMailbox in $sharedMailboxes) {
        # If Full Access then permission is returned
        $getFullAccessPermissionsParams = @{
            Identity    = $sharedMailbox.Guid
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop'
        }
        $fullAccessUsers = @()
        $fullAccessPermissions = Get-EXOMailboxPermission @getFullAccessPermissionsParams | Where-Object { $_.AccessRights -eq 'FullAccess' -and $_.Deny -eq $false } | Select-Object User
        foreach ($record in $fullAccessPermissions) {
            $fullAccessUser = $userMailboxesUpnGrouped[$record.User].guid
            if ($fullAccessUser) { $fullAccessUsers += $fullAccessUser }
        }
        $numberOfAccounts = $fullAccessUsers.Count
        $numberOfFullAccess += $numberOfAccounts

        $permission = @{
            PermissionReference = @{
                Id = $sharedMailbox.Guid
            }       
            Description         = $sharedMailbox.UserPrincipalName
            DisplayName         = 'Shared Mailbox - ' + $sharedMailbox.DisplayName
        }
        # Batch permissions based on the amount of account references, 
        # to make sure the output objects are not above the limit
        $accountsBatchSize = 500
        if ($numberOfAccounts -gt 0) {
            $batches = 0..($numberOfAccounts - 1) | Group-Object { [math]::Floor($_ / $accountsBatchSize ) }
            foreach ($batch in $batches) {
                $permission.AccountReferences = [array]($batch.Group | ForEach-Object { @($fullAccessUsers[$_]) })
                Write-Output $permission
            }
        }
    }
    Write-Information "Target permission import for shared mailboxes completed. Full Access: [$numberOfFullAccess] accounts. Total shared mailboxes: [$($sharedMailboxes.count)]"
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