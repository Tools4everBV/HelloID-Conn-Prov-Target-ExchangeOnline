#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-SharedMailboxes-Revoke
# Revoke shared mailbox permission (full access, send as or send on behalf) from account
# PowerShell V2
#####################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Permission definition settings
$accessRights = @("FullAccess", "SendAs") # SendOnBehalf

# PowerShell commands to import
$commands = @(
    "Remove-MailboxPermission"
    , "Remove-RecipientPermission"
    , "Set-Mailbox"
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
#endregion functions

try {
    #region Verify account reference
    $actionMessage = "verifying account reference"
    
    if ([string]::IsNullOrEmpty($($actionContext.References.Account))) {
        throw "The account reference could not be found"
    }
    #endregion Verify account reference

    #region Import module
    $actionMessage = "importing module [ExchangeOnlineManagement]"
    
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $null = Import-Module @importModuleSplatParams

    Write-Information "Imported module [$($importModuleSplatParams.Name)]"
    #endregion Create access token

    #region Create access token
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
    #endregion Create access token

    #region Connect to Microsoft Exchange Online
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
    #endregion Connect to Microsoft Exchange Online

    #region Revoke Mailbox permission
    foreach ($accessRight in $accessRights) {
        switch ($accessRight) {
            "FullAccess" {
                #region Revoke Full Access from account
                try {
                    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/remove-mailboxpermission?view=exchange-ps
                    $actionMessage = "revoking [FullAccess] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                    $revokeFullAccessPermissionSplatParams = @{
                        Identity        = $actionContext.References.Permission.id
                        User            = $actionContext.References.Account
                        AccessRights    = 'FullAccess'
                        InheritanceType = 'All'
                        Confirm         = $false
                        Verbose         = $false
                        ErrorAction     = "Stop"
                    }

                    if (-Not($actionContext.DryRun -eq $true)) {
                        Write-Information "SplatParams: $($revokeFullAccessPermissionSplatParams | ConvertTo-Json)"

                        $null = Remove-MailboxPermission @revokeFullAccessPermissionSplatParams

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Revoked [FullAccess] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning "DryRun: Would revoke [FullAccess] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                    }
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
        
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Permission.id)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        throw $auditMessage
                    }
                }
                #endregion Revoke Full Access from account
            }
            "SendAs" {
                #region Revoke Send As from account
                try {
                    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/remove-recipientpermission?view=exchange-ps
                    $actionMessage = "revoking [SendAs] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                    $revokeSendAsPermissionSplatParams = @{
                        Identity     = $actionContext.References.Permission.id
                        Trustee      = $actionContext.References.Account
                        AccessRights = 'SendAs'
                        Confirm      = $false
                        Verbose      = $false
                        ErrorAction  = "Stop"
                    }

                    if (-Not($actionContext.DryRun -eq $true)) {
                        Write-Information "SplatParams: $($revokeSendAsPermissionSplatParams | ConvertTo-Json)"

                        $null = Remove-RecipientPermission @revokeSendAsPermissionSplatParams

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Revoked [SendAs] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning "DryRun: Would revoke [SendAs] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                    }
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
        
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Permission.id)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        throw $auditMessage
                    }
                }
                #endregion Revoke Send As from account
            }
            "SendOnBehalf" {
                #region Revoke Send On Behalf from account
                try {
                    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
                    $actionMessage = "revoking [SendOnBehalf] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                    $revokeSendOnBehalfPermissionSplatParams = @{
                        Identity            = $actionContext.References.Permission.id
                        GrantSendOnBehalfTo = @{remove = "$($actionContext.References.Account)" }
                        Confirm             = $false
                        Verbose             = $false
                        ErrorAction         = "Stop"
                    }

                    if (-Not($actionContext.DryRun -eq $true)) {
                        Write-Information "SplatParams: $($revokeSendOnBehalfPermissionSplatParams | ConvertTo-Json)"

                        $null = Set-Mailbox @revokeSendOnBehalfPermissionSplatParams

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Revoked [SendOnBehalf] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning "DryRun: Would revoke [SendOnBehalf] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                    }
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
        
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Permission.id)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        throw $auditMessage
                    }
                }
                #endregion Revoke Send On Behalf from account
            }
        }
    }
    #endregion Revoke Mailbox permission
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

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            # Action  = "" # Optional
            Message = $auditMessage
            IsError = $true
        })
}
finally {
    #region Disconnect from Microsoft Exchange Online
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $actionMessage = "disconnecting to Microsoft Exchange Online"

    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }

    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    
    Write-Information "Disconnected from Microsoft Exchange Online"
    #endregion Disconnect from Microsoft Exchange Online

    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
}