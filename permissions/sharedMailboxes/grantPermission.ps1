#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-SharedMailboxes-Grant
# Grant shared mailbox permission (full access, send as or send on behalf) to account
# PowerShell V2
#####################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# PowerShell commands to import
$commands = @(
    "Add-MailboxPermission"
    , "Add-RecipientPermission"
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

    #region Grant Mailbox permission
    foreach ($accessRight in $actionContext.References.Permission.AccessRights) {
        switch ($accessRight) {
            "Full Access" {
                #region Grant Full Access to account
                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/add-mailboxpermission?view=exchange-ps
                $actionMessage = "granting [FullAccess] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantFullAccessPermissionSplatParams = @{
                    Identity        = $actionContext.References.Permission.id
                    User            = $actionContext.References.Account
                    AccessRights    = 'FullAccess'
                    InheritanceType = 'All'
                    AutoMapping     = $true
                    Confirm         = $false
                    Verbose         = $false
                    ErrorAction     = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Information "SplatParams: $($grantFullAccessPermissionSplatParams | ConvertTo-Json)"

                    $null = Add-MailboxPermission @grantFullAccessPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Granted [FullAccess] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant [FullAccess] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
                #endregion Grant Full Access to account
            }
            "Send As" {
                #region Grant Send As to account
                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/add-recipientpermission?view=exchange-ps
                $actionMessage = "granting [SendAs] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantSendAsPermissionSplatParams = @{
                    Identity     = $actionContext.References.Permission.id
                    Trustee      = $actionContext.References.Account
                    AccessRights = 'SendAs'
                    Confirm      = $false
                    Verbose      = $false
                    ErrorAction  = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Information "SplatParams: $($grantSendAsPermissionSplatParams | ConvertTo-Json)"

                    $null = Add-RecipientPermission @grantSendAsPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Granted [SendAs] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant [SendAs] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
                #endregion Grant Send As to account
            }
            "Send on Behalf" {
                #region Grant Send On Behalf to account
                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
                $actionMessage = "granting [SendOnBehalf] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantSendOnBehalfPermissionSplatParams = @{
                    Identity            = $actionContext.References.Permission.id
                    GrantSendOnBehalfTo = @{add = "$($actionContext.References.Account)" }
                    Confirm             = $false
                    Verbose             = $false
                    ErrorAction         = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Information "SplatParams: $($grantSendOnBehalfPermissionSplatParams | ConvertTo-Json)"

                    $null = Set-Mailbox @grantSendOnBehalfPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Granted [SendOnBehalf] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant [SendOnBehalf] to mailbox [$($actionContext.PermissionDisplayName)] with id [$($actionContext.References.Permission.id)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
                #endregion Grant Send On Behalf to account
            }
        }
    }
    #endregion Grant Mailbox permission
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