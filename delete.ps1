#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Delete
# Sets auto-reply configuration
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# PowerShell commands to import
$commands = @(
    "Get-MailboxAutoReplyConfiguration",
    "Set-MailboxAutoReplyConfiguration"
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
 
    Write-Information "Created access token."
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

    #region Get account
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-user?view=exchange-ps
    $actionMessage = "querying account where [Identity] = [$($actionContext.References.Account)]"

    $getMailboxAutoReplyConfigSplatParams = @{
        Identity    = $actionContext.References.Account
        Verbose     = $false
        ErrorAction = "Stop"
    }
    try {
        $correlatedAccount = Get-MailboxAutoReplyConfiguration @getMailboxAutoReplyConfigSplatParams
    }
    catch {
        if ($_ -like "*The operation couldn't be performed because Identity:`"$($actionContext.References.Account)`" couldn't be found*") {
            $correlatedAccount = $null
        }
        else {
            throw
        }
    }
        
    Write-Information "Queried account where [Identity] = [$($actionContext.References.Account)]. Result: $($correlatedAccount | ConvertTo-Json)"
    #endregion Get account

    #region Calulate action
    $actionMessage = "calculating action"

    if (($correlatedAccount | Measure-Object).count -eq 1) {
        if ($correlatedAccount.AutoReplyState -eq $actionContext.Data.AutoReplyState) {
            $actionAccount = "NoChanges"
        }
        else {
            $actionAccount = "Delete"
        }
    }
    else {
        $actionAccount = "NotFound"
    }

    Write-Information "Action: $actionAccount"
    #endregion Calulate action

    #region Process
    switch ($actionAccount) {
        "Delete" {
            $actionMessage = "setting autoreply to account"

            $setMicrosoftExchangeOnlineAccountSplatParams = @{
                Identity         = $actionContext.References.Account
                AutoReplyState   = $actionContext.Data.AutoReplyState
                InternalMessage  = $actionContext.Data.InternalMessage
                ExternalMessage  = $actionContext.Data.ExternalMessage
                ExternalAudience = $actionContext.Data.ExternalAudience
                Verbose          = $false
                ErrorAction      = "Stop"
            }
    
            Write-Information "SplatParams: $($setMicrosoftExchangeOnlineAccountSplatParams | ConvertTo-Json)"

            if (-Not($actionContext.DryRun -eq $true)) {       
                $null = Set-MailboxAutoReplyConfiguration  @setMicrosoftExchangeOnlineAccountSplatParams

                Write-Information "Account with id [$($actionContext.References.Account)] successfully deleted [AutoReplyState = $($actionContext.Data.AutoReplyState)]"

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Account with id [$($actionContext.References.Account)] successfully deleted [AutoReplyState = $($actionContext.Data.AutoReplyState)]"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Would set account with id [$($actionContext.References.Account)] to [AutoReplyState = $($actionContext.Data.AutoReplyState)]"
            }

            break
        }

        "NoChanges" {
            $actionMessage = "skipping setting autoreply to account"

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = "Account with AccountReference [$($actionContext.References.Account)] successfully deleted (skipped AutoReplyState already configured)"
                    IsError = $false
                })

            break
        }

        "NotFound" {
            $actionMessage = "skipping deleting account with AccountReference [$($actionContext.References.Account)]"
    
            Write-Information "Account with AccountReference [$($actionContext.References.Account)] successfully deleted (skipped not found)"
                
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = "Account with AccountReference [$($actionContext.References.Account)] successfully deleted (skipped not found)"
                    IsError = $false
                })

            break
        }
    }
    #endregion Process
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