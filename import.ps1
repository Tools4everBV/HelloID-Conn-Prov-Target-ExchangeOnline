#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Import
# Correlate to account
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
    Write-Information 'Starting target account import'

    # Define properties to query
    $importFields = $($actionContext.ImportFields)
    $importFields = $importFields -replace '\..*', ''

    # Remove fieldmapping that are not avaliable in 'Get-EXOMailbox'
    $importFields = $importFields | Where-Object { $_ -ne 'AutoReplyState' }
    $importFields = $importFields | Where-Object { $_ -ne 'InternalMessage' }
    $importFields = $importFields | Where-Object { $_ -ne 'ExternalMessage' }
    
    # Add mandatory fields for HelloID to query and return
    if ('Guid' -notin $importFields) { $importFields += 'Guid' }
    if ('DisplayName ' -notin $importFields) { $importFields += 'DisplayName ' }
    if ('UserPrincipalName' -notin $importFields) { $importFields += 'UserPrincipalName' }
    if ('hiddenFromAddressListsEnabled' -notin $importFields) { $importFields += 'hiddenFromAddressListsEnabled ' }

    Write-Information "Querying fields [$importFields]"

    $actionMessage = "importing module [ExchangeOnlineManagement]"
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = 'Get-User,Get-EXOMailbox'
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
    Write-Information "Created access token."

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

    $getExistingMailboxesSplatParams = @{
        ResultSize           = 'Unlimited'
        Properties           = @($importFields)
        RecipientTypeDetails = "UserMailbox"
        Verbose              = $false
        ErrorAction          = 'Stop'
    }

    $existingMailboxes = Get-EXOMailbox @getExistingMailboxesSplatParams

    # Map the imported data to the account field mappings
    foreach ($account in $existingMailboxes) {
        $enabled = $false
        if ($account.HiddenFromAddressListsEnabled -eq $false) {
            $enabled = $true
        }

        # Make sure the DisplayName has a value
        if ([string]::IsNullOrEmpty($account.displayName)) {
            $account.displayName = $account.Guid
        }
        # Make sure the Username has a value
        if ([string]::IsNullOrEmpty($account.UserPrincipalName)) {
            $account.UserPrincipalName = $account.Guid
        }
        # Return the result
        Write-Output @{
            AccountReference = $account.Guid
            DisplayName      = $account.DisplayName
            UserName         = $account.UserPrincipalName
            Enabled          = $enabled
            Data             = $account
        }
    }
    Write-Information 'Target account import completed'
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
    $actionMessage = "disconnecting to Microsoft Exchange Online"
    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }
    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    Write-Information "Disconnected from Microsoft Exchange Online"
}