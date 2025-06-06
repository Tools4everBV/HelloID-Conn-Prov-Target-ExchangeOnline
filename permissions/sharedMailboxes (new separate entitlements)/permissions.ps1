#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-SharedMailboxes-List
# List shared mailboxes as permissions
# PowerShell V2
#####################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Define PowerShell commands to import
$commands = @(
    "Get-Recipient"
    , "Get-EXORecipient"
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

    #region Get Shared Mailboxes
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroup?view=exchange-ps
    $actionMessage = "querying Microsoft Exchange Online Shared Mailboxes"
    
    $getMicrosoftExchangeOnlineSharedMailboxesSplatParams = @{
        Properties           = @("Guid", "DisplayName")
        RecipientTypeDetails = "SharedMailbox"
        ResultSize           = "Unlimited"
        Verbose              = $false
        ErrorAction          = "Stop"
    }

    $getMicrosoftExchangeOnlineSharedMailboxesResponse = $null
    $getMicrosoftExchangeOnlineSharedMailboxesResponse = Get-EXORecipient @getMicrosoftExchangeOnlineSharedMailboxesSplatParams
    $microsoftExchangeOnlineSharedMailboxes = $getMicrosoftExchangeOnlineSharedMailboxesResponse | Select-Object Guid, DisplayName

    Write-Information "Queried Microsoft Exchange Online Shared Mailboxes. Result count: $(($microsoftExchangeOnlineSharedMailboxes | Measure-Object).Count)"
    #endregion Get Shared Mailboxes

    #region Send results to HelloID
    $microsoftExchangeOnlineSharedMailboxes | ForEach-Object {
        # Shorten DisplayName to max. 100 chars (83 because ' - Send on Behalf' is 17 char)
        $displayName = "Shared Mailbox - $($_.DisplayName)"
        $displayName = $displayName.substring(0, [System.Math]::Min(83, $displayName.Length))

        $outputContext.Permissions.Add(
            @{
                displayName    = $displayName + ' - Full Access'
                identification = @{
                    Id         = $_.Guid
                    Permission = "FullAccess"
                }
            }
        )
        $outputContext.Permissions.Add(
            @{
                displayName    = $displayName + ' - Send As'
                identification = @{
                    Id         = $_.Guid
                    Permission = "SendAs"
                }
            }
        )
        $outputContext.Permissions.Add(
            @{
                displayName    = $displayName + ' - Send on Behalf'
                identification = @{
                    Id         = $_.Guid
                    Permission = "SendOnBehalf"
                }
            }
        )
    }
    #endregion Send results to HelloID
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
    
    # Set Success to false
    $outputContext.Success = $false

    Write-Warning $warningMessage

    # Required to write an error as the listing of permissions doesn't show auditlog
    Write-Error $auditMessage
}
finally {
    #region Disconnect from Microsoft Exchange Online
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $actionMessage = "disconnecting from Microsoft Exchange Online"

    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }

    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    
    Write-Information "Disconnected from Microsoft Exchange Online"
    #endregion Disconnect from Microsoft Exchange Online
}