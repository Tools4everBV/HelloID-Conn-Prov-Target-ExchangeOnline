#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-LitigationHold-Grant
# Enable litigation hold and set the duration on mailbox
# PowerShell V2
#####################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Define PowerShell commands to import
$commands = @(
    "Set-Mailbox"
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
    #endregion Import module

    if ($actionContext.Configuration.UseCertificate -eq $true) {
        Write-Information "Connecting to Exchange Online with certificate"

        #region Retrieving certificate
        $actionMessage = "retrieving certificate"
        $certificate = Get-MSEntraCertificate
        #endregion Retrieving certificate

        #region Connect to Microsoft Exchange Online
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
        #endregion Connect to Microsoft Exchange Online
    }
    else {
        Write-Information "Connecting to Exchange Online with secret"
        
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
    }

    #region Enable litigation hold
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
    $actionMessage = "enabling litigation hold and set the duration to [$($actionContext.References.Permission.Duration)] days on mailbox [$($actionContext.References.Account)]"

    $enableLitigationHoldSplatParams = @{
        Identity               = $actionContext.References.Account
        LitigationHoldEnabled  = $true
        LitigationHoldDuration = $actionContext.References.Permission.Duration
        Verbose                = $false
        ErrorAction            = "Stop"
    }

    Write-Information "SplatParams: $($enableLitigationHoldSplatParams | ConvertTo-Json)"

    if (-Not($actionContext.DryRun -eq $true)) {
        $null = Set-Mailbox @enableLitigationHoldSplatParams

        $outputContext.AuditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Enabled litigation hold and set the duration to [$($actionContext.References.Permission.Duration)] days on mailbox [$($actionContext.References.Account)]."
                IsError = $false
            })
    }
    else {
        Write-Warning "DryRun: Would enabled litigation hold and set the duration to [$($actionContext.References.Permission.Duration)] days on mailbox [$($actionContext.References.Account)]."
    }
    #endregion Enable litigation hold
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