#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Import
# Correlate to account
# PowerShell V2
#################################################

$entraIdCorrelationProperty = 'employeeId'

$entraPropertiesToSelect = @(
    "id",
    "$entraIdCorrelationProperty"
)

# Define PowerShell commands to import
$commands = @(
    "Get-User",
    "Get-EXOMailbox"
)

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

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Certificate
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($actionContext.Configuration.AppId)"
            'sub' = "$($actionContext.Configuration.AppId)"
            'aud' = "https://login.microsoftonline.com/$($actionContext.Configuration.TenantID)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $actionContext.Configuration.AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($actionContext.Configuration.TenantID)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
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
    Write-Information 'Starting target account import'

    # Define properties to query
    $importFields = @($actionContext.ImportFields) -replace '\..*', ''

    # Remove fieldmapping that are not available in 'Get-EXOMailbox'
    $importFields = $importFields | Where-Object { $_ -ne 'AutoReplyState' }
    $importFields = $importFields | Where-Object { $_ -ne 'InternalMessage' }
    $importFields = $importFields | Where-Object { $_ -ne 'ExternalMessage' }
    $importFields = $importFields | Where-Object { $_ -ne 'ExternalAudience' }
    $importFields = $importFields | Where-Object { $_ -ne $entraIdCorrelationProperty }
    
    # Add mandatory fields for HelloID to query and return
    if ('Guid' -notin $importFields) { $importFields += 'Guid' }
    if ('DisplayName' -notin $importFields) { $importFields += 'DisplayName' }
    if ('UserPrincipalName' -notin $importFields) { $importFields += 'UserPrincipalName' }
    if ('hiddenFromAddressListsEnabled' -notin $importFields) { $importFields += 'hiddenFromAddressListsEnabled' }
    if ('ExternalDirectoryObjectId' -notin $importFields) { $importFields += 'ExternalDirectoryObjectId' }

    Write-Information "Querying fields [$importFields]"

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

    $actionMessage = "connecting to Microsoft Graph"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate
    $headers = @{
        Authorization = "Bearer $entraToken"
        Accept        = 'application/json'
    }

    $actionMessage = "querying Entra ID users"
    $entraUsersById = @{}
    $getEntraIDAccountsSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/users?`$select=$($entraPropertiesToSelect -join ',')&`$top=999"
        Method      = 'GET'
        Headers     = $headers
        Verbose     = $false
        ErrorAction = 'Stop'
    }
    do {
        $entraUsersResponse = Invoke-RestMethod @getEntraIDAccountsSplatParams
        foreach ($entraUser in $entraUsersResponse.value) {
            $entraUsersById[$entraUser.id] = $entraUser
        }
        $getEntraIDAccountsSplatParams.Uri = $entraUsersResponse.'@odata.nextLink'
    } while (-not [string]::IsNullOrEmpty($getEntraIDAccountsSplatParams.Uri))

    Write-Information "Queried [$($entraUsersById.Count)] Entra ID users"
    
    $getExistingMailboxesSplatParams = @{
        ResultSize           = 'Unlimited'
        Properties           = @($importFields)
        RecipientTypeDetails = "UserMailbox"
        Verbose              = $false
        ErrorAction          = 'Stop'
    }

    $actionMessage = "querying Exchange Online mailboxes"
    $existingMailboxes = Get-EXOMailbox @getExistingMailboxesSplatParams
    Write-Information "Queried [$($existingMailboxes.Count)] Exchange Online user mailboxes"

    # Map the imported data to the account field mappings
    $actionMessage = "mapping Exchange Online mailboxes to Entra ID users"
    foreach ($importedAccount in $existingMailboxes) {
        $entraIdAccount = $null
        if ($entraUsersById.ContainsKey($importedAccount.ExternalDirectoryObjectId)) {
            $entraIdAccount = $entraUsersById[$importedAccount.ExternalDirectoryObjectId]
        }
        else {
            Write-Warning "No Entra ID user found for mailbox [$($importedAccount.Guid)] with ExternalDirectoryObjectId [$($importedAccount.ExternalDirectoryObjectId)]."
        }

        $data = @{}
        foreach ($field in $actionContext.ImportFields) {
            if ($null -ne $importedAccount.PSObject.Properties[$field]) {
                $data[$field] = $importedAccount.$field
            }
            elseif ($null -ne $entraIdAccount -and $null -ne $entraIdAccount.PSObject.Properties[$field]) {
                $data[$field] = $entraIdAccount.$field
            }
            else {
                $data[$field] = $null
            }
        }

        $enabled = $false
        if ($account.HiddenFromAddressListsEnabled -eq $false) {
            $enabled = $true
        }

        # Make sure the DisplayName has a value
        if ([string]::IsNullOrEmpty($importedAccount.displayName)) {
            $importedAccount.displayName = $importedAccount.Guid
        }
        # Make sure the Username has a value
        if ([string]::IsNullOrEmpty($importedAccount.UserPrincipalName)) {
            $importedAccount.UserPrincipalName = $importedAccount.Guid
        }
        # Return the result
        Write-Output @{
            AccountReference = $importedAccount.Guid
            DisplayName      = $importedAccount.DisplayName
            UserName         = $importedAccount.UserPrincipalName
            Enabled          = $enabled
            Data             = $data
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
