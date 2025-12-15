#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Update
# Updates custom attributes
# Updates Emailadresses (proxyaddresses) preserving existing lines.
# Optionally removes any SPO:SPO_ addresses from the emailAddresses list. This will be regenerated upon next SharePoint Online license assignment.
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Define PowerShell commands to import
$commands = @(
    "Get-EXOMailbox",
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
    #region Verify account reference and action context data
    $actionMessage = "verifying account reference"
    
    if ([string]::IsNullOrEmpty($($actionContext.References.Account))) {
        throw "The account reference could not be found"
    }
    if ([string]::IsNullOrEmpty($($actionContext.Data))) {
        throw "Action context data is empty, add fields to update action or remove the update script"
    }
    #endregion Verify account reference and action context data

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

    #region Get account
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-user?view=exchange-ps
    $actionMessage = "querying account where [Identity] = [$($actionContext.References.Account)]"

    $accountPropertiesToQuery = @("guid", "displayname") + $($outputContext.Data.PsObject.Properties.Name).ToLower() | Select-Object -Unique
    $getMicrosoftExchangeOnlineAccountSplatParams = @{
        Identity    = $actionContext.References.Account
        Properties  = $accountPropertiesToQuery
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $correlatedAccount = Get-EXOMailbox  @getMicrosoftExchangeOnlineAccountSplatParams | Select-Object $accountPropertiesToQuery
    
    $outputContext.PreviousData = $correlatedAccount | Select-Object $outputContext.Data.PsObject.Properties.Name

    Write-Information "Queried account where [Identity] = [$($actionContext.References.Account)]. Result: $($correlatedAccount  | ConvertTo-Json)"
    #endregion Get account

    #region Calulate action
    $actionMessage = "calculating action"

    if (($correlatedAccount | Measure-Object).count -eq 1) {
        # Check if we're processing email addresses
        if ($actionContext.Data.PSObject.Properties.Name -contains "emailAddresses") {
            # Merge and ensure uniqueness of existing and new emailAddresses
            $mergedEmailAddresses = @($correlatedAccount.emailAddresses) + $actionContext.Data.emailAddresses | Sort-Object -Unique
            # Get the primary SMTP address from the mapped properties
            $primarySMTP = $actionContext.Data.emailAddresses | Where-Object { $_ -cmatch '^SMTP:' }
            if ($primarySMTP.Count -gt 1) {
                throw 'Multiple primary SMTP addresses found in the mapped properties. Please ensure only one is set.'
            }
            
            # Optionaly remove any SPO:SPO_ addresses from the merged list
            #$mergedEmailAddresses = $mergedEmailAddresses | Where-Object { $_ -notmatch '^SPO:SPO_' }
            
            # Ensure the primary SMTP is set correctly in the merged list
            $mergedEmailAddresses = $mergedEmailAddresses | ForEach-Object {
                if ($_ -cmatch '^SMTP:') {
                    $_.ToLower() -replace '^smtp:', 'smtp:'
                }
                else {
                    $_
                }
            }
            # Add the primary SMTP address at the beginning of the list
            $mergedEmailAddresses = @($primarySMTP) + @(($mergedEmailAddresses | Where-Object { $_ -ne $primarySMTP }))
            
            # Update the actionContext.Data with the merged email addresses
            $actionContext.Data.emailAddresses = $mergedEmailAddresses
        }

        $accountPropertiesToCompare = $actionContext.Data | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name

        $accountSplatCompareProperties = @{
            ReferenceObject  = $correlatedAccount.PSObject.Properties | Where-Object { $_.Name -in $accountPropertiesToCompare }
            DifferenceObject = $actionContext.Data.PSObject.Properties | Where-Object { $_.Name -in $accountPropertiesToCompare }
        }

        if ($null -ne $accountSplatCompareProperties.ReferenceObject -and $null -ne $accountSplatCompareProperties.DifferenceObject) {
            $accountPropertiesChanged = Compare-Object @accountSplatCompareProperties -PassThru
            $accountOldProperties = $accountPropertiesChanged | Where-Object { $_.SideIndicator -eq "<=" }
            $accountNewProperties = $accountPropertiesChanged | Where-Object { $_.SideIndicator -eq "=>" }
        }

        if ($accountNewProperties) {
            # Create custom object with old and new values
            $accountChangedPropertiesObject = [PSCustomObject]@{
                OldValues = @{}
                NewValues = @{}
            }

            # Add the old properties to the custom object with old and new values
            foreach ($accountOldProperty in $accountOldProperties) {
                $accountChangedPropertiesObject.OldValues.$($accountOldProperty.Name) = $accountOldProperty.Value
            }

            # Add the new properties to the custom object with old and new values
            foreach ($accountNewProperty in $accountNewProperties) {
                $accountChangedPropertiesObject.NewValues.$($accountNewProperty.Name) = $accountNewProperty.Value
            }

            $actionAccount = "Update"
        }
        else {
            $actionAccount = "NoChanges"
        }
    }
    else {
        $actionAccount = "NotFound"
    }
    #endregion Calulate action
    
    #region Process
    switch ($actionAccount) {
        "Update" {
            $actionMessage = "updating account"

            Write-Information "Account property(s) required to update: $($propertiesChanged.Name -join ', ')"

            $setMicrosoftExchangeOnlineAccountSplatParams = @{
                Identity    = $actionContext.References.Account
                Verbose     = $false
                ErrorAction = "Stop"
            }

            foreach ($accountNewProperty in $accountNewProperties) {
                $setMicrosoftExchangeOnlineAccountSplatParams["$($accountNewProperty.Name)"] = $accountNewProperty.Value
            }
            if (-Not($actionContext.DryRun -eq $true)) {       
                $null = Set-Mailbox  @setMicrosoftExchangeOnlineAccountSplatParams

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Account with id [$($actionContext.References.Account)] successfully updated. Account property(s) updated: [$($propertiesChanged.name -join ',')]" 
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Would update account with id [$($actionContext.References.Account)]. Account property(s) to update: [$($propertiesChanged.name -join ',')]"
            }

            break
        }

        "NoChanges" {
            $actionMessage = "no changes to account"

            $outputContext.Data = $correlatedAccount | Select-Object $outputContext.Data.PsObject.Properties.Name
            $outputContext.PreviousData = $correlatedAccount | Select-Object $outputContext.Data.PsObject.Properties.Name

            Write-Information "Account with id [$($actionContext.References.Account)] successfully checked. No changes required"

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = "Account with id [$($actionContext.References.Account)] successfully checked. No changes required"
                    IsError = $false
                })

            break
        }

        "NotFound" {
            $actionMessage = "updating account"
        
            Write-Information "No account found where [Identity] = [$($actionContext.References.Account)]. Possibly indicating that it could be deleted, or not correlated."
                
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = "No account found where [Identity] = [$($actionContext.References.Account)]. Possibly indicating that it could be deleted, or not correlated."
                    IsError = $true
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
