#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Update
# Updates custom attributes
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

            Write-Information "Changed properties: $($accountChangedPropertiesObject | ConvertTo-Json)"

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

            $setMicrosoftExchangeOnlineAccountSplatParams = @{
                Identity    = $actionContext.References.Account
                Verbose     = $false
                ErrorAction = "Stop"
            }

            foreach ($accountNewProperty in $accountNewProperties) {
                $setMicrosoftExchangeOnlineAccountSplatParams["$($accountNewProperty.Name)"] = $accountNewProperty.Value
            }

            Write-Information "SplatParams: $($setMicrosoftExchangeOnlineAccountSplatParams | ConvertTo-Json)"

            if (-Not($actionContext.DryRun -eq $true)) {       
                $null = Set-Mailbox  @setMicrosoftExchangeOnlineAccountSplatParams

                Write-Information "Account with id [$($actionContext.References.Account)] successfully updated Old values: $($accountChangedPropertiesObject.oldValues | ConvertTo-Json). New values: $($accountChangedPropertiesObject.newValues | ConvertTo-Json)"

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Account with id [$($actionContext.References.Account)] successfully updated Old values: $($accountChangedPropertiesObject.oldValues | ConvertTo-Json). New values: $($accountChangedPropertiesObject.newValues | ConvertTo-Json)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Would update account with id [$($actionContext.References.Account)]. Old values: $($accountChangedPropertiesObject.oldValues | ConvertTo-Json). New values: $($accountChangedPropertiesObject.newValues | ConvertTo-Json)"
            }

            break
        }

        "NoChanges" {
            $actionMessage = "no changes to account"

            $outputContext.Data = $actionContext.Data
            $outputContext.PreviousData = $actionContext.Data

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
