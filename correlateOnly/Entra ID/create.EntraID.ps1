#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Create
# Correlate to account
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($actionContext.Configuration.isDebug) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions

function Resolve-MicrosoftGraphAPIError {
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

function Convert-StringToBoolean($obj) {
    if ($obj -is [PSCustomObject]) {
        foreach ($property in $obj.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [string]) {
                $lowercaseValue = $value.ToLower()
                if ($lowercaseValue -eq "true") {
                    $obj.$($property.Name) = $true
                }
                elseif ($lowercaseValue -eq "false") {
                    $obj.$($property.Name) = $false
                }
            }
            elseif ($value -is [PSCustomObject] -or $value -is [System.Collections.IDictionary]) {
                $obj.$($property.Name) = Convert-StringToBoolean $value
            }
            elseif ($value -is [System.Collections.IList]) {
                for ($i = 0; $i -lt $value.Count; $i++) {
                    $value[$i] = Convert-StringToBoolean $value[$i]
                }
                $obj.$($property.Name) = $value
            }
        }
    }
    return $obj
}
#endregion functions

try {
    #region account
    # Define correlation
    $correlationField = $actionContext.CorrelationConfiguration.accountField
    $correlationValue = $actionContext.CorrelationConfiguration.personFieldValue

    # Define account object
    $account = [PSCustomObject]$actionContext.Data.PsObject.Copy()

    # Define properties to query
    $accountPropertiesToQuery = @("id") + $account.PsObject.Properties.Name | Select-Object -Unique

    # Remove properties of account object with null-values
    $account.PsObject.Properties | ForEach-Object {
        # Remove properties with null-values
        if ($_.Value -eq $null) {
            $account.PsObject.Properties.Remove("$($_.Name)")
        }
    }
    # Convert the properties of account object containing "TRUE" or "FALSE" to boolean 
    $account = Convert-StringToBoolean $account
    #endRegion account

    #region Verify correlation configuration and properties
    $actionMessage = "verifying correlation configuration and properties"

    if ($actionContext.CorrelationConfiguration.Enabled -eq $true) {
        if ([string]::IsNullOrEmpty($correlationField)) {
            throw "Correlation is enabled but not configured correctly."
        }
        
        if ([string]::IsNullOrEmpty($correlationValue)) {
            throw "The correlation value for [$correlationField] is empty. This is likely a mapping issue."
        }
    }
    else {
        Write-Warning "Correlation is disabled."
    }
    #endregion Verify correlation configuration and properties

    #region Create access token
    $actionMessage = "creating access token"
    
    $createAccessTokenBody = @{
        grant_type    = "client_credentials"
        client_id     = $actionContext.Configuration.AppId
        client_secret = $actionContext.Configuration.AppSecret
        resource      = "https://graph.microsoft.com"
    }

    $createAccessTokenSplatParams = @{
        Uri         = "https://login.microsoftonline.com/$($actionContext.Configuration.TenantID)/oauth2/token"
        Headers     = $headers
        Method      = "POST"
        Body        = $createAccessTokenBody
        ContentType = "application/x-www-form-urlencoded"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $createAccessTokenResonse = Invoke-RestMethod @createAccessTokenSplatParams

    Write-Verbose "Created access token. Result: $($createAccessTokenResonse | ConvertTo-Json)"
    #endregion Create access token

    #region Create headers
    $actionMessage = "creating headers"

    $headers = @{
        "Authorization"    = "Bearer $($createAccessTokenResonse.access_token)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json;charset=utf-8"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    Write-Verbose "Created headers. Result: $($headers | ConvertTo-Json)."
    #endregion Create headers

    #region Get account
    # Docs: https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying account where [$($correlationField)] = [$($correlationValue)]"

    $getMicrosoftEntraIDAccountSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/users?`$filter=$correlationField eq '$correlationValue'&`$select=$($accountPropertiesToQuery -join ',')"
        Method      = "GET"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    Write-Verbose "SplatParams: $($getMicrosoftEntraIDAccountSplatParams | ConvertTo-Json)"

    # Add header after printing splat
    $getMicrosoftEntraIDAccountSplatParams['Headers'] = $headers

    $getMicrosoftEntraIDAccountResponse = $null
    $getMicrosoftEntraIDAccountResponse = Invoke-RestMethod @getMicrosoftEntraIDAccountSplatParams
    $correlatedAccount = $getMicrosoftEntraIDAccountResponse.Value
        
    Write-Verbose "Queried account where [$($correlationField)] = [$($correlationValue)]. Result: $($correlatedAccount  | ConvertTo-Json)"
    #endregion Get account

    #region Calulate action
    $actionMessage = "calculating action"
    if (($correlatedAccount | Measure-Object).count -eq 1) {
        $actionAccount = "Correlate"
    }
    elseif (($correlatedAccount | Measure-Object).count -eq 0) {
        $actionAccount = "NotFound"
    }
    elseif (($correlatedAccount | Measure-Object).count -gt 1) {
        $actionAccount = "MultipleFound"
    }
    #endregion Calulate action
    
    #region Process
    switch ($actionAccount) {
        "Correlate" {
            #region Correlate account
            $actionMessage = "correlating to account"

            $outputContext.AccountReference = "$($correlatedAccount.id)"
            $outputContext.Data = $correlatedAccount.PsObject.Copy()

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = "CorrelateAccount" # Optionally specify a different action for this audit log
                    Message = "Correlated to account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json) on [$($correlationField)] = [$($correlationValue)]."
                    IsError = $false
                })

            $outputContext.AccountCorrelated = $true
            #endregion Correlate account

            break
        }

        "MultipleFound" {
            #region Multiple accounts found
            $actionMessage = "correlating to account"

            # Throw terminal error
            throw "Multiple accounts found where [$($correlationField)] = [$($correlationValue)]. Please correct this so the persons are unique."
            #endregion Multiple accounts found

            break
        }

        "NotFound" {
            #region No account found
            $actionMessage = "correlating to account"
        
            # Throw terminal error
            throw "No account found where [$($correlationField)] = [$($correlationValue)] while this connector only supports correlation."
            #endregion No account found

            break
        }
    }
    #endregion Process
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
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
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }

    # Check if accountreference is set, if not set, set this with default value as this must contain a value
    if ([String]::IsNullOrEmpty($outputContext.AccountReference) -and $actionContext.DryRun -eq $true) {
        $outputContext.AccountReference = "DryRun: Currently not available"
    }
}
