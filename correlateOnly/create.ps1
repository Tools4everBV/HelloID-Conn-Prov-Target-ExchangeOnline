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

# Define PowerShell commands to import
$commands = @(
    "Get-User"
)

#region functions
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

function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ''
        }
        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.Powershell.Commands.HttpResponseException') {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}
#endregion functions

#region account
# Define correlation
$correlationField = $actionContext.CorrelationConfiguration.accountField
$correlationValue = $actionContext.CorrelationConfiguration.accountFieldValue

$account = [PSCustomObject]$actionContext.Data
# Remove properties with null-values
$account.PsObject.Properties | ForEach-Object {
    # Remove properties with null-values
    if ($_.Value -eq $null) {
        $account.PsObject.Properties.Remove("$($_.Name)")
    }
}
# Convert the properties containing "TRUE" or "FALSE" to boolean
$account = Convert-StringToBoolean $account

# Define properties to query
$accountPropertiesToQuery = @("guid") + $account.PsObject.Properties.Name | Select-Object -Unique
#endRegion account

try {
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
        throw "Correlation is disabled while this connector only supports correlation."
    }
    #endregion Verify correlation configuration and properties

    #region Import module
    $actionMessage = "importing module"

    $moduleName = "ExchangeOnlineManagement"

    # Check if module is already imported or available on disk
    $module = Get-Module -Name $moduleName -ListAvailable -ErrorAction SilentlyContinue -Verbose:$false

    if ($module) {
        Write-Verbose "Module [$moduleName] is already imported."
    }
    else {
        # Check if module is available in online gallery
        if (Find-Module -Name $moduleName) {
            # Import module with specified commands
            $module = Import-Module $ModuleName -Cmdlet $commands -Verbose:$false
            Write-Verbose "Imported module [$ModuleName]"
        }
        else {
            # If the module is not imported, not available and not in the online gallery then abort
            throw "Module [$ModuleName] is not available. Please install the module using: Install-Module -Name [$ModuleName] -Force"
        }
    }
    #endregion Import module

    #region Create access token
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
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

    $createdAccessToken = Invoke-RestMethod @createAccessTokenSplatParams
    $accessToken = $createdAccessToken.access_token

    Write-Verbose "Created access token. Result: $($accessToken | ConvertTo-Json)"
    #endregion Create access token

    #region Connect to Microsoft Exchange Online
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
    $actionMessage = "connecting to Microsoft Exchange Online"

    # Connect to Microsoft Exchange Online Online in an unattended scripting scenario using an access token.
    Write-Verbose "Connecting to Microsoft Exchange Online"

    $createExchangeSessionSplatParams = @{
        Organization          = $actionContext.Configuration.Organization
        AppID                 = $actionContext.Configuration.AppId
        AccessToken           = $accessToken
        CommandName           = $commands
        ShowBanner            = $false
        ShowProgress          = $false
        TrackPerformance      = $false
        SkipLoadingCmdletHelp = $true
        SkipLoadingFormatData = $true
        ErrorAction           = "Stop"
    }

    $createdExchangeSession = Connect-ExchangeOnline @createExchangeSessionSplatParams
        
    Write-Verbose "Successfully connected to Microsoft Exchange Online"
    #endregion Connect to Microsoft Exchange Online

    #region Get Microsoft Exchange Online account
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-user?view=exchange-ps
    $actionMessage = "querying Microsoft Exchange Online account where [$($correlationField)] = [$($correlationValue)]"

    $getMicrosoftExchangeOnlineAccountSplatParams = @{
        Filter      = "$($correlationField) -eq '$($correlationValue)'"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $currentMicrosoftExchangeOnlineAccount = $null
    $currentMicrosoftExchangeOnlineAccount = Get-User @getMicrosoftExchangeOnlineAccountSplatParams | Select-Object $accountPropertiesToQuery

    Write-Verbose "Queried Microsoft Exchange Online account where [$($correlationField)] = [$($correlationValue)]. Result: $($currentMicrosoftExchangeOnlineAccount | ConvertTo-Json)"
    #endregion Get Microsoft Exchange Online account

    #region Account
    #region Calulate action
    $actionMessage = "calculating action"
    if (($currentMicrosoftExchangeOnlineAccount | Measure-Object).count -eq 0) {
        $actionAccount = "NotFound"
    }
    elseif (($currentMicrosoftExchangeOnlineAccount | Measure-Object).count -eq 1) {
        $actionAccount = "Correlate"
    }
    elseif (($currentMicrosoftExchangeOnlineAccount | Measure-Object).count -gt 1) {
        $actionAccount = "MultipleFound"
    }
    #endregion Calulate action
    
    #region Process
    switch ($actionAccount) {
        "Correlate" {
            #region Correlate account
            $actionMessage = "correlating to account"

            $outputContext.AccountReference = "$($currentMicrosoftExchangeOnlineAccount.Guid)"
            $outputContext.Data = $currentMicrosoftExchangeOnlineAccount

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
    #endregion Account
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        Write-Warning "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        Write-Warning "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            # Action  = "" # Optional
            Message = $auditMessage
            IsError = $true
        })
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if ($outputContext.AuditLogs.IsError -contains $true) {
        $outputContext.Success = $false
    }
    else {
        $outputContext.Success = $true
    }

    # Check if accountreference is set, if not set, set this with default value as this must contain a value
    if ([String]::IsNullOrEmpty($outputContext.AccountReference) -and $actionContext.DryRun -eq $true) {
        $outputContext.AccountReference = "DryRun: Currently not available"
    }
}