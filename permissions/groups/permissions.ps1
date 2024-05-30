#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-Groups-List
# List groups as permissions
# Currently only Mail-enabled Security Group of Distribution Group are supported by the Exchange Online Management module
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
    "Get-DistributionGroup"
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

try {
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

    #region Connect to Exchange
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
    $actionMessage = "connecting to exchange"

    # Connect to Exchange Online in an unattended scripting scenario using an access token.
    Write-Verbose "Connecting to Exchange Online"

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
        
    Write-Verbose "Successfully connected to Exchange Online"
    #endregion Connect to Exchange

    #region Mail-enabled Security Groups
    #region Get Exchange Online Mail-enabled Security Groups
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroup?view=exchange-ps
    $actionMessage = "querying Microsoft Exchange Online Mail-enabled Security Groups"

    $getMicrosoftExchangeOnlineMailEnabledSecurityGroupsSplatParams = @{
        Filter      = "RecipientTypeDetails -eq 'MailUniversalSecurityGroup' -and IsDirSynced -eq 'False'"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $microsoftExchangeOnlineMailEnabledSecurityGroups = $null
    $microsoftExchangeOnlineMailEnabledSecurityGroups = Get-DistributionGroup @getMicrosoftExchangeOnlineMailEnabledSecurityGroupsSplatParams

    Write-Information "Queried Microsoft Exchange Online Mail-enabled Security Groups. Result count: $(($microsoftExchangeOnlineMailEnabledSecurityGroups | Measure-Object).Count)"
    #endregion Get Microsoft Exchange Online Mail-enabled Security Groups

    #region Send results to HelloID
    $microsoftExchangeOnlineMailEnabledSecurityGroups | ForEach-Object {
        # Shorten DisplayName to max. 100 chars
        $displayName = "Mail-enabled Security Group - $($_.displayName)"
        $displayName = $displayName.substring(0, [System.Math]::Min(100, $displayName.Length)) 
        
        $outputContext.Permissions.Add(
            @{
                displayName    = $displayName
                identification = @{
                    Id   = $_.Guid
                    Name = $_.displayName
                    Type = "Mail-enabled Security Group"
                }
            }
        )
    }
    #endregion Send results to HelloID
    #endregion Mail-enabled Security Groups

    #region Distribution Groups
    #region Get Exchange Online Distribution Groups
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroup?view=exchange-ps
    $actionMessage = "querying Microsoft Exchange Online Distribution Groups"

    $getMicrosoftExchangeOnlineDistributionGroupsSplatParams = @{
        Filter      = "RecipientTypeDetails -ne 'MailUniversalSecurityGroup' -and IsDirSynced -eq 'False'"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $microsoftExchangeOnlineDistributionGroups = $null
    $microsoftExchangeOnlineDistributionGroups = Get-DistributionGroup @getMicrosoftExchangeOnlineDistributionGroupsSplatParams

    Write-Information "Queried Microsoft Exchange Online Distribution Groups. Result count: $(($microsoftExchangeOnlineDistributionGroups | Measure-Object).Count)"
    #endregion Get Microsoft Exchange Online Distribution Groups

    #region Send results to HelloID
    $microsoftExchangeOnlineDistributionGroups | ForEach-Object {
        # Shorten DisplayName to max. 100 chars
        $displayName = "Distribution Group - $($_.displayName)"
        $displayName = $displayName.substring(0, [System.Math]::Min(100, $displayName.Length)) 
        
        $outputContext.Permissions.Add(
            @{
                displayName    = $displayName
                identification = @{
                    Id   = $_.Guid
                    Name = $_.displayName
                    Type = "Distribution Group"
                }
            }
        )
    }
    #endregion Send results to HelloID
    #endregion Distribution Groups
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

    # Set Success to false
    $outputContext.Success = $false

    # Required to write an error as the listing of permissions doesn't show auditlog
    Write-Error $auditMessage
}