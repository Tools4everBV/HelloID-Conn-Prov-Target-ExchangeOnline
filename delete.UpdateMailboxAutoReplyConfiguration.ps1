#####################################################
# HelloID-Conn-Prov-Target-ExchangeOnline-Delete-Update-MailboxAutoReplyConfiguration
#
# Version: 2.0.0
#####################################################
# Initialize default values
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json
# The accountReference object contains the Identification object provided in the account create call
$aRef = $accountReference | ConvertFrom-Json
$success = $false # Set to false at start, at the end, only when no error occurs it is set to true
$auditLogs = [System.Collections.Generic.List[PSCustomObject]]::new()

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($($c.isDebug)) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Define configuration properties as required
$requiredConfigurationFields = @("AzureADOrganization", "AzureADTenantId", "AzureADAppId", "AzureADAppSecret")

# Used to connect to Exchange Online in an unattended scripting scenario using an App ID and App Secret to create an Access Token.
$AADOrganization = $c.AzureADOrganization
$AADTenantId = $c.AzureADTenantId
$AADAppID = $c.AzureADAppId
$AADAppSecret = $c.AzureADAppSecret

# PowerShell commands to import
$commands = @(
    "Get-User" # Always required
    , "Get-EXOMailbox"
    , "Set-MailboxAutoReplyConfiguration"
)

# Change mapping here
# Remove externalId from manager name
$primaryManagerName = ($($p.PrimaryManager.DisplayName) -replace " \($($p.PrimaryManager.ExternalId)\)", '')
$primaryManagerEmail = $($p.PrimaryManager.Email)
$account = [PSCustomObject]@{
    AutoReplyState    = 'Enabled'
    InternalMessage   = "Dear colleague, thank you for your message. I am no longer employed at Enyoi. Your mail will be forwarded to $($primaryManagerName)"
    ExternalMessage   = "Dear Sir, Madam, Thank you for your email. I am no longer employed at Enyoi. Your mail is automatically forwarded to my colleague $($primaryManagerName) with mail address $($primaryManagerEmail)"
}

# Define account properties as required
$requiredAccountFields = @("AutoReplyState", "InternalMessage", "ExternalMessage")


#region functions
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
            ErrorMessage          = ""
        }
        if ($ErrorObject.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException") {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}

function Get-ErrorMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $errorMessage = [PSCustomObject]@{
            VerboseErrorMessage = $null
            AuditErrorMessage   = $null
        }

        if ( $($ErrorObject.Exception.GetType().FullName -eq "Microsoft.PowerShell.Commands.HttpResponseException") -or $($ErrorObject.Exception.GetType().FullName -eq "System.Net.WebException")) {
            $httpErrorObject = Resolve-HTTPError -Error $ErrorObject

            $errorMessage.VerboseErrorMessage = $httpErrorObject.ErrorMessage

            $errorMessage.AuditErrorMessage = $httpErrorObject.ErrorMessage
        }

        # If error message empty, fall back on $ex.Exception.Message
        if ([String]::IsNullOrEmpty($errorMessage.VerboseErrorMessage)) {
            $errorMessage.VerboseErrorMessage = $ErrorObject.Exception.Message
        }
        if ([String]::IsNullOrEmpty($errorMessage.AuditErrorMessage)) {
            $errorMessage.AuditErrorMessage = $ErrorObject.Exception.Message
        }

        Write-Output $errorMessage
    }
}
#endregion functions

try {
    # Check if required fields are available in configuration object
    $incompleteConfiguration = $false
    foreach ($requiredConfigurationField in $requiredConfigurationFields) {
        if ($requiredConfigurationField -notin $c.PsObject.Properties.Name) {
            $incompleteConfiguration = $true
            Write-Warning "Required configuration object field [$requiredConfigurationField] is missing"
        }

        if ([String]::IsNullOrEmpty($c.$requiredConfigurationField)) {
            $incompleteConfiguration = $true
            Write-Warning "Required configuration object field [$requiredConfigurationField] has a null or empty value"
        }
    }

    if ($incompleteConfiguration -eq $true) {
        throw "Configuration object incomplete, cannot continue."
    }

    # Check if required fields are available in account object
    $incompleteAccount = $false
    foreach ($requiredAccountField in $requiredAccountFields) {
        if ($requiredAccountField -notin $account.PsObject.Properties.Name) {
            $incompleteAccount = $true
            Write-Warning "Required account object field [$requiredAccountField] is missing"
        }

        if ([String]::IsNullOrEmpty($account.$requiredAccountField)) {
            $incompleteAccount = $true
            Write-Warning "Required account object field [$requiredAccountField] has a null or empty value"
        }
    }

    if ($incompleteAccount -eq $true) {
        throw "Account object incomplete, cannot continue."
    }

    try {           
        # Import module
        $moduleName = "ExchangeOnlineManagement"

        # If module is imported say that and do nothing
        if (Get-Module | Where-Object { $_.Name -eq $ModuleName }) {
            Write-Verbose "Module [$ModuleName] is already imported."
        }
        else {
            # If module is not imported, but available on disk then import
            if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $ModuleName }) {
                $module = Import-Module $ModuleName -Cmdlet $commands -Verbose:$false
                Write-Verbose "Imported module [$ModuleName]"
            }
            else {
                # If the module is not imported, not available and not in the online gallery then abort
                throw "Module [$ModuleName] is not available. Please install the module using: Install-Module -Name [$ModuleName] -Force"
            }
        }
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error importing module [$ModuleName]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        continue
    }

    # Connect to Exchange
    try {
        # Create access token
        Write-Verbose "Creating Access Token"

        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$AADTenantId/oauth2/token"
        
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$AADAppID"
            client_secret = "$AADAppSecret"
            resource      = "https://outlook.office365.com"
        }
        
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType "application/x-www-form-urlencoded"
        $accessToken = $Response.access_token

        # Connect to Exchange Online in an unattended scripting scenario using an access token.
        Write-Verbose "Connecting to Exchange Online"

        $exchangeSessionParams = @{
            Organization     = $AADOrganization
            AppID            = $AADAppID
            AccessToken      = $accessToken
            CommandName      = $commands
            ShowBanner       = $false
            ShowProgress     = $false
            TrackPerformance = $false
            ErrorAction      = "Stop"
        }
        $exchangeSession = Connect-ExchangeOnline @exchangeSessionParams
        
        Write-Information "Successfully connected to Exchange Online"
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error connecting to Exchange Online. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        continue
    }

    # Get Exchange Online Mailbox
    try {
        Write-Verbose "Querying EXO mailbox [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
            
        $mailbox = Get-EXOMailbox -Identity $aRef.Guid -ErrorAction Stop

        if (($mailbox | Measure-Object).Count -eq 0) {
            throw "Could not find a EXO mailbox [$($aRef.UserPrincipalName) ($($aRef.Guid))]" 
        }

        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Successfully queried EXO mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]"
                IsError = $false
            })
    }
    catch { 
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error querying EXO mailbox [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })
    }

    # Update Mailbox AutoReply Configuration
    try {
        Write-Verbose "Updating autoreply configuration for mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]: $($mailboxSplatParams | ConvertTo-Json)"

        $mailboxSplatParams = @{
            Identity        = $($aRef.Guid)
            AutoReplyState  = $($account.AutoReplyState)
            InternalMessage = $($account.InternalMessage)
            ExternalMessage = $($account.ExternalMessage)
        }

        if ($dryRun -eq $false) {
            $updateMailbox = Set-MailboxAutoReplyConfiguration  @mailboxSplatParams -ErrorAction Stop

            $auditLogs.Add([PSCustomObject]@{
                    Action  = "CreateAccount"
                    Message = "Successfully updated autoreply configuration for mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]: $($mailboxSplatParams | ConvertTo-Json)"
                    IsError = $false
                })
        }
        else {
            Write-Warning "DryRun: would update autoreply configuration for mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]: $($mailboxSplatParams | ConvertTo-Json)"
        }
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error updating autoreply configuration for mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]: $($mailboxSplatParams | ConvertTo-Json). Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })
    }
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($auditLogs.IsError -contains $true)) {
        $success = $true
    }

    # Send results
    $result = [PSCustomObject]@{
        Success          = $success
        AccountReference = $aRef
        AuditLogs        = $auditLogs
        Account          = $account

        # Optionally return data for use in other systems
        ExportData       = [PSCustomObject]@{
            DisplayName       = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            Guid              = $mailbox.Guid
        }
    }

    Write-Output ($result | ConvertTo-Json -Depth 10)
}