#####################################################
# HelloID-Conn-Prov-Target-ExchangeOnline-Delete-Update-MailboxAutoReplyConfiguration
#
# Version: 3.0.0 | new-powershell-connector
#####################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set to false at start, at the end, only when no error occurs it is set to true
$outputContext.Success = $false 

# The accountReference object contains the Identification object provided in the create account call
$aRef = $actionContext.References.Account 

# Initialize default values
$c = $actionContext.Configuration

# Set debug logging
switch ($($c.isDebug)) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}

# Define configuration properties as required
$requiredConfigurationFields = @("AzureADOrganization", "AzureADTenantId", "AzureADAppId", "AzureADAppSecret")

# Used to connect to Exchange Online in an unattended scripting scenario using an App ID and App Secret to create an Access Token.
$AADOrganization = $c.AzureADOrganization
$AADTenantId = $c.AzureADTenantId
$AADAppID = $c.AzureADAppId
$AADAppSecret = $c.AzureADAppSecret

# PowerShell commands to import
$commands = @(
    "Get-User"
    , "Get-EXOMailbox"
    , "Set-MailboxAutoReplyConfiguration"
)

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
    try {
        # Verify if [aRef] has a value
        if ([string]::IsNullOrEmpty($($actionContext.References.Account))) {      
            throw 'The account reference could not be found'
        }

        # Check if required fields are available in configuration object
        $incompleteConfiguration = $false
        foreach ($requiredConfigurationField in $requiredConfigurationFields) {
            if ($requiredConfigurationField -notin $c.PsObject.Properties.Name) {
                $incompleteConfiguration = $true
                Write-Warning "Required configuration object field [$requiredConfigurationField] is missing"
            }
            elseif ([String]::IsNullOrEmpty($c.$requiredConfigurationField)) {
                $incompleteConfiguration = $true
                Write-Warning "Required configuration object field [$requiredConfigurationField] has a null or empty value"
            }
        }

        if ($incompleteConfiguration -eq $true) {
            throw "Configuration object incomplete, cannot continue."
        }

        $account = $actionContext.Data

        # Check if required fields are available in account object
        $incompleteAccount = $false
        foreach ($requiredAccountField in $requiredAccountFields) {
            if ($requiredAccountField -notin $account.PsObject.Properties.Name) {
                $incompleteAccount = $true
                Write-Warning "Required account object field [$requiredAccountField] is missing"
            }
            elseif ([String]::IsNullOrEmpty($account.$requiredAccountField)) {
                $incompleteAccount = $true
                Write-Warning "Required account object field [$requiredAccountField] has a null or empty value"
            }
        }
    
        if ($incompleteAccount -eq $true) {
            throw "Account object incomplete, cannot continue."
        }

    }
    catch {
        $ex = $PSItem
        $outputContext.AuditLogs.Add([PSCustomObject]@{
                Action  = "DeleteAccount"
                Message = "$($ex.Exception.Message)"
                IsError = $true
            })

        throw $_
    }
 
    try {           
        # Import module
        $moduleName = "ExchangeOnlineManagement"

        # If module is imported say that and do nothing
        if (Get-Module -Verbose:$false | Where-Object { $_.Name -eq $ModuleName }) {
            Write-Verbose "Module [$ModuleName] is already imported."
        }
        else {
            # If module is not imported, but available on disk then import
            if (Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -eq $ModuleName }) {
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
        $outputContext.AuditLogs.Add([PSCustomObject]@{
                Action  = "DeleteAccount"
                Message = "Error importing module [$ModuleName]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        throw "Error importing module [$ModuleName]"
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
        
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType "application/x-www-form-urlencoded" -UseBasicParsing:$true
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
        $outputContext.AuditLogs.Add([PSCustomObject]@{
                Action  = "DeleteAccount"
                Message = "Error connecting to Exchange Online. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })

        # Skip further actions, as this is a critical error
        throw "Error connecting to Exchange Online"
    }

    ## Get Exchange Online Mailbox
    try {
        Write-Verbose "Querying EXO mailbox [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
            
        $mailbox = Get-EXOMailbox -Identity $aRef.Guid -ErrorAction Stop

        if (($mailbox | Measure-Object).Count -eq 0) {
            throw "Could not find a EXO mailbox [$($aRef.UserPrincipalName) ($($aRef.Guid))]" 
        }

        $outputContext.AuditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Successfully queried EXO mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]"
                IsError = $false
            })
    }
    catch { 
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex

        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        $outputContext.AuditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Error querying EXO mailbox [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })
    
        # Skip further actions, as this is a critical error
        throw "Error querying EXO mailbox"
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

        if (-Not($actionContext.DryRun -eq $true)) {
            $updateMailbox = Set-MailboxAutoReplyConfiguration  @mailboxSplatParams -ErrorAction Stop

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = "DeleteAccount"
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
        $outputContext.AuditLogs.Add([PSCustomObject]@{
                Action  = "DeleteAccount"
                Message = "Error updating autoreply configuration for mailbox [$($aRef.userPrincipalName) ($($aRef.Guid))]: $($mailboxSplatParams | ConvertTo-Json). Error Message: $($errorMessage.AuditErrorMessage)"
                IsError = $True
            })
    }
}
catch {
    Write-Verbose $_
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
    $outputContext.AccountReference = $aRef
    $outputContext.Data = $account
}