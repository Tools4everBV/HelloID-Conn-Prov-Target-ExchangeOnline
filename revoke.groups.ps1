#####################################################
# HelloID-Conn-Prov-Target-ExchangeOnline-RevokePermission-Group
#
# Version: 2.0.0
#####################################################
# Initialize default values
$c = $configuration | ConvertFrom-Json
# The accountReference object contains the Identification object provided in the account create call
$aRef = $accountReference | ConvertFrom-Json
# The permissionReference object contains the Identification object provided in the retrieve permissions call
$pRef = $permissionReference | ConvertFrom-Json
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
    , "Remove-DistributionGroupMember"
)

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
    
    # Revoke Exchange Online Groupmembership
    try {
        Write-Verbose "Revoking permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"           

        $dgSplatParams = @{
            Identity                        = $pRef.id
            Member                          = $aRef.Guid
            BypassSecurityGroupManagerCheck = $true
        }

        if ($dryRun -eq $false) {
            $removeDGMember = Remove-DistributionGroupMember @dgSplatParams -Confirm:$false -ErrorAction Stop

            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "Successfully revoked permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                    IsError = $false
                })
        }
        else {
            Write-Warning "DryRun: would revoke permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
        }
    }
    catch {
        $ex = $PSItem
        $errorMessage = Get-ErrorMessage -ErrorObject $ex
        
        Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
        if ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Management.Tasks.MemberNotFoundException*") {
            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "User [$($aRef.UserPrincipalName) ($($aRef.Guid))] isn't a member of the group [$($pRef.Name) ($($pRef.id))]. Skipped revoke of permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                    IsError = $false
                }
            )
        }
        elseif ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($pRef.id)*") {
            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "Group [$($pRef.Name) ($($pRef.id))] couldn't be found. Possibly no longer exists. Skipped revoke of permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                    IsError = $false
                }
            )
        }
        elseif ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($aRef.Guid)*") {
            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "User [$($aRef.UserPrincipalName) ($($aRef.Guid))] couldn't be found. Possibly no longer exists. Skipped revoke of permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                    IsError = $false
                }
            )
        }
        else {
            $auditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "Error revoking permission to group [$($pRef.Name) ($($pRef.id))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                    IsError = $True
                })
        }
    }
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($auditLogs.IsError -contains $true)) {
        $success = $true
    }

    # Send results
    $result = [PSCustomObject]@{
        Success   = $success
        AuditLogs = $auditLogs
    }

    Write-Output ($result | ConvertTo-Json -Depth 10)
}