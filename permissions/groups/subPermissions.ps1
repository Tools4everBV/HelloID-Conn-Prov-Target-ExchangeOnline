#####################################################
# HelloID-Conn-Prov-Target-ExchangeOnline-subPermissions-Groups
# Grant and Revoke groupmemberships from account
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

# PowerShell commands to import
$commands = @(
    , "Get-DistributionGroup"
    , "Add-DistributionGroupMember"
    , "Remove-DistributionGroupMember"
)

# Determine all the sub-permissions that needs to be Granted/Updated/Revoked
$currentPermissions = @{ }
foreach ($permission in $actionContext.CurrentPermissions) {
    $currentPermissions[$permission.Reference.Id] = $permission.DisplayName
}

#region functions
function Remove-StringLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

function Get-SanitizedGroupName {
    # The names of security principal objects can contain all Unicode characters except the special LDAP characters defined in RFC 2253.
    # This list of special characters includes: a leading space a trailing space and any of the following characters: # , + " \ < > 
    # A group account cannot consist solely of numbers, periods (.), or spaces. Any leading periods or spaces are cropped.
    # https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2003/cc776019(v=ws.10)?redirectedfrom=MSDN
    # https://www.ietf.org/rfc/rfc2253.txt  
    param(
        [parameter(Mandatory = $true)][String]$Name
    )
    $newName = $name.trim()
    $newName = $newName -replace " - ", "_"
    $newName = $newName -replace "[`,~,!,#,$,%,^,&,*,(,),+,=,<,>,?,/,',`",,:,\,|,},{,.]", ""
    $newName = $newName -replace "\[", ""
    $newName = $newName -replace "]", ""
    $newName = $newName -replace " ", "_"
    $newName = $newName -replace "\.\.\.\.\.", "."
    $newName = $newName -replace "\.\.\.\.", "."
    $newName = $newName -replace "\.\.\.", "."
    $newName = $newName -replace "\.\.", "."

    # Remove diacritics
    $newName = Remove-StringLatinCharacters $newName
  
    return $newName
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

#region Get Access Token
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

    $importModuleResponse = Import-Module @importModuleSplatParams

    Write-Verbose "Imported module [$($importModuleSplatParams.Name)]"
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

    Write-Verbose "Created access token"
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

    $createExchangeSessionResponse = Connect-ExchangeOnline @createExchangeSessionSplatParams
  
    Write-Verbose "Connected to Microsoft Exchange Online"
    #endregion Connect to Microsoft Exchange Online

    #region Define desired permissions
    $actionMessage = "calculating desired permission"

    $desiredPermissions = @{}
    if (-Not($actionContext.Operation -eq "revoke")) {
        # Example: Contract Based Logic:
        foreach ($contract in $personContext.Person.Contracts) {
            Write-Verbose "Contract: $($contract.ExternalId). In condition: $($contract.Context.InConditions)"
            if ($contract.Context.InConditions -OR ($actionContext.DryRun -eq $true)) {
                $actionMessage = "querying Exchange Online Group for department: $($contract.Department | ConvertTo-Json)"

                # Get group to use objectGuid to avoid name change issues
                # Avaliable properties: https://learn.microsoft.com/en-us/powershell/exchange/filter-properties?view=exchange-ps
                $correlationField = "CustomAttribute1"
                $correlationValue = $contract.Department.ExternalId


                $getMicrosoftExchangeOnlineGroupsSplatParams = @{
                    Filter      = "$correlationField -eq '$correlationValue' -and IsDirSynced -eq 'False'"
                    ResultSize  = "Unlimited"
                    Verbose     = $false
                    ErrorAction = "Stop"
                }

                $getMicrosoftExchangeOnlineGroupsResponse = $null
                $getMicrosoftExchangeOnlineGroupsResponse = Get-DistributionGroup @getMicrosoftExchangeOnlineGroupsSplatParams
                $microsoftExchangeOnlineGroups = $getMicrosoftExchangeOnlineGroupsResponse | Select-Object -Property (@("Guid", "DisplayName", $correlationField) | Select-Object -Unique)

                if ($microsoftExchangeOnlineGroups.Guid.count -eq 0) {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "No Group found where [$($correlationField)] = [$($correlationValue)]"
                            IsError = $true
                        })
                }
                elseif ($microsoftExchangeOnlineGroups.Guid.count -gt 1) {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "Multiple Groups found where [$($correlationField)] = [$($correlationValue)]. Please correct this so the Groups are unique."
                            IsError = $true
                        })
                }
                else {
                    # Add group to desired permissions with the guid as key and the displayname as value (use id to avoid issues with name changes and for uniqueness)
                    $desiredPermissions["$($microsoftExchangeOnlineGroups.Guid)"] = "$($microsoftExchangeOnlineGroups.DisplayName)"
                }
            }
        }
    }
    #endregion Define desired permissions
  
    Write-Information ("Desired Permissions: {0}" -f ($desiredPermissions | ConvertTo-Json))
    Write-Information ("Existing Permissions: {0}" -f ($actionContext.CurrentPermissions | ConvertTo-Json))

    #region Compare current with desired permissions and revoke permissions
    $newCurrentPermissions = @{}
    foreach ($permission in $currentPermissions.GetEnumerator()) {
        if (-Not $desiredPermissions.ContainsKey($permission.Name) -AND $permission.Name -ne "No permissions defined") {
            #region Remove account from group
            try {
                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/remove-distributiongroupmember?view=exchange-ps
                $actionMessage = "revoking group [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $revokePermissionSplatParams = @{
                    Identity                        = $permission.Name
                    Member                          = $actionContext.References.Account
                    BypassSecurityGroupManagerCheck = $true
                    Confirm                         = $false
                    Verbose                         = $false
                    ErrorAction                     = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Verbose "SplatParams: $($revokePermissionSplatParams | ConvertTo-Json)"

                    $revokePermissionResponse = Remove-DistributionGroupMember @revokePermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Revoked group [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would revoke group [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
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

                if ($auditMessage -like "*Microsoft.Exchange.Management.Tasks.MemberNotFoundException*") {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Skipped $($actionMessage). Reason: User is already no longer a member."
                            IsError = $false
                        })
                }
                elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($permission.Name)*") {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Skipped $($actionMessage). Reason: Group no longer exists."
                            IsError = $false
                        })
                }
                elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Skipped $($actionMessage). Reason: User no longer exists."
                            IsError = $false
                        })
                }
                else {
                    throw $auditMessage
                }
            }
            #endregion Remove account from group
        }
        else {
            $newCurrentPermissions[$permission.Name] = $permission.Value
        }
    }
    #endregion Compare current with desired permissions and revoke permissions

    #region Compare desired with current permissions and grant permissions
    foreach ($permission in $desiredPermissions.GetEnumerator()) {
        $outputContext.SubPermissions.Add([PSCustomObject]@{
                DisplayName = $permission.Value
                Reference   = [PSCustomObject]@{ Id = $permission.Name }
            })
  
        if (-Not $currentPermissions.ContainsKey($permission.Name)) {
            #region Add account to group
            try {
                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/add-distributiongroupmember?view=exchange-ps
                $actionMessage = "granting group [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantPermissionSplatParams = @{
                    Identity                        = $permission.Name
                    Member                          = $actionContext.References.Account
                    BypassSecurityGroupManagerCheck = $true
                    Confirm                         = $false
                    Verbose                         = $false
                    ErrorAction                     = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Verbose "SplatParams: $($grantPermissionSplatParams | ConvertTo-Json)"

                    $grantPermissionResponse = Add-DistributionGroupMember @grantPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Granted group [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant group [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
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

                if ($auditMessage -like "*Microsoft.Exchange.Management.Tasks.MemberAlreadyExistsException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Skipped $($actionMessage). Reason: User is already a member."
                            IsError = $false
                        })
                }
                else {
                    throw $auditMessage
                }
            }
            #endregion Add account to group
        }
    }
    #endregion Compare desired with current permissions and grant permissions
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
            # Action = "" # Optional
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

    $deleteExchangeSessionResponse = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
  
    Write-Verbose "Disconnected from Microsoft Exchange Online"
    #endregion Disconnect from Microsoft Exchange Online

    # Handle case of empty defined dynamic permissions. Without this the entitlement will error.
    if ($actionContext.Operation -match "update|grant" -AND $outputContext.SubPermissions.count -eq 0) {
        $outputContext.SubPermissions.Add([PSCustomObject]@{
                DisplayName = "No permissions defined"
                Reference   = [PSCustomObject]@{ Id = "No permissions defined" }
            })

        Write-Warning "Skipped granting permissions for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json). Reason: No permissions defined."
    }

    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
}