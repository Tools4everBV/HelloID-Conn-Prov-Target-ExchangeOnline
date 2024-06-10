#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-subPermissions-SharedMailboxes
#
# Grant and Revoke shared mailbox permissions (full access, send as or send on behalf) from account
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
    "Add-MailboxPermission"    
    , "Remove-MailboxPermission"
    , "Add-RecipientPermission"
    , "Remove-RecipientPermission"
    , "Set-Mailbox"
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
    #region Import module
    $actionMessage = "importing module"
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = "Stop"
    }
    Import-Module @importModuleSplatParams
    Write-Verbose "Imported module [$($importModuleSplatParams.Name)]"
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

    #region Define desired permissions
    $actionMessage = "calculating desired permission"

    $desiredPermissions = @{}
    if (-Not($actionContext.Operation -eq "revoke")) {
        # Example: Contract Based Logic:
        foreach ($contract in $personContext.Person.Contracts) {
            $actionMessage = "querying Exchange Online Sharedmailbox for resource: $($resource | ConvertTo-Json)"

            Write-Verbose "Contract: $($contract.ExternalId). In condition: $($contract.Context.InConditions)"
            if ($contract.Context.InConditions -OR ($actionContext.DryRun -eq $true)) {
                # Get group to use objectGuid to avoid name change issues
                # Avaliable properties: https://learn.microsoft.com/en-us/powershell/exchange/cmdlet-property-sets?view=exchange-ps#get-exomailbox-property-sets
                $correlationField = "DisplayName" # Examples "Name" "CustomAttribute1"

                # Example: department_<departmentname>
                $correlationValue = $contract.Department.DisplayName

                # Example: title_<titlename>
                # $correlationValue = "title_" + $contract.Title.Name
                
                # Sanitize group name, e.g. replace " - " with "_" or other sanitization actions 
                $correlationValue = Get-SanitizedGroupName -Name $correlationValue

                $getExOSharedMailboxesSplatParams = @{
                    Filter               = "$correlationField -eq '$correlationValue'"
                    RecipientTypeDetails = "SharedMailbox"
                    ResultSize           = "Unlimited"
                    Verbose              = $false
                    ErrorAction          = "Stop"
                }
                
                Write-Verbose "Quering ExO Mailbox where [$correlationField -eq '$correlationValue']"

                $sharedMailboxes = $null
                $sharedMailboxes = Get-EXOMailbox @getExOSharedMailboxesSplatParams
    
                if ($sharedMailboxes.Guid.count -eq 0) {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "No SharedMailbox found where [$($correlationField)] = [$($correlationValue)]"
                            IsError = $true
                        })
                }
                elseif ($sharedMailboxes.Guid.count -gt 1) {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "Multiple SharedMailboxes found where [$($correlationField)] = [$($correlationValue)]. Please correct this so the SharedMailboxes are unique."
                            IsError = $true
                        })
                }
                else {
                    # Add group to desired permissions with the id as key and the displayname as value (use id to avoid issues with name changes and for uniqueness)
                    $desiredPermissions["$($sharedMailboxes.Guid)"] = $sharedMailboxes.DisplayName
                }
            }
        }
    }
    #endregion Define desired permissions
    
    Write-Warning ("Desired Permissions: {0}" -f ($desiredPermissions.Values | ConvertTo-Json))
    Write-Warning ("Existing Permissions: {0}" -f ($actionContext.CurrentPermissions.DisplayName | ConvertTo-Json))

    #region Compare current with desired permissions and revoke permissions
    $newCurrentPermissions = @{}
    foreach ($permission in $currentPermissions.GetEnumerator()) {    
        if (-Not $desiredPermissions.ContainsKey($permission.Name) -AND $permission.Name -ne "No permissions defined") {
            #region Revoke permission from account
            # Revoke FullAccess
            $actionMessage = "revoking sharedMailbox [FullAccess] [$($permission.Value)] with id [$($permission.Name)] from account"

            $revokeFullAccessPermissionSplatParams = @{
                Identity        = $permission.Name
                User            = $actionContext.References.Account
                AccessRights    = 'FullAccess'
                InheritanceType = 'All'
                Confirm         = $false
                Verbose         = $false
                ErrorAction     = 'Stop'
            } 

            if (-Not($actionContext.DryRun -eq $true)) {
                Write-Verbose "SplatParams: $($revokeFullAccessPermissionSplatParams | ConvertTo-Json)"

                try {
                    $removeFullAccessPermission = Remove-MailboxPermission @revokeFullAccessPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "RevokePermission"
                            Message = "Revoked sharedMailbox [FullAccess] [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
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
                
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Permission.id)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning $warningMessage
                
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = $auditMessage
                                IsError = $true
                            })
                    }   
                }
            }
            else {
                Write-Warning "DryRun: Would revoke sharedMailbox [FullAccess] [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
            }

            # Revoke SendAs
            $actionMessage = "revoking sharedMailbox [SendAs] [$($permission.Value)] with id [$($permission.Name)] from account"

            $revokeSendAsPermissionSplatParams = @{
                Identity     = $permission.Name
                Trustee      = $actionContext.References.Account
                AccessRights = 'SendAs'
                Confirm      = $false
                Verbose      = $false
                ErrorAction  = 'Stop'
            } 

            if (-Not($actionContext.DryRun -eq $true)) {
                Write-Verbose "SplatParams: $($revokeSendAsPermissionSplatParams | ConvertTo-Json)"

                try {
                    $removeSendAsPermission = Remove-RecipientPermission @revokeSendAsPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "RevokePermission"
                            Message = "Revoked sharedMailbox [SendAs] [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
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
    
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Permission.id)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning $warningMessage
    
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = $auditMessage
                                IsError = $true
                            })
                    }   
                }
            }
            else {
                Write-Warning "DryRun: Would revoke sharedMailbox [SendAs] [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
            }

            # Revoke SendonBehalf
            $actionMessage = "revoking sharedMailbox [SendonBehalf] [$($permission.Value)] with id [$($permission.Name)] from account"

            $revokeSendonBehalfPermissionSplatParams = @{
                Identity            = $permission.Name
                GrantSendOnBehalfTo = @{remove = "$($actionContext.References.Account)" }
                Confirm             = $false
                Verbose             = $false
                ErrorAction         = 'Stop'
            } 

            if (-Not($actionContext.DryRun -eq $true)) {
                Write-Verbose "SplatParams: $($revokeSendonBehalfPermissionSplatParams | ConvertTo-Json)"

                try {
                    $removeSendonBehalfPermission = Set-Mailbox @revokeSendonBehalfPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "RevokePermission"
                            Message = "Revoked sharedMailbox [SendonBehalf] [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
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

                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Permission.id)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning $warningMessage

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                Action  = "RevokePermission"
                                Message = $auditMessage
                                IsError = $true
                            })
                    }   
                }
            }
            else {
                Write-Warning "DryRun: Would revoke sharedMailbox [SendonBehalf] [$($permission.Value)] with id [$($permission.Name)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
            }
            #endregion Revoke permission from account
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
            #region Grant permission to account
            # Grant FullAccess
            $actionMessage = "granting sharedMailbox [FullAccess] [$($permission.Value)] with id [$($permission.Name)] to account"

            $grantFullAccessPermissionSplatParams = @{
                Identity        = $permission.Name
                User            = $actionContext.References.Account
                AccessRights    = 'FullAccess'
                InheritanceType = 'All'
                AutoMapping     = $true
                Confirm         = $false
                Verbose         = $false
                ErrorAction     = 'Stop'
            } 

            if (-Not($actionContext.DryRun -eq $true)) {
                Write-Verbose "SplatParams: $($grantFullAccessPermissionSplatParams | ConvertTo-Json)"

                try {
                    $addFullAccessPermission = Add-MailboxPermission @grantFullAccessPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "Granted sharedMailbox [FullAccess] [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
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
                            Action  = "GrantPermission"
                            Message = $auditMessage
                            IsError = $true
                        })   
                }
            }
            else {
                Write-Warning "DryRun: Would grant sharedMailbox [FullAccess] [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
            }

            # Grant SendAs
            $actionMessage = "granting sharedMailbox [SendAs] [$($permission.Value)] with id [$($permission.Name)] to account"

            $grantSendAsPermissionSplatParams = @{
                Identity     = $permission.Name
                Trustee      = $actionContext.References.Account
                AccessRights = 'SendAs'
                Confirm      = $false
                Verbose      = $false
                ErrorAction  = 'Stop'
            } 

            if (-Not($actionContext.DryRun -eq $true)) {
                Write-Verbose "SplatParams: $($grantSendAsPermissionSplatParams | ConvertTo-Json)"

                try {
                    $addSendAsPermission = Add-RecipientPermission @grantSendAsPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "Granted sharedMailbox [SendAs] [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
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
                            Action  = "GrantPermission"
                            Message = $auditMessage
                            IsError = $true
                        })   
                }
            }
            else {
                Write-Warning "DryRun: Would grant sharedMailbox [SendAs] [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
            }

            # Grant SendonBehalf
            $actionMessage = "granting sharedMailbox [SendonBehalf] [$($permission.Value)] with id [$($permission.Name)] to account"

            $grantSendonBehalfPermissionSplatParams = @{
                Identity            = $permission.Name
                GrantSendOnBehalfTo = @{add = "$($actionContext.References.Account)" }
                Confirm             = $false
                Verbose             = $false
                ErrorAction         = 'Stop'
            } 

            if (-Not($actionContext.DryRun -eq $true)) {
                Write-Verbose "SplatParams: $($grantSendonBehalfPermissionSplatParams | ConvertTo-Json)"

                try {
                    $addSendonBehalfPermission = Set-Mailbox @grantSendonBehalfPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "RevokePermission"
                            Message = "Granted sharedMailbox [SendonBehalf] [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
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
                            Action  = "GrantPermission"
                            Message = $auditMessage
                            IsError = $true
                        })   
                }
            }
            else {
                Write-Warning "DryRun: Would grant sharedMailbox [SendonBehalf] [$($permission.Value)] with id [$($permission.Name)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
            }
            #endregion Grant permission to account
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
            # Action  = "" # Optional
            Message = $auditMessage
            IsError = $true
        })
}
finally { 
    # Handle case of empty defined dynamic permissions.  Without this the entitlement will error.
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