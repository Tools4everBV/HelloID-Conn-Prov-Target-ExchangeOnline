#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-subPermissions-SharedMailboxes
# Grant and Revoke shared mailbox permissions (full access, send as or send on behalf) from account
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

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

    Write-Information "Created access token"
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

    #region Define desired permissions
    $actionMessage = "calculating desired permission"

    $desiredPermissions = @{}
    if (-Not($actionContext.Operation -eq "revoke")) {
        # Example: Contract Based Logic:
        foreach ($contract in $personContext.Person.Contracts) {
            Write-Information "Contract: $($contract.ExternalId). In condition: $($contract.Context.InConditions)"
            if ($contract.Context.InConditions -OR ($actionContext.DryRun -eq $true)) {
                $actionMessage = "querying Exchange Online Sharedmailbox for department: $($contract.Department | ConvertTo-Json)"
        
                # Get mailbox to use objectGuid to avoid name change issues
                # Avaliable properties: https://learn.microsoft.com/en-us/powershell/exchange/filter-properties?view=exchange-ps
                $correlationField = "CustomAttribute1"
                $correlationValue = $contract.Department.ExternalId

                $getMicrosoftExchangeOnlineSharedMailboxesSplatParams = @{
                    Properties           = (@("Guid", "DisplayName", $correlationField) | Select-Object -Unique)
                    Filter               = "$correlationField -eq '$correlationValue'"
                    RecipientTypeDetails = "SharedMailbox"
                    ResultSize           = "Unlimited"
                    Verbose              = $false
                    ErrorAction          = "Stop"
                }
        
                Write-Information "Quering ExO Mailbox where [$correlationField -eq '$correlationValue']"

                $getMicrosoftExchangeOnlineSharedMailboxesResponse = $null
                $getMicrosoftExchangeOnlineSharedMailboxesResponse = Get-EXORecipient @getMicrosoftExchangeOnlineSharedMailboxesSplatParams
                $microsoftExchangeOnlineSharedMailboxes = $getMicrosoftExchangeOnlineSharedMailboxesResponse | Select-Object -Property (@("Guid", "DisplayName", $correlationField) | Select-Object -Unique)
  
                if ($microsoftExchangeOnlineSharedMailboxes.Guid.count -eq 0) {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "No SharedMailbox found where [$($correlationField)] = [$($correlationValue)]"
                            IsError = $true
                        })
                }
                elseif ($microsoftExchangeOnlineSharedMailboxes.Guid.count -gt 1) {
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "GrantPermission"
                            Message = "Multiple SharedMailboxes found where [$($correlationField)] = [$($correlationValue)]. Please correct this so the SharedMailboxes are unique."
                            IsError = $true
                        })
                }
                else {
                    $accessRights = @("FullAccess", "SendAs") # Options: FullAccess, SendAs, SendOnBehalf
                    foreach ($accessRight in $accessRights) {
                        # Add shared mailbox to desired permissions with the desired access right + the guid as key and the displayname as value (use id to avoid issues with name changes and for uniqueness)
                        $desiredPermissions["$accessRight-$($microsoftExchangeOnlineSharedMailboxes.Guid)"] = "$accessRight-$($microsoftExchangeOnlineSharedMailboxes.DisplayName)"
                    }
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
            #region Revoke Mailbox permission
            if ($permission.Name.StartsWith("FullAccess-", [System.StringComparison]::CurrentCultureIgnoreCase)) {
                #region Revoke Full Access from account
                try {
                    $mailboxId = $permission.Name -replace 'FullAccess-', ''
                    $mailboxName = $permission.Value -replace 'FullAccess-', ''

                    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/remove-mailboxpermission?view=exchange-ps
                    $actionMessage = "revoking [FullAccess] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"
          
                    $revokeFullAccessPermissionSplatParams = @{
                        Identity        = $mailboxId
                        User            = $actionContext.References.Account
                        AccessRights    = 'FullAccess'
                        InheritanceType = 'All'
                        Confirm         = $false
                        Verbose         = $false
                        ErrorAction     = "Stop"
                    }

                    if (-Not($actionContext.DryRun -eq $true)) {
                        Write-Information "SplatParams: $($revokeFullAccessPermissionSplatParams | ConvertTo-Json)"

                        $null = Remove-MailboxPermission @revokeFullAccessPermissionSplatParams

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Revoked [FullAccess] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning "DryRun: Would revoke [FullAccess] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
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
          
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($permission.Name)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        throw $auditMessage
                    }
                }
            }
            elseif ($permission.Name.StartsWith("SendAs-", [System.StringComparison]::CurrentCultureIgnoreCase)) {
                #region Revoke Send As from account
                try {
                    $mailboxId = $permission.Name -replace 'SendAs-', ''
                    $mailboxName = $permission.Value -replace 'SendAs-', ''

                    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/remove-recipientpermission?view=exchange-ps
                    $actionMessage = "revoking [SendAs] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                    $revokeSendAsPermissionSplatParams = @{
                        Identity     = $mailboxId
                        Trustee      = $actionContext.References.Account
                        AccessRights = 'SendAs'
                        Confirm      = $false
                        Verbose      = $false
                        ErrorAction  = "Stop"
                    }

                    if (-Not($actionContext.DryRun -eq $true)) {
                        Write-Information "SplatParams: $($revokeSendAsPermissionSplatParams | ConvertTo-Json)"

                        $null = Remove-RecipientPermission @revokeSendAsPermissionSplatParams

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Revoked [SendAs] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning "DryRun: Would revoke [SendAs] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
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
        
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($permission.Name)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        throw $auditMessage
                    }
                }
                #endregion Revoke Send As from account
            }
            elseif ($permission.Name.StartsWith("SendOnBehalf-", [System.StringComparison]::CurrentCultureIgnoreCase)) {
                #region Revoke Send On Behalf from account
                try {
                    $mailboxId = $permission.Name -replace 'SendOnBehalf-', ''
                    $mailboxName = $permission.Value -replace 'SendOnBehalf-', ''

                    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
                    $actionMessage = "revoking [SendOnBehalf] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                    $revokeSendOnBehalfPermissionSplatParams = @{
                        Identity            = $mailboxId
                        GrantSendOnBehalfTo = @{remove = "$($actionContext.References.Account)" }
                        Confirm             = $false
                        Verbose             = $false
                        ErrorAction         = "Stop"
                    }

                    if (-Not($actionContext.DryRun -eq $true)) {
                        Write-Information "SplatParams: $($revokeSendOnBehalfPermissionSplatParams | ConvertTo-Json)"

                        $null = Set-Mailbox @revokeSendOnBehalfPermissionSplatParams

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Revoked [SendOnBehalf] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                                IsError = $false
                            })
                    }
                    else {
                        Write-Warning "DryRun: Would revoke [SendOnBehalf] to mailbox [$($mailboxName)] with id [$($mailboxId)] from account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
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
        
                    if ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($actionContext.References.Account)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: User no longer exists."
                                IsError = $false
                            })
                    }
                    elseif ($auditMessage -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $warningMessage -like "*$($mailboxId)*") {
                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action = "" # Optional
                                Message = "Skipped $($actionMessage). Reason: Mailbox no longer exists."
                                IsError = $false
                            })
                    }
                    else {
                        throw $auditMessage
                    }
                }
                #endregion Revoke Send On Behalf from account
            }
            #endregion Revoke Mailbox permission
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
            #region Grant Mailbox permission
            if ($permission.Name.StartsWith("FullAccess-", [System.StringComparison]::CurrentCultureIgnoreCase)) {
                #region Grant Full Access to account
                $mailboxId = $permission.Name -replace 'FullAccess-', ''
                $mailboxName = $permission.Value -replace 'FullAccess-', ''

                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/add-mailboxpermission?view=exchange-ps
                $actionMessage = "granting [FullAccess] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantFullAccessPermissionSplatParams = @{
                    Identity        = $mailboxId
                    User            = $actionContext.References.Account
                    AccessRights    = 'FullAccess'
                    InheritanceType = 'All'
                    AutoMapping     = $true
                    Confirm         = $false
                    Verbose         = $false
                    ErrorAction     = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Information "SplatParams: $($grantFullAccessPermissionSplatParams | ConvertTo-Json)"

                    $null = Add-MailboxPermission @grantFullAccessPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Granted [FullAccess] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant [FullAccess] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
                #endregion Grant Full Access to account
            }
            elseif ($permission.Name.StartsWith("SendAs-", [System.StringComparison]::CurrentCultureIgnoreCase)) {
                #region Grant Send As to account
                $mailboxId = $permission.Name -replace 'SendAs-', ''
                $mailboxName = $permission.Value -replace 'SendAs-', ''

                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/add-recipientpermission?view=exchange-ps
                $actionMessage = "granting [SendAs] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantSendAsPermissionSplatParams = @{
                    Identity     = $mailboxId
                    Trustee      = $actionContext.References.Account
                    AccessRights = 'SendAs'
                    Confirm      = $false
                    Verbose      = $false
                    ErrorAction  = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Information "SplatParams: $($grantSendAsPermissionSplatParams | ConvertTo-Json)"

                    $null = Add-RecipientPermission @grantSendAsPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Granted [SendAs] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant [SendAs] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
                #endregion Grant Send As to account
            }
            elseif ($permission.Name.StartsWith("SendOnBehalf-", [System.StringComparison]::CurrentCultureIgnoreCase)) {
                #region Grant Send On Behalf to account
                $mailboxId = $permission.Name -replace 'SendOnBehalf-', ''
                $mailboxName = $permission.Value -replace 'SendOnBehalf-', ''

                # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/set-mailbox?view=exchange-ps
                $actionMessage = "granting [SendOnBehalf] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

                $grantSendOnBehalfPermissionSplatParams = @{
                    Identity            = $mailboxId
                    GrantSendOnBehalfTo = @{add = "$($actionContext.References.Account)" }
                    Confirm             = $false
                    Verbose             = $false
                    ErrorAction         = "Stop"
                }

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Information "SplatParams: $($grantSendOnBehalfPermissionSplatParams | ConvertTo-Json)"

                    $null = Set-Mailbox @grantSendOnBehalfPermissionSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action = "" # Optional
                            Message = "Granted [SendOnBehalf] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would grant [SendOnBehalf] to mailbox [$($mailboxName)] with id [$($mailboxId)] to account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)."
                }
                #endregion Grant Send On Behalf to account
            }
            #endregion Grant Mailbox permission
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

    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
  
    Write-Information "Disconnected from Microsoft Exchange Online"
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