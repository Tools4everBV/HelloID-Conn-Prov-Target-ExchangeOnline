#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-subPermissions-SharedMailboxes
# Grant and Revoke shared mailbox permissions (full access, send as or send on behalf) from account
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

#region Determine permission type
# Extract permission type from the permission definitions
$permissionType = $actionContext.References.Permission.Reference # Permission.Reference is empty in preview.
# $permissionType = "FullAccess"
Write-Information "Processing permission type: [$permissionType]"

if ([string]::IsNullOrEmpty($permissionType)) {
    throw "Permission type could not be determined from context."
}

if ($permissionType -notin @('FullAccess', 'SendAs', 'SendOnBehalf')) {
    throw "Invalid permission type [$permissionType]. Must be one of: FullAccess, SendAs, SendOnBehalf"
}
#endregion Determine permission type

# PowerShell commands to import
$commands = switch ($permissionType) {
    'FullAccess' { @("Add-MailboxPermission", "Remove-MailboxPermission") }
    'SendAs' { @("Add-RecipientPermission", "Remove-RecipientPermission") }
    'SendOnBehalf' { @("Set-Mailbox") }
}

# Determine all the sub-permissions that needs to be Granted/Updated/Revoked
$currentPermissions = @{}
foreach ($permission in $actionContext.CurrentPermissions) {
    $currentPermissions[$permission.Reference.Id] = $permission.DisplayName
}

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

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param()
    try {
        $rawCertificate = [system.convert]::FromBase64String($actionContext.Configuration.AppCertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $actionContext.Configuration.AppCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions

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
    Write-Information "Imported module [ExchangeOnlineManagement]"
    #endregion Import module

    if ($actionContext.Configuration.UseCertificate -eq $true) {
        Write-Information "Connecting to Exchange Online with certificate"
        $actionMessage = "retrieving certificate"
        $certificate = Get-MSEntraCertificate

        $actionMessage = "connecting to Microsoft Exchange Online"
        $createExchangeSessionSplatParams = @{
            Organization          = $actionContext.Configuration.Organization
            AppID                 = $actionContext.Configuration.AppId
            Certificate           = $certificate
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
    }
    else {
        Write-Information "Connecting to Exchange Online with secret"
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
        $createAccessTokenResponse = Invoke-RestMethod @createAccessTokenSplatParams
        Write-Information "Created access token."

        $actionMessage = "connecting to Microsoft Exchange Online"
        $createExchangeSessionSplatParams = @{
            Organization          = $actionContext.Configuration.Organization
            AppID                 = $actionContext.Configuration.AppId
            AccessToken           = $createAccessTokenResponse.access_token
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
    }

    #region Define desired permissions
    $actionMessage = "calculating desired permission for type [$permissionType]"
    $desiredPermissions = @{}
    
    if (-Not($actionContext.Operation -eq "revoke")) {
        foreach ($contract in $personContext.Person.Contracts) {
            Write-Information "Contract: $($contract.ExternalId). In condition: $($contract.Context.InConditions)"
            if ($contract.Context.InConditions -OR ($actionContext.DryRun -eq $true)) {
                $actionMessage = "querying Exchange Online Sharedmailbox for department: $($contract.Department | ConvertTo-Json)"
                
                $correlationField = "CustomAttribute1"
                $correlationValue = $contract.Department.ExternalId

                $getMicrosoftExchangeOnlineSharedMailboxesSplatParams = @{
                    Properties           = @("Guid", "DisplayName", $correlationField)
                    Filter               = "$correlationField -eq '$correlationValue'"
                    RecipientTypeDetails = "SharedMailbox"
                    ResultSize           = "Unlimited"
                    Verbose              = $false
                    ErrorAction          = "Stop"
                }
        
                Write-Information "Querying ExO Mailbox where [$correlationField -eq '$correlationValue']"
                $getMicrosoftExchangeOnlineSharedMailboxesResponse = Get-EXORecipient @getMicrosoftExchangeOnlineSharedMailboxesSplatParams
                $microsoftExchangeOnlineSharedMailboxes = $getMicrosoftExchangeOnlineSharedMailboxesResponse | Select-Object Guid, DisplayName, $correlationField
  
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
                    # Add ONLY the current permission type to desired permissions
                    $desiredPermissions["$($microsoftExchangeOnlineSharedMailboxes.Guid)"] = @{
                        MailboxId   = $microsoftExchangeOnlineSharedMailboxes.Guid
                        MailboxName = $microsoftExchangeOnlineSharedMailboxes.DisplayName
                        Permission  = $permissionType
                    }
                }
            }
        }
    }
    #endregion Define desired permissions
  
    Write-Information ("Desired Permissions: {0}" -f ($desiredPermissions | ConvertTo-Json))
    Write-Information ("Current Permissions: {0}" -f ($currentPermissions | ConvertTo-Json))

    #region Compare current with desired permissions and revoke permissions
    foreach ($permission in $currentPermissions.GetEnumerator()) {
        if (-Not $desiredPermissions.ContainsKey($permission.Name) -AND $permission.Name -ne "No permissions defined") {
            switch ($permissionType) {
                'FullAccess' {
                    try {
                        $mailboxId = $permission.Name
                        $actionMessage = "revoking [$permissionType] to mailbox with id [$mailboxId] from account [$($actionContext.References.Account)]"
          
                        $revokePermissionSplatParams = @{
                            Identity        = $mailboxId
                            User            = $actionContext.References.Account
                            AccessRights    = 'FullAccess'
                            InheritanceType = 'All'
                            Confirm         = $false
                            ErrorAction     = "Stop"
                        }

                        if (-Not($actionContext.DryRun -eq $true)) {
                            Write-Information "SplatParams: $($revokePermissionSplatParams | ConvertTo-Json)"
                            $null = Remove-MailboxPermission @revokePermissionSplatParams
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Revoked [$permissionType] from account [$($actionContext.References.Account)]."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would revoke [$permissionType]"
                        }
                    }
                    catch {
                        $ex = $PSItem
                        if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
                            $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                            $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
                            $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
                        }
                        else {
                            $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                        }
          
                        if ($auditMessage -like "*ManagementObjectNotFoundException*") {
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Skipped revoke. Reason: User or Mailbox no longer exists."
                                    IsError = $false
                                })
                        }
                        else {
                            throw $auditMessage
                        }
                    }
                }
                'SendAs' {
                    try {
                        $mailboxId = $permission.Name
                        $actionMessage = "revoking [$permissionType] to mailbox with id [$mailboxId] from account [$($actionContext.References.Account)]"

                        $revokePermissionSplatParams = @{
                            Identity     = $mailboxId
                            Trustee      = $actionContext.References.Account
                            AccessRights = 'SendAs'
                            Confirm      = $false
                            ErrorAction  = "Stop"
                        }

                        if (-Not($actionContext.DryRun -eq $true)) {
                            Write-Information "SplatParams: $($revokePermissionSplatParams | ConvertTo-Json)"
                            $null = Remove-RecipientPermission @revokePermissionSplatParams
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Revoked [$permissionType] from account [$($actionContext.References.Account)]."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would revoke [$permissionType]"
                        }
                    }
                    catch {
                        $ex = $PSItem
                        if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
                            $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                            $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
                            $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
                        }
                        else {
                            $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                        }
        
                        if ($auditMessage -like "*ManagementObjectNotFoundException*") {
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Skipped revoke. Reason: User or Mailbox no longer exists."
                                    IsError = $false
                                })
                        }
                        else {
                            throw $auditMessage
                        }
                    }
                }
                'SendOnBehalf' {
                    try {
                        $mailboxId = $permission.Name
                        $actionMessage = "revoking [$permissionType] to mailbox with id [$mailboxId] from account [$($actionContext.References.Account)]"

                        $revokePermissionSplatParams = @{
                            Identity            = $mailboxId
                            GrantSendOnBehalfTo = @{remove = "$($actionContext.References.Account)" }
                            Confirm             = $false
                            ErrorAction         = "Stop"
                        }

                        if (-Not($actionContext.DryRun -eq $true)) {
                            Write-Information "SplatParams: $($revokePermissionSplatParams | ConvertTo-Json)"
                            $null = Set-Mailbox @revokePermissionSplatParams
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Revoked [$permissionType] from account [$($actionContext.References.Account)]."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would revoke [$permissionType]"
                        }
                    }
                    catch {
                        $ex = $PSItem
                        if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
                            $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                            $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
                            $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
                        }
                        else {
                            $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                        }
        
                        if ($auditMessage -like "*ManagementObjectNotFoundException*") {
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Skipped revoke. Reason: User or Mailbox no longer exists."
                                    IsError = $false
                                })
                        }
                        else {
                            throw $auditMessage
                        }
                    }
                }
            }
        }
    }
    #endregion Compare current with desired permissions and revoke permissions
  
    #region Compare desired with current permissions and grant permissions
    foreach ($permission in $desiredPermissions.GetEnumerator()) {
        $mailboxId = $permission.Value.MailboxId
        $mailboxName = $permission.Value.MailboxName

        $outputContext.SubPermissions.Add([PSCustomObject]@{
                DisplayName = "$mailboxName"
                Reference   = [PSCustomObject]@{
                    Id         = $mailboxId
                    Permission = $permissionType
                }
            })
        
        if (-Not $currentPermissions.ContainsKey($permission.Name)) {
            switch ($permissionType) {
                'FullAccess' {
                    try {
                        $actionMessage = "granting [$permissionType] to mailbox [$mailboxName] with id [$mailboxId] to account [$($actionContext.References.Account)]"

                        $grantPermissionSplatParams = @{
                            Identity        = $mailboxId
                            User            = $actionContext.References.Account
                            AccessRights    = 'FullAccess'
                            InheritanceType = 'All'
                            AutoMapping     = $true
                            Confirm         = $false
                            ErrorAction     = "Stop"
                        }

                        if (-Not($actionContext.DryRun -eq $true)) {
                            Write-Information "SplatParams: $($grantPermissionSplatParams | ConvertTo-Json)"
                            $null = Add-MailboxPermission @grantPermissionSplatParams
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Granted [$permissionType] to account [$($actionContext.References.Account)]."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would grant [$permissionType]"
                        }
                    }
                    catch {
                        $ex = $PSItem
                        if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
                            $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                            $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
                            $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
                        }
                        else {
                            $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                        }
                        throw $auditMessage
                    }
                }
                'SendAs' {
                    try {
                        $actionMessage = "granting [$permissionType] to mailbox [$mailboxName] with id [$mailboxId] to account [$($actionContext.References.Account)]"

                        $grantPermissionSplatParams = @{
                            Identity     = $mailboxId
                            Trustee      = $actionContext.References.Account
                            AccessRights = 'SendAs'
                            Confirm      = $false
                            ErrorAction  = "Stop"
                        }

                        if (-Not($actionContext.DryRun -eq $true)) {
                            Write-Information "SplatParams: $($grantPermissionSplatParams | ConvertTo-Json)"
                            $null = Add-RecipientPermission @grantPermissionSplatParams
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Granted [$permissionType] to account [$($actionContext.References.Account)]."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would grant [$permissionType]"
                        }
                    }
                    catch {
                        $ex = $PSItem
                        if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
                            $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                            $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
                            $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
                        }
                        else {
                            $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                        }
                        throw $auditMessage
                    }
                }
                'SendOnBehalf' {
                    try {
                        $actionMessage = "granting [$permissionType] to mailbox [$mailboxName] with id [$mailboxId] to account [$($actionContext.References.Account)]"

                        $grantPermissionSplatParams = @{
                            Identity            = $mailboxId
                            GrantSendOnBehalfTo = @{add = "$($actionContext.References.Account)" }
                            Confirm             = $false
                            ErrorAction         = "Stop"
                        }

                        if (-Not($actionContext.DryRun -eq $true)) {
                            Write-Information "SplatParams: $($grantPermissionSplatParams | ConvertTo-Json)"
                            $null = Set-Mailbox @grantPermissionSplatParams
                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    Message = "Granted [$permissionType] to account [$($actionContext.References.Account)]."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would grant [$permissionType]"
                        }
                    }
                    catch {
                        $ex = $PSItem
                        if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
                            $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                            $errorObj = Resolve-ExchangeOnlineError -ErrorObject $ex
                            $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
                        }
                        else {
                            $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
                        }
                        throw $auditMessage
                    }
                }
            }
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
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            Message = $auditMessage
            IsError = $true
        })
}
finally {
    $actionMessage = "disconnecting from Microsoft Exchange Online"
    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }
    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    Write-Information "Disconnected from Microsoft Exchange Online"

    if ($actionContext.Operation -match "update|grant" -AND $outputContext.SubPermissions.count -eq 0) {
        $outputContext.SubPermissions.Add([PSCustomObject]@{
                DisplayName = "No permissions defined"
                Reference   = [PSCustomObject]@{ Id = "No permissions defined" }
            })

        Write-Warning "Skipped granting permissions for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json). Reason: No permissions defined."
    }

    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
}
