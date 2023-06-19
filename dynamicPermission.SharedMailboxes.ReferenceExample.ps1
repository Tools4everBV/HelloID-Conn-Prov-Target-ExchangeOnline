#####################################################
# HelloID-Conn-Prov-Target-ExchangeOnline-DynamicPermissions-SharedMailboxes
#
# Version: 2.0.0
#####################################################
# Initialize default values
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json
# The accountReference object contains the Identification object provided in the account create call
$aRef = $accountReference | ConvertFrom-Json
# Operation is a script parameter which contains the action HelloID wants to perform for this permission
# It has one of the following values: "grant", "revoke", "update"
$o = $operation | ConvertFrom-Json
# The entitlementContext contains the sub permissions (Previously the $permissionReference variable)
$eRef = $entitlementContext | ConvertFrom-Json
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

$currentPermissions = @{}
foreach ($permission in $eRef.CurrentPermissions) {
    $currentPermissions[$permission.Reference.Id] = $permission.DisplayName
}

# Determine all the sub-permissions that needs to be Granted/Updated/Revoked
$subPermissions = New-Object Collections.Generic.List[PSCustomObject]

# PowerShell commands to import
$commands = @(
    , "Get-EXOMailbox"
    , "Add-MailboxPermission"
    , "Add-RecipientPermission"
    , "Remove-MailboxPermission"
    , "Remove-RecipientPermission"
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

    #region Change mapping here
    $desiredPermissions = @{}
    if ($o -ne "revoke") {
        # Example: Contract Based Logic:
        foreach ($contract in $p.Contracts) {
            Write-Verbose ("Contract in condition: {0}" -f $contract.Context.InConditions)
            if ($contract.Context.InConditions -OR ($dryRun -eq $True)) {
                # Example: department_<departmentname>
                $mailboxName = "shared_mailbox_department_" + $contract.Department.DisplayName

                # Example: title_<titlename>
                # $mailboxName = "title_" + $contract.Title.Name
            
                # Get mailbox to use objectGuid to avoid name change issues
                $filter = "DisplayName -eq '$mailboxName'"
                Write-Verbose "Querying EXO mailbox that matches filter [$($filter)]"

                $mailbox = Get-EXOMailbox -Filter $filter -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

                if (($mailbox | Measure-Object).count -eq 0) {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "No mailbox found that matches filter [$($filter)]"
                            IsError = $true
                        })
                    continue
                }
                elseif (($mailbox | Measure-Object).count -gt 1) {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Multiple mailboxes found that matches filter [$($filter)]. Please correct this so the mailboxs are unique."
                            IsError = $true
                        })
                    continue
                }

                # Add mailbox to desired permissions with the id as key and the displayname as value (use id to avoid issues with name changes and for uniqueness)
                $desiredPermissions["$($mailbox.Guid)"] = $mailbox.DisplayName
            }
        }
    
        # Example: Person Based Logic:
        # Example: location_<locationname>
        # $mailboxName = "location_" + $p.Location.Name

        # # Get mailbox to use objectGuid to avoid name change issues
        # $filter = "IsDirSynced -eq 'False' -and DisplayName -eq '$mailboxName'"
        # Write-Verbose "Querying EXO mailbox that matches filter [$($filter)]"

        # $mailbox = Get-EXOMailbox -Filter $filter -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

        # if (($mailbox | Measure-Object).count -eq 0) {
        #     $auditLogs.Add([PSCustomObject]@{
        #             # Action  = "" # Optional
        #             Message = "No mailbox found that matches filter [$($filter)]"
        #             IsError = $true
        #         })
        # }
        # elseif (($mailbox | Measure-Object).count -gt 1) {
        #     $auditLogs.Add([PSCustomObject]@{
        #             # Action  = "" # Optional
        #             Message = "Multiple mailboxes found that matches filter [$($filter)]. Please correct this so the mailboxs are unique."
        #             IsError = $true
        #         })
        # }

        # # Add mailbox to desired permissions with the id as key and the displayname as value (use id to avoid issues with name changes and for uniqueness)
        # $desiredPermissions["$($mailbox.Guid)"] = $mailbox.DisplayName
    }

    Write-Warning ("Desired Permissions: {0}" -f ($desiredPermissions.Values | ConvertTo-Json))

    Write-Warning ("Existing Permissions: {0}" -f ($eRef.CurrentPermissions.DisplayName | ConvertTo-Json))

    #region Execute
    # Compare desired with current permissions and grant permissions
    foreach ($permission in $desiredPermissions.GetEnumerator()) {
        $subPermissions.Add([PSCustomObject]@{
                DisplayName = $permission.Value
                Reference   = [PSCustomObject]@{ Id = $permission.Name }
            })

        if (-Not $currentPermissions.ContainsKey($permission.Name)) {
            # Grant Exchange Online Mailbox permission
            try {
                Write-Verbose "Granting permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"

                $FullAccessPermissionSplatParams = @{
                    Identity        = $permission.Name
                    User            = $aRef.Guid
                    AccessRights    = 'FullAccess'
                    InheritanceType = 'All'
                    AutoMapping     = $AutoMapping
                } 

                if ($dryRun -eq $false) {
                    $addFullAccessPermission = Add-MailboxPermission @FullAccessPermissionSplatParams -ErrorAction Stop

                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Successfully granted permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: would grant permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                }
            }
            catch {
                $ex = $PSItem
                $errorMessage = Get-ErrorMessage -ErrorObject $ex
                    
                Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
                $auditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "Error granting permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                        IsError = $True
                    })
            }

            try {
                Write-Verbose "Granting permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"

                $sendAsPermissionSplatParams = @{
                    Identity     = $permission.Name
                    Trustee      = $aRef.Guid
                    AccessRights = 'SendAs'
                    Confirm      = $false
                } 

                if ($dryRun -eq $false) {
                    $addSendAsPermission = Add-RecipientPermission @sendAsPermissionSplatParams -ErrorAction Stop

                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Successfully granted permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: would grant permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                }
            }
            catch {
                $ex = $PSItem
                $errorMessage = Get-ErrorMessage -ErrorObject $ex
                    
                Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
                $auditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "Error granting permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                        IsError = $True
                    })
            }

            # try {
            #     Write-Verbose "Granting permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"

            #     # Can only be assigned to mailbox (so just a user account isn't sufficient, there has to be a mailbox for the user)
            #     $SendonBehalfPermissionSplatParams = @{
            #         Identity            = $permission.Name
            #         GrantSendOnBehalfTo = @{add = "$($aRef.Guid)" }
            #         Confirm             = $false
            #     } 

                    
            #     if ($dryRun -eq $false) {
            #         $addSendonBehalfPermission = Set-Mailbox @SendonBehalfPermissionSplatParams -ErrorAction Stop

            #         $auditLogs.Add([PSCustomObject]@{
            #                 # Action  = "" # Optional
            #                 Message = "Successfully granted permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
            #                 IsError = $false
            #             })
            #     }
            #     else {
            #         Write-Warning "DryRun: would grant permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
            #     }
            # }
            # catch {
            #     $ex = $PSItem
            #     $errorMessage = Get-ErrorMessage -ErrorObject $ex
                    
            #     Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
            #     $auditLogs.Add([PSCustomObject]@{
            #             # Action  = "" # Optional
            #             Message = "Error granting permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
            #             IsError = $True
            #         })
            # }
        }
    }

    # Compare current with desired permissions and revoke permissions
    $newCurrentPermissions = @{}
    foreach ($permission in $currentPermissions.GetEnumerator()) {    
        if (-Not $desiredPermissions.ContainsKey($permission.Name) -AND $permission.Name -ne "No Mailboxes Defined") {
            # Revoke Exchange Online Mailbox permission
            try {
                Write-Verbose "Revoking permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"

                $FullAccessPermissionSplatParams = @{
                    Identity        = $permission.Name
                    User            = $aRef.Guid
                    AccessRights    = 'FullAccess'
                    InheritanceType = 'All'
                    AutoMapping     = $AutoMapping
                } 

                if ($dryRun -eq $false) {
                    $addFullAccessPermission = Add-MailboxPermission @FullAccessPermissionSplatParams -ErrorAction Stop

                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Successfully revoked permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: would revoke permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                }
            }
            catch {
                $ex = $PSItem
                $errorMessage = Get-ErrorMessage -ErrorObject $ex
                    
                Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"

                if ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($permission.Name)*") {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Mailbox [$($permission.Value) ($($permission.Name))] couldn't be found. Possibly no longer exists. Skipped revoke of permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        }
                    )
                }
                elseif ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($aRef.Guid)*") {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "User [$($aRef.UserPrincipalName) ($($aRef.Guid))] couldn't be found. Possibly no longer exists. Skipped revoke of permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        }
                    )
                }
                else {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Error Revoking permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                            IsError = $True
                        })
                }
            }

            try {
                Write-Verbose "Revoking permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"

                $sendAsPermissionSplatParams = @{
                    Identity     = $permission.Name
                    Trustee      = $aRef.Guid
                    AccessRights = 'SendAs'
                    Confirm      = $false
                } 

                if ($dryRun -eq $false) {
                    $addSendAsPermission = Add-RecipientPermission @sendAsPermissionSplatParams -ErrorAction Stop

                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Successfully revoked permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: would revoke permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                }
            }
            catch {
                $ex = $PSItem
                $errorMessage = Get-ErrorMessage -ErrorObject $ex
                    
                Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"

                if ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($permission.Name)*") {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Mailbox [$($permission.Value) ($($permission.Name))] couldn't be found. Possibly no longer exists. Skipped revoke of permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        }
                    )
                }
                elseif ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($aRef.Guid)*") {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "User [$($aRef.UserPrincipalName) ($($aRef.Guid))] couldn't be found. Possibly no longer exists. Skipped revoke of permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        }
                    )
                }
                else {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Error revoking permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                            IsError = $True
                        })
                }
            }

            try {
                Write-Verbose "Revoking permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"

                # Can only be assigned to mailbox (so just a user account isn't sufficient, there has to be a mailbox for the user)
                $SendonBehalfPermissionSplatParams = @{
                    Identity             = $permission.Name
                    revokeSendOnBehalfTo = @{add = "$($aRef.Guid)" }
                    Confirm              = $false
                } 

                    
                if ($dryRun -eq $false) {
                    $addSendonBehalfPermission = Set-Mailbox @SendonBehalfPermissionSplatParams -ErrorAction Stop

                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Successfully revoked permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: would revoke permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                }
            }
            catch {
                $ex = $PSItem
                $errorMessage = Get-ErrorMessage -ErrorObject $ex
                    
                Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
                    
                if ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($permission.Name)*") {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Mailbox [$($permission.Value) ($($permission.Name))] couldn't be found. Possibly no longer exists. Skipped revoke of permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        }
                    )
                }
                elseif ($($errorMessage.AuditErrorMessage) -like "*Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException*" -and $($errorMessage.AuditErrorMessage) -like "*$($aRef.Guid)*") {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "User [$($aRef.UserPrincipalName) ($($aRef.Guid))] couldn't be found. Possibly no longer exists. Skipped revoke of permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
                            IsError = $false
                        }
                    )
                }
                else {
                    $auditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Error Revoking permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
                            IsError = $True
                        })
                }
            }
        }
        else {
            $newCurrentPermissions[$permission.Name] = $permission.Value
        }
    }
    

    # # Update current permissions
    # # Warning! This example will grant all permissions again! Only uncomment this when this is needed (e.g. force update)
    # if ($o -eq "update") {
    #     # Grant all desired permissions, ignoring current permissions
    #     foreach ($permission in $desiredPermissions.GetEnumerator()) {
    #         $subPermissions.Add([PSCustomObject]@{
    #                 DisplayName = $permission.Value
    #                 Reference   = [PSCustomObject]@{ Id = $permission.Name }
    #             })
    
    #         if (-Not $currentPermissions.ContainsKey($permission.Name)) {
    #             # Grant Exchange Online Mailbox permission
    #             try {
    #                 Write-Verbose "Granting permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    
    #                 $FullAccessPermissionSplatParams = @{
    #                     Identity        = $permission.Name
    #                     User            = $aRef.Guid
    #                     AccessRights    = 'FullAccess'
    #                     InheritanceType = 'All'
    #                     AutoMapping     = $AutoMapping
    #                 } 
    
    #                 if ($dryRun -eq $false) {
    #                     $addFullAccessPermission = Add-MailboxPermission @FullAccessPermissionSplatParams -ErrorAction Stop
    
    #                     $auditLogs.Add([PSCustomObject]@{
    #                             # Action  = "" # Optional
    #                             Message = "Successfully granted permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    #                             IsError = $false
    #                         })
    #                 }
    #                 else {
    #                     Write-Warning "DryRun: would grant permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    #                 }
    #             }
    #             catch {
    #                 $ex = $PSItem
    #                 $errorMessage = Get-ErrorMessage -ErrorObject $ex
                        
    #                 Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
    #                 $auditLogs.Add([PSCustomObject]@{
    #                         # Action  = "" # Optional
    #                         Message = "Error granting permission [FullAccess] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
    #                         IsError = $True
    #                     })
    #             }
    
    #             try {
    #                 Write-Verbose "Granting permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    
    #                 $sendAsPermissionSplatParams = @{
    #                     Identity     = $permission.Name
    #                     Trustee      = $aRef.Guid
    #                     AccessRights = 'SendAs'
    #                     Confirm      = $false
    #                 } 
    
    #                 if ($dryRun -eq $false) {
    #                     $addSendAsPermission = Add-RecipientPermission @sendAsPermissionSplatParams -ErrorAction Stop
    
    #                     $auditLogs.Add([PSCustomObject]@{
    #                             # Action  = "" # Optional
    #                             Message = "Successfully granted permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    #                             IsError = $false
    #                         })
    #                 }
    #                 else {
    #                     Write-Warning "DryRun: would grant permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    #                 }
    #             }
    #             catch {
    #                 $ex = $PSItem
    #                 $errorMessage = Get-ErrorMessage -ErrorObject $ex
                        
    #                 Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
    #                 $auditLogs.Add([PSCustomObject]@{
    #                         # Action  = "" # Optional
    #                         Message = "Error granting permission [SendAs] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
    #                         IsError = $True
    #                     })
    #             }
    
    #             # try {
    #             #     Write-Verbose "Granting permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    
    #             #     # Can only be assigned to mailbox (so just a user account isn't sufficient, there has to be a mailbox for the user)
    #             #     $SendonBehalfPermissionSplatParams = @{
    #             #         Identity            = $permission.Name
    #             #         GrantSendOnBehalfTo = @{add = "$($aRef.Guid)" }
    #             #         Confirm             = $false
    #             #     } 
    
                        
    #             #     if ($dryRun -eq $false) {
    #             #         $addSendonBehalfPermission = Set-Mailbox @SendonBehalfPermissionSplatParams -ErrorAction Stop
    
    #             #         $auditLogs.Add([PSCustomObject]@{
    #             #                 # Action  = "" # Optional
    #             #                 Message = "Successfully granted permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    #             #                 IsError = $false
    #             #             })
    #             #     }
    #             #     else {
    #             #         Write-Warning "DryRun: would grant permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]"
    #             #     }
    #             # }
    #             # catch {
    #             #     $ex = $PSItem
    #             #     $errorMessage = Get-ErrorMessage -ErrorObject $ex
                        
    #             #     Write-Verbose "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($errorMessage.VerboseErrorMessage)"
    #             #     $auditLogs.Add([PSCustomObject]@{
    #             #             # Action  = "" # Optional
    #             #             Message = "Error granting permission [SendonBehalf] to mailbox [$($permission.Value) ($($permission.Name))] for user [$($aRef.UserPrincipalName) ($($aRef.Guid))]. Error Message: $($errorMessage.AuditErrorMessage)"
    #             #             IsError = $True
    #             #         })
    #             # }
    #         }
    #     }
    # }

    # Handle case of empty defined dynamic permissions.  Without this the entitlement will error.
    if ($o -match "update|grant" -AND $subPermissions.count -eq 0) {
        $subPermissions.Add([PSCustomObject]@{
                DisplayName = "No Mailboxes Defined"
                Reference   = [PSCustomObject]@{ Id = "No Mailboxes Defined" }
            })
    }
}
#endregion Execute
finally { 
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($auditLogs.IsError -contains $true)) {
        $success = $true
    }

    #region Build up result
    $result = [PSCustomObject]@{
        Success        = $success
        SubPermissions = $subPermissions
        AuditLogs      = $auditLogs
    }
    Write-Output ($result | ConvertTo-Json -Depth 10)
    #endregion Build up result
}