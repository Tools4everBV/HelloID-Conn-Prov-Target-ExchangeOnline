#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Resources-SharedMailboxes
# Creates sharedMailboxes dynamically based on HR data
# PowerShell V2
#####################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($actionContext.Configuration.isDebug) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Determine all the sub-permissions that needs to be Granted/Updated/Revoked
$currentPermissions = @{ }
foreach ($permission in $actionContext.CurrentPermissions) {
    $currentPermissions[$permission.Reference.Id] = $permission.DisplayName
}

# PowerShell commands to import
$commands = @(
    "Get-User"
    , "Get-EXOMailbox"
    , "New-Mailbox"
    , "Set-Mailbox"
)

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
    $newName = $newName -replace " ", ""
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

#region group
# Change mapping here
# Make sure the resourceContext data is unique. Fill in the required fields after -unique
# Example: department
$resourceData = $resourceContext.SourceData | Select-Object -Unique ExternalId, DisplayName
# Example: title
# $resourceData = $resourceContext.SourceData | Select-Object -Unique ExternalId, Name
# Define correlation
# $correlationField = "displayName"
$correlationField = "CustomAttribute1"
$correlationValue = "" # Defined later in script
#endRegion group

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

    Write-Verbose "Created access token"
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

    #region Get Exchange Online Shared Mailboxes
    # Microsoft docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroup?view=exchange-ps
    $actionMessage = "querying Microsoft Exchange Online Shared Mailboxes"
    
    # Change mapping here
    $getMicrosoftExchangeOnlineSharedMailboxesSplatParams = @{
        # Filter               = "Name -like `"*Shared*`""
        Properties           = @("Guid", "DisplayName", "CustomAttribute1") # If more properties are needed please add them here
        RecipientTypeDetails = "SharedMailbox"
        ResultSize           = "Unlimited"
        Verbose              = $false
        ErrorAction          = "Stop"
    }

    $microsoftExchangeOnlineSharedMailboxes = $null
    $microsoftExchangeOnlineSharedMailboxes = Get-EXORecipient @getMicrosoftExchangeOnlineSharedMailboxesSplatParams

    # Group on correlation property to check if group exists (as correlation property has to be unique for a group)
    $microsoftExchangeOnlineSharedMailboxesGrouped = $microsoftExchangeOnlineSharedMailboxes | Group-Object $correlationField -AsHashTable -AsString

    Write-Information "Queried Microsoft Exchange Online Shared Mailboxes. Result count: $(($microsoftExchangeOnlineSharedMailboxes | Measure-Object).Count)"
    #endregion Get Microsoft Exchange Online Shared Mailboxes

    foreach ($resource in $resourceData) {
        $actionMessage = "querying sharedMailbox for resource: $($resource | ConvertTo-Json)"
        
        # Change mapping here
        # Example: department_<departmentname>
        # $groupName = "department_" + $resource.DisplayName
        $displayName = $resource.DisplayName

        # write-warning "primarySmtpAddress [$primarySmtpAddress]"
        # Example: title_<titlename>
        # $groupName = "title_" + $resource.Name

        # Determine primarySmtpAddress
        $primarySmtpAddress = $resource.DisplayName
        $domain = '@yourdomain.com'
        $primarySmtpAddress = Get-SanitizedGroupName -Name $primarySmtpAddress
        $primarySmtpAddress = $primarySmtpAddress + $domain
        $primarySmtpAddress = $primarySmtpAddress.ToLower()

        # Sanitize group name, e.g. replace " - " with "_" or other sanitization actions 
        # $groupName = Get-SanitizedGroupName -Name $groupName
       
        $correlationValue = $resource.ExternalId

        $correlatedResource = $null
        $correlatedResource = $microsoftExchangeOnlineSharedMailboxesGrouped["$($correlationValue)"]

        #region Calulate action
        if (($correlatedResource | Measure-Object).count -eq 0) {
            $actionResource = "CreateResource"
        }
        elseif (($correlatedResource | Measure-Object).count -eq 1) {
            # Exmple how to update a resource
            # if ($displayName -eq $correlatedResource.DisplayName) {
            $actionResource = "CorrelateResource"
            # }
            # else {
            #     $actionResource = "CorrelateUpdateResource"
            # }
        }
        else {
            $actionResource = "MultipleFoundResource"
        }
        #endregion Calulate action

        #region Process
        switch ($actionResource) {
            "CreateResource" {
                #region Create group
                $actionMessage = "creating sharedMailbox for resource: $($resource | ConvertTo-Json)"

                $createSharedMailboxSplatParams = @{
                    Shared             = $true
                    Name               = $displayName
                    PrimarySmtpAddress = $primarySmtpAddress
                    Verbose            = $false
                    ErrorAction        = "Stop"
                }

                Write-Verbose "SplatParams: $($createSharedMailboxSplatParams | ConvertTo-Json)"

                if (-Not($actionContext.DryRun -eq $true)) {                    
                    $response = New-Mailbox @createSharedMailboxSplatParams

                    # Change mapping here
                    # Set-Mailbox because CustomAttribute1 cannot be set with the new-mailbox command
                    $updateSharedMailboxSplatParams = @{
                        Identity         = $response.ExternalDirectoryObjectId 
                        CustomAttribute1 = $correlationValue
                        Verbose          = $false
                        ErrorAction      = "Stop"
                    }
                        
                    $null = Set-Mailbox @updateSharedMailboxSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            Action  = "CreateResource"
                            Message = "Created sharedMailbox with displayName [$($displayName)] with id [$($response.ExternalDirectoryObjectId)]."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would create sharedMailbox with displayName [$($displayName)] for resource: $($resource | ConvertTo-Json)."
                }
                #endregion Create group

                break
            }

            "CorrelateResource" {
                #region Correlate group
                $actionMessage = "correlating to sharedMailbox for resource: $($resource | ConvertTo-Json)"

                Write-Verbose "Correlated to sharedMailbox with id [$($correlatedResource.ExternalDirectoryObjectId)] and displayName [$($correlatedResource.DisplayName)] on [$($correlationField)] = [$($correlationValue)]."
                #endregion Correlate group

                break
            }

            # Exmple how to update a resource
            # "CorrelateUpdateResource" {
            #     #region Correlate update group
            #     $actionMessage = "updating to sharedMailbox for resource: $($resource | ConvertTo-Json)"   

            #     $updateSharedMailboxSplatParams = @{
            #         Identity    = $correlatedResource.ExternalDirectoryObjectId
            #         DisplayName = $displayName
            #         Name        = $displayName
            #     }
                    
            #     Write-Verbose "SplatParams: $($updateSharedMailboxSplatParams | ConvertTo-Json)"

            #     if (-Not($actionContext.DryRun -eq $true)) {      

            #         $null = Set-Mailbox @updateSharedMailboxSplatParams

            #         $outputContext.AuditLogs.Add([PSCustomObject]@{
            #                 Action      = "CreateResource"
            #                 Message     = "Updated sharedMailbox with id [$($correlatedResource.ExternalDirectoryObjectId)]. From [$($correlatedResource.DisplayName)] to [$displayName]"
            #                 IsError     = $false
            #                 Verbose     = $false
            #                 ErrorAction = "Stop"
            #             })
            #     }
            #     else {
            #         Write-Warning "DryRun: Would update sharedMailbox with id [$($correlatedResource.ExternalDirectoryObjectId)]. From [$($correlatedResource.DisplayName)] to [$displayName]"
            #     }
            # }

            "MultipleFoundResource" {

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = "CreateResource"
                        Message = "Multiple sharedMailboxes found on [$($correlationField)] = [$($correlationValue)]. DisplayNames: [$($correlatedResource.DisplayName -join ', ')]"
                        IsError = $false
                    })

            }
        }
        #endregion Process
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

    Write-Warning $warningMessage

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            # Action  = "" # Optional
            Message = $auditMessage
            IsError = $true
        })
}
finally { 
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
}