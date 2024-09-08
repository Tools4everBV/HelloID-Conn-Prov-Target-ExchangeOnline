#####################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Resources-DistributionGroups
# Creates DistributionGroups dynamically based on HR data
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
    "Get-DistributionGroup"
    , "New-DistributionGroup"
    , "Set-DistributionGroup"
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

# Define correlation field
$correlationField = "CustomAttribute1"

#region Get Access Token
try {
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

    Write-Verbose "Created access token. Result: $($createAccessTokenResonse | ConvertTo-Json)"
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

    #region Get distribution groups
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-distributiongroup?view=exchange-ps
    $actionMessage = "querying Microsoft Exchange Online distribution groups"
    
    $getMicrosoftExchangeOnlineDistributionGroupsSplatParams = @{
        Filter      = "RecipientTypeDetails -ne 'MailUniversalSecurityGroup' -and IsDirSynced -eq 'False'"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $getMicrosoftExchangeOnlineDistributionGroupsResponse = $null
    $getMicrosoftExchangeOnlineDistributionGroupsResponse = Get-DistributionGroup @getMicrosoftExchangeOnlineDistributionGroupsSplatParams
    $microsoftExchangeOnlineDistributionGroups = $getMicrosoftExchangeOnlineDistributionGroupsResponse | Select-Object -Property (@("Guid", "DisplayName", $correlationField) | Select-Object -Unique)

    Write-Information "Queried Microsoft Exchange Online distribution groups. Result count: $(($microsoftExchangeOnlineDistributionGroups | Measure-Object).Count)"
    #endregion Get distribution groups

    #region Process resources
    # Ensure the resourceContext data is unique based on ExternalId and DisplayName
    # and always sorted in the same order (by ExternalId and DisplayName)
    $resourceData = $resourceContext.SourceData |
    Select-Object -Property ExternalId, DisplayName -Unique | # Ensure uniqueness
    Sort-Object -Property ExternalId, DisplayName # Ensure consistent order by sorting on ExternalId and then by DisplayName

    # Group on $correlationField to check if distribution group exists (as correlation property has to be unique for a distribution group)
    $microsoftExchangeOnlineDistributionGroupsGrouped = $microsoftExchangeOnlineDistributionGroups | Group-Object -Property $correlationField -AsHashTable -AsString

    foreach ($resource in $resourceData) {
        #region get distribution group for resource
        $actionMessage = "querying distribution group for resource: $($resource | ConvertTo-Json)"
 
        $correlationValue = $resource.ExternalId

        $correlatedResource = $null
        if (($microsoftExchangeOnlineDistributionGroupsGrouped | Measure-Object).Count -gt 0) {
            $correlatedResource = $microsoftExchangeOnlineDistributionGroupsGrouped["$($correlationValue)"]
        }
        #endregion get distribution group for resource
        
        #region Calulate action
        if (($correlatedResource | Measure-Object).count -eq 0) {
            $actionResource = "CreateResource"
        }
        elseif (($correlatedResource | Measure-Object).count -eq 1) {
            $actionResource = "CorrelateResource"
        }
        #endregion Calulate action

        #region Process
        switch ($actionResource) {
            "CreateResource" {
                #region Create distribution group
                # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/new-distributiongroup?view=exchange-ps
                $actionMessage = "creating distribution group for resource: $($resource | ConvertTo-Json)"

                $createDistributionGroupSplatParams = @{
                    Name               = "dg_$($resource.DisplayName)"
                    PrimarySmtpAddress = "dg_$(Get-SanitizedGroupName $resource.DisplayName)@schoutenenzn.nl"
                    Verbose            = $false
                    ErrorAction        = "Stop"
                }

                Write-Verbose "SplatParams: $($createDistributionGroupSplatParams | ConvertTo-Json)"

                if (-Not($actionContext.DryRun -eq $true)) {     
                    $createDistributionGroupResponse = New-DistributionGroup @createDistributionGroupSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Created distribution group with Name [$($createDistributionGroupSplatParams.Name)] and PrimarySmtpAddress [$($createDistributionGroupSplatParams.PrimarySmtpAddress)] with id [$($createDistributionGroupResponse.Guid)] for resource: $($resource | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would create distribution group with Name [$($createDistributionGroupSplatParams.Name)] and PrimarySmtpAddress [$($createDistributionGroupSplatParams.PrimarySmtpAddress)] with id [$($createDistributionGroupResponse.Guid)] for resource: $($resource | ConvertTo-Json)."
                }
                #endregion Create distribution group

                # Update distribution group after creation, as CustomAttribute1 cannot be set with the new-distributiongroup command
                #region Update distribution group
                # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/set-distributiongroup?view=exchange-ps
                $actionMessage = "updating [$correlationField] with [$correlationValue] for created distribution group with id [$($createDistributionGroupResponse.Guid)] for resource: $($resource | ConvertTo-Json)"

                $updateDistributionGroupSplatParams = @{
                    Identity          = $createDistributionGroupResponse.Guid 
                    $correlationField = $correlationValue
                    Verbose           = $false
                    ErrorAction       = "Stop"
                }

                Write-Verbose "SplatParams: $($updateDistributionGroupSplatParams | ConvertTo-Json)"

                if (-Not($actionContext.DryRun -eq $true)) {     
                    $updateDistributionGroupResponse = Set-DistributionGroup @updateDistributionGroupSplatParams

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                            # Action  = "" # Optional
                            Message = "Updated [$correlationField] with [$correlationValue] for created distribution group with id [$($createDistributionGroupResponse.Guid)] for resource: $($resource | ConvertTo-Json)."
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Would [$correlationField] with [$correlationValue] for created distribution group with id [$($createDistributionGroupResponse.Guid)] for resource: $($resource | ConvertTo-Json)."
                }
                #endregion Update distribution group

                break
            }

            "CorrelateResource" {
                #region Correlate distribution group
                $actionMessage = "correlating to distribution group for resource: $($resource | ConvertTo-Json)"

                if (-Not($actionContext.DryRun -eq $true)) {
                    Write-Verbose "Correlated to distribution group with id [$($correlatedResource.id)] on [$($correlationField)] = [$($correlationValue)]."
                }
                else {
                    Write-Warning "DryRun: Would correlate to distribution group with id [$($correlatedResource.id)] on [$($correlationField)] = [$($correlationValue)]."
                }
                #endregion Correlate distribution group

                break
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
    #region Disconnect from Microsoft Exchange Online
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $actionMessage = "connecting to Microsoft Exchange Online"

    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }

    $deleteExchangeSessionResponse = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    
    Write-Verbose "Disconnected from Microsoft Exchange Online"
    #endregion Disconnect from Microsoft Exchange Online

    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
}