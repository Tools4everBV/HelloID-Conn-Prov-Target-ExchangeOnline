#################################################
# HelloID-Conn-Prov-Target-Microsoft-Exchange-Online-Permissions-folderPermission-Grant
# Set permission for user on folder of mailbox
# PowerShell V2
#################################################
# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Define PowerShell commands to import
$commands = @(
    "Get-MailboxFolderStatistics"
    , "Set-MailboxFolderPermission"
)

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

function Convert-StringToBoolean($obj) {
    if ($obj -is [PSCustomObject]) {
        foreach ($property in $obj.PSObject.Properties) {
            $value = $property.Value
            if ($value -is [string]) {
                $lowercaseValue = $value.ToLower()
                if ($lowercaseValue -eq "true") {
                    $obj.$($property.Name) = $true
                }
                elseif ($lowercaseValue -eq "false") {
                    $obj.$($property.Name) = $false
                }
            }
            elseif ($value -is [PSCustomObject] -or $value -is [System.Collections.IDictionary]) {
                $obj.$($property.Name) = Convert-StringToBoolean $value
            }
            elseif ($value -is [System.Collections.IList]) {
                for ($i = 0; $i -lt $value.Count; $i++) {
                    $value[$i] = Convert-StringToBoolean $value[$i]
                }
                $obj.$($property.Name) = $value
            }
        }
    }
    return $obj
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

    #region Get Mailbox "Calendar" folder name
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/get-user?view=exchange-ps
    $actionMessage = "querying Mailbox [Calendar] folder name for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

    $getMicrosoftExchangeOnlineMailboxFolderStatisticsSplatParams = @{
        Identity    = $actionContext.References.Account
        FolderScope = "Calendar"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $getMicrosoftExchangeOnlineMailboxFolderStatisticsResponse = Get-MailboxFolderStatistics @getMicrosoftExchangeOnlineMailboxFolderStatisticsSplatParams
    $mailboxFolderName = ($getMicrosoftExchangeOnlineMailboxFolderStatisticsResponse | Where-Object { $_.FolderType -eq 'Calendar' }).Name
        
    Write-Information "Queried Mailbox [Calendar] folder name for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json). Result: $($mailboxFolderName | ConvertTo-Json)"
    #endregion Get Mailbox "Calendar" folder name

    #region Set Mailbox Folder Permission
    # Docs: https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailboxfolderpermission?view=exchange-ps
    $actionMessage = "setting permission for [$($actionContext.References.Permission.mailboxFolderUser)] to [$($actionContext.References.Permission.mailboxFolderAccessRight)] on the [$($mailboxFolderName)] folder of mailbox [$($actionContext.References.Account)]"

    $setMailboxFolderPermissionSplatParams = @{
        Identity     = "$($actionContext.References.Account):\$($mailboxFolderName)"
        User         = $actionContext.References.Permission.mailboxFolderUser
        AccessRights = $actionContext.References.Permission.mailboxFolderAccessRight
        Verbose      = $false
        ErrorAction  = "Stop"
    }

    Write-Information "SplatParams: $($setMailboxFolderPermissionSplatParams | ConvertTo-Json)"

    if (-Not($actionContext.DryRun -eq $true)) {
        $null = Set-MailboxFolderPermission @setMailboxFolderPermissionSplatParams

        $outputContext.AuditLogs.Add([PSCustomObject]@{
                # Action  = "" # Optional
                Message = "Set permission for [$($actionContext.References.Permission.mailboxFolderUser)] to [$($actionContext.References.Permission.mailboxFolderAccessRight)] on the [$($mailboxFolderName)] folder of mailbox [$($actionContext.References.Account)]."
                IsError = $false
            })
    }
    else {
        Write-Warning "DryRun: Would set permission for [$($actionContext.References.Permission.mailboxFolderUser)] to [$($actionContext.References.Permission.mailboxFolderAccessRight)] on the [$($mailboxFolderName)] folder of mailbox [$($actionContext.References.Account)]."
    }
    #endregion Set Mailbox Folder Permission
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
    $actionMessage = "disconnecting to Microsoft Exchange Online"

    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }

    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    
    Write-Information "Disconnected from Microsoft Exchange Online"
    #endregion Disconnect from Microsoft Exchange Online

    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }
}