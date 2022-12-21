#####################################################
# HelloID-Conn-Prov-Target-ExchangeOnline-Create-CorrelateUser
#
# Version: 1.2.1
#####################################################
# Initialize default values
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json
$success = $false # Set to false at start, at the end, only when no error occurs it is set to true
$auditLogs = [System.Collections.Generic.List[PSCustomObject]]::new()

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($($c.isDebug)) {
    $true { $VerbosePreference = 'Continue' }
    $false { $VerbosePreference = 'SilentlyContinue' }
}
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Used to connect to Exchange Online in an unattended scripting scenario using a certificate.
# Follow the Microsoft Docs on how to set up the Azure App Registration: https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
$AADOrganization = $c.AzureADOrganization
$AADAppID = $c.AzureADAppId
$AADCertificateThumbprint = $c.AzureADCertificateThumbprint # Certificate has to be locally installed

# PowerShell commands to import
$sessionName = 'HelloID_Prov_Exchange_Online_CRUD'
$commands = @(
    "Get-User" # Always required
    , "Get-Mailbox"
    , "Get-EXOMailbox"
    , "Set-Mailbox"
    , "Set-MailboxFolderPermission"
    , "Set-MailboxRegionalConfiguration"
)

# Change mapping here
$account = [PSCustomObject]@{
    userPrincipalName = $p.Accounts.MicrosoftActiveDirectory.userPrincipalName
}

# # Troubleshooting
# $account = [PSCustomObject]@{
#     UserPrincipalName = "user@enyoi.onmicrosoft.com"
# }
# $dryRun = $false

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
            ErrorMessage          = ''
        }
        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}

function Set-PSSession {
    <#
    .SYNOPSIS
        Get or create a "remote" Powershell session
    .DESCRIPTION
        Get or create a "remote" Powershell session at the local computer
    .EXAMPLE
        PS C:\> $remoteSession = Set-PSSession -PSSessionName ($psSessionName + $mutex.Number) # Test1
       Get or Create a "remote" Powershell session at the local computer with computername and number: Test1 And assign to a $varaible which can be used to make remote calls.
    .OUTPUTS
        $remoteSession [System.Management.Automation.Runspaces.PSSession]
    .NOTES
        Make sure you always disconnect the PSSession, otherwise the PSSession is blocked to reconnect. 
        Place the following code in the finally block to make sure the session will be disconnected
        if ($null -ne $remoteSession) {  
            Disconnect-PSSession $remoteSession 
        }
    #>
    [OutputType([System.Management.Automation.Runspaces.PSSession])]  
    param(       
        [Parameter(mandatory)]
        [string]$PSSessionName
    )
    try {       
        $sessionObject = $null              
        $sessionObject = Get-PSSession -ComputerName $env:computername -Name $PSSessionName -ErrorAction stop
        if ($null -eq $sessionObject) {
            # Due to some inconsistency, the Get-PSSession does not always throw an error  
            throw "The command cannot find a PSSession that has the name '$PSSessionName'."
        }
        # To Avoid using mutliple sessions at the same time.
        if ($sessionObject.length -gt 1) {
            Remove-PSSession -Id ($sessionObject.id | Sort-Object | Select-Object -first 1)
            $sessionObject = Get-PSSession -ComputerName $env:computername -Name $PSSessionName -ErrorAction stop
        }        
        Write-Verbose "Remote Powershell session is found, Name: $($sessionObject.Name), ComputerName: $($sessionObject.ComputerName)"
    }
    catch {
        Write-Verbose "Remote Powershell session not found: $($_)"
    }

    if ($null -eq $sessionObject) { 
        try {
            $remotePSSessionOption = New-PSSessionOption -IdleTimeout (New-TimeSpan -Minutes 5).TotalMilliseconds
            $sessionObject = New-PSSession -ComputerName $env:computername -EnableNetworkAccess:$true -Name $PSSessionName -SessionOption $remotePSSessionOption
            Write-Verbose "Successfully created new Remote Powershell session, Name: $($sessionObject.Name), ComputerName: $($sessionObject.ComputerName)"
        }
        catch {
            throw "Could not create PowerShell Session with name '$PSSessionName' at computer with name '$env:computername': $($_.Exception.Message)"
        }
    }

    Write-Verbose "Remote Powershell Session '$($sessionObject.Name)' State: '$($sessionObject.State)' Availability: '$($sessionObject.Availability)'"
    if ($sessionObject.Availability -eq "Busy") {
        throw "Remote Powershell Session '$($sessionObject.Name)' is in Use"
    }

    Write-Output $sessionObject
}
#endregion functions

try {
    $remoteSession = Set-PSSession -PSSessionName $sessionName
    Connect-PSSession $remoteSession | out-null

    try {
        # if it does not exist create new session to exchange online in remote session     
        $createSessionResult = Invoke-Command -Session $remoteSession -ScriptBlock {
            try {
                # Set TLS to accept TLS, TLS 1.1 and TLS 1.2
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

                $auditLogs = [System.Collections.Generic.List[PSCustomObject]]::new()

                $dryRun = $using:dryRun

                # Create array for logging since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands
                $verboseLogs = [System.Collections.ArrayList]::new()
                $informationLogs = [System.Collections.ArrayList]::new()
                $warningLogs = [System.Collections.ArrayList]::new()
                    
                # Import module
                $moduleName = "ExchangeOnlineManagement"
                $commands = $using:commands

                # If module is imported say that and do nothing
                if (Get-Module | Where-Object { $_.Name -eq $ModuleName }) {
                    [Void]$verboseLogs.Add("Module $ModuleName is already imported.")
                }
                else {
                    # If module is not imported, but available on disk then import
                    if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $ModuleName }) {
                        $module = Import-Module $ModuleName -Cmdlet $commands
                        [Void]$verboseLogs.Add("Imported module $ModuleName")
                    }
                    else {
                        # If the module is not imported, not available and not in the online gallery then abort
                        throw "Module $ModuleName not imported, not available. Please install the module using: Install-Module -Name $ModuleName -Force"
                    }
                }

                # Check if Exchange Connection already exists
                try {
                    $checkCmd = Get-User -ResultSize 1 -ErrorAction Stop | Out-Null
                    $connectedToExchange = $true
                }
                catch {
                    if ($_.Exception.Message -like "The term 'Get-User' is not recognized as the name of a cmdlet, function, script file, or operable program.*") {
                        $connectedToExchange = $false
                    }
                }
                
                # Connect to Exchange
                try {
                    if ($connectedToExchange -eq $false) {
                        [Void]$verboseLogs.Add("Connecting to Exchange Online..")

                        # Connect to Exchange Online in an unattended scripting scenario using a certificate thumbprint (certificate has to be locally installed).
                        $exchangeSessionParams = @{
                            Organization          = $using:AADOrganization
                            AppID                 = $using:AADAppID
                            CertificateThumbPrint = $using:AADCertificateThumbprint
                            CommandName           = $commands
                            ShowBanner            = $false
                            ShowProgress          = $false
                            TrackPerformance      = $false
                            ErrorAction           = 'Stop'
                        }
                        $exchangeSession = Connect-ExchangeOnline @exchangeSessionParams
                        
                        [Void]$informationLogs.Add("Successfully connected to Exchange Online")
                    }
                    else {
                        [Void]$informationLogs.Add("Successfully connected to Exchange Online (already connected)")
                    }
                }
                catch {
                    $ex = $PSItem
                    if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                        $errorObject = Resolve-HTTPError -Error $ex
                
                        $verboseErrorMessage = $errorObject.ErrorMessage
                
                        $auditErrorMessage = $errorObject.ErrorMessage
                    }
                
                    # If error message empty, fall back on $ex.Exception.Message
                    if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
                        $verboseErrorMessage = $ex.Exception.Message
                    }
                    if ([String]::IsNullOrEmpty($auditErrorMessage)) {
                        $auditErrorMessage = $ex.Exception.Message
                    }

                    [Void]$verboseLogs.Add("Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)")
                    $auditLogs.Add([PSCustomObject]@{
                            Action  = "CreateAccount"
                            Message = "Error connecting to Exchange Online. Error Message: $auditErrorMessage"
                            IsError = $True
                        })

                    # Clean up error variables
                    Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
                    Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
                }
            }
            finally {
                $returnobject = @{
                    auditLogs       = $auditLogs
                    verboseLogs     = $verboseLogs
                    informationLogs = $informationLogs
                    warningLogs     = $warningLogs
                }
                $returnobject.Keys | ForEach-Object { Remove-Variable $_ -ErrorAction SilentlyContinue }
                Write-Output $returnobject
            }
        }
    }
    catch {
        $ex = $PSItem
        if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
            $errorObject = Resolve-HTTPError -Error $ex

            $verboseErrorMessage = $errorObject.ErrorMessage

            $auditErrorMessage = $errorObject.ErrorMessage
        }

        # If error message empty, fall back on $ex.Exception.Message
        if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
            $verboseErrorMessage = $ex.Exception.Message
        }
        if ([String]::IsNullOrEmpty($auditErrorMessage)) {
            $auditErrorMessage = $ex.Exception.Message
        }

        Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"
        $auditLogs.Add([PSCustomObject]@{
                Action  = "CreateAccount"
                Message = "Error connecting to Exchange Online. Error Message: $auditErrorMessage"
                IsError = $True
            })

        # Clean up error variables
        Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
        Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
    }
    finally {
        $auditLogs += $createSessionResult.auditLogs

        # Log the data from logging arrays (since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands)
        $verboseLogs = $createSessionResult.verboseLogs
        foreach ($verboseLog in $verboseLogs) { Write-Verbose $verboseLog }
        $informationLogs = $createSessionResult.informationLogs
        foreach ($informationLog in $informationLogs) { Write-Information $informationLog }
        $warningLogs = $createSessionResult.warningLogs
        foreach ($warningLog in $warningLogs) { Write-Warning $warningLog }
    }

    if (-NOT($auditLogs.IsError -contains $true)) {
        try {
            # Get Exchange Online User
            $getExoUser = Invoke-Command -Session $remoteSession -ScriptBlock {
                try {
                    # Set TLS to accept TLS, TLS 1.1 and TLS 1.2
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

                    $auditLogs = [System.Collections.Generic.List[PSCustomObject]]::new()

                    $dryRun = $using:dryRun
                    $account = $using:account

                    # Create array for logging since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands
                    $verboseLogs = [System.Collections.ArrayList]::new()
                    $informationLogs = [System.Collections.ArrayList]::new()
                    $warningLogs = [System.Collections.ArrayList]::new()

                    [Void]$verboseLogs.Add("Querying user with UserPrincipalName '$($account.userPrincipalName)'")

                    if ([string]::IsNullOrEmpty($account.userPrincipalName)) { throw "No UserPrincipalName provided" }  
                    
                    $user = Get-User -Identity $account.userPrincipalName -ErrorAction Stop

                    if ($null -eq $user.Guid) { throw "Failed to return a user with UserPrincipalName '$($account.userPrincipalName)'" }

                    $aRef = @{
                        Guid              = $user.Guid
                        UserPrincipalName = $user.UserPrincipalName
                    }

                    $auditLogs.Add([PSCustomObject]@{
                            Action  = "CreateAccount"
                            Message = "Successfully queried and correlated to user $($aRef.userPrincipalName) ($($aRef.Guid))"
                            IsError = $false
                        })
                }
                catch { 
                    $ex = $PSItem
                    if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                        $errorObject = Resolve-HTTPError -Error $ex
                    
                        $verboseErrorMessage = $errorObject.ErrorMessage
                    
                        $auditErrorMessage = $errorObject.ErrorMessage
                    }
                    
                    # If error message empty, fall back on $ex.Exception.Message
                    if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
                        $verboseErrorMessage = $ex.Exception.Message
                    }
                    if ([String]::IsNullOrEmpty($auditErrorMessage)) {
                        $auditErrorMessage = $ex.Exception.Message
                    }

                    [Void]$verboseLogs.Add("Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)")
                    $auditLogs.Add([PSCustomObject]@{
                            Action  = "CreateAccount"
                            Message = "Error querying user with UserPrincipalName '$($account.userPrincipalName)'. Error Message: $auditErrorMessage"
                            IsError = $True
                        })

                    # Clean up error variables
                    Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
                    Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
                }
                finally {
                    $returnobject = @{
                        user            = $user
                        aRef            = $aRef
                        auditLogs       = $auditLogs
                        verboseLogs     = $verboseLogs
                        informationLogs = $informationLogs
                        warningLogs     = $warningLogs
                    }
                    $returnobject.Keys | ForEach-Object { Remove-Variable $_ -ErrorAction SilentlyContinue }
                    Write-Output $returnobject 
                }
            }
        }
        catch {
            $ex = $PSItem
            if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                $errorObject = Resolve-HTTPError -Error $ex
        
                $verboseErrorMessage = $errorObject.ErrorMessage
        
                $auditErrorMessage = $errorObject.ErrorMessage
            }
        
            # If error message empty, fall back on $ex.Exception.Message
            if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
                $verboseErrorMessage = $ex.Exception.Message
            }
            if ([String]::IsNullOrEmpty($auditErrorMessage)) {
                $auditErrorMessage = $ex.Exception.Message
            }
        
            Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)" 
            $auditLogs.Add([PSCustomObject]@{
                    Action  = "CreateAccount"
                    Message = "Error querying user with UserPrincipalName '$($account.userPrincipalName)'. Error Message: $auditErrorMessage"
                    IsError = $True
                })

            # Clean up error variables
            Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
            Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
        }
        finally {
            $aRef = $getExoUser.aRef
            $auditLogs += $getExoUser.auditLogs
            $user = $getExoUser.user

            # Log the data from logging arrays (since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands)
            $verboseLogs = $getExoUser.verboseLogs
            foreach ($verboseLog in $verboseLogs) { Write-Verbose $verboseLog }
            $informationLogs = $getExoUser.informationLogs
            foreach ($informationLog in $informationLogs) { Write-Information $informationLog }
            $warningLogs = $getExoUser.warningLogs
            foreach ($warningLog in $warningLogs) { Write-Warning $warningLog }
        }
    }
}
finally {
    Start-Sleep 1
    if ($null -ne $remoteSession) {
        Disconnect-PSSession $remoteSession -WarningAction SilentlyContinue | out-null # Suppress Warning: PSSession Connection was created using the EnableNetworkAccess parameter and can only be reconnected from the local computer. # to fix the warning the session must be created with a elevated prompt
        Write-Verbose "Remote Powershell Session '$($remoteSession.Name)' State: '$($remoteSession.State)' Availability: '$($remoteSession.Availability)'"
    }

    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($auditLogs.IsError -contains $true)) {
        $success = $true
    }
    
    # Send results
    $result = [PSCustomObject]@{
        Success          = $success
        AccountReference = $aRef
        AuditLogs        = $auditLogs
        Account          = $account

        # Optionally return data for use in other systems
        ExportData       = [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Guid              = $user.Guid
        }
    }

    Write-Output ($result | ConvertTo-Json -Depth 10)
}