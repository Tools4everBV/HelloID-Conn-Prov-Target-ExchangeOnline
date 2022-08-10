#region Initialize default properties
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json
$m = $manager | ConvertFrom-Json
$aRef = $accountReference | ConvertFrom-Json
$mRef = $managerAccountReference | ConvertFrom-Json

# The permissionReference object contains the Identification object provided in the retrieve permissions call
$pRef = $permissionReference | ConvertFrom-Json

$success = $True
$auditLogs = [Collections.Generic.List[PSCustomObject]]::new()

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Used to connect to Exchange Online using user credentials (MFA not supported).
$Domain = $c.Domain
$Username = $c.Username
$Password = $c.Password

$autoMapping = $true

# Troubleshooting
# $aRef = @{
#     Guid = "ae71715a-2964-4ce6-844a-b684d61aa1e5"
#     UserPrincipalName = "user@enyoi.onmicrosoft.com"
# }
# $dryRun = $false

#region functions
# Write functions logic here
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
        $sessionObject = Get-PSSession -ComputerName $env:computername -Name $PSSessionName -ErrorAction stop
        if ($null -eq $sessionObject) {
            # Due to some inconsistency, the Get-PSSession does not always throw an error  
            throw "The command cannot find a PSSession that has the name '$PSSessionName'."
        }
        # To Avoid using mutliple sessions at the same time.
        if ($sessionObject.length -gt 1) {
            remove-pssession -Id ($sessionObject.id | Sort-Object | select-object -first 1)
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
            Write-Verbose "Remote Powershell session is created, Name: $($sessionObject.Name), ComputerName: $($sessionObject.ComputerName)"
        }
        catch {
            throw "Couldn't created a PowerShell Session: $($_.Exception.Message)"
        }
    }
    Write-Verbose "Remote Powershell Session '$($sessionObject.Name)' State: '$($sessionObject.State)' Availability: '$($sessionObject.Availability)'"
    if ($sessionObject.Availability -eq "Busy") {
        throw "Remote session is in Use" 
    }
    Write-Output $sessionObject
}
#endregion functions

try {
    if ($dryRun -eq $false) {
        $remoteSession = Set-PSSession -PSSessionName 'HelloID_Prov_Exchange_Online'
        Connect-PSSession $remoteSession | out-null                                                                            

        # if it does not exist create new session to exchange online in remote session     
        $createSessionResult = Invoke-Command -Session $remoteSession -ScriptBlock {
            # Set TLS to accept TLS, TLS 1.1 and TLS 1.2
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

            # Create array for logging since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands
            $verboseLogs = [System.Collections.ArrayList]::new()
            $informationLogs = [System.Collections.ArrayList]::new()
            $warningLogs = [System.Collections.ArrayList]::new()
            $errorLogs = [System.Collections.ArrayList]::new()
                
            # Import module
            $moduleName = "ExchangeOnlineManagement"
            $commands = @(
                "Get-User",
                "Get-DistributionGroup",
                "Add-DistributionGroupMember",
                "Remove-DistributionGroupMember",
                "Get-EXOMailbox",
                "Add-MailboxPermission",
                "Add-RecipientPermission",
                "Set-Mailbox",
                "Remove-MailboxPermission",
                "Remove-RecipientPermission"
            )

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
    
                    # Connect to Exchange Online in an unattended scripting scenario using user credentials (MFA not supported).
                    $securePassword = ConvertTo-SecureString $using:Password -AsPlainText -Force
                    $credential = [System.Management.Automation.PSCredential]::new($using:Username, $securePassword)
                    $exchangeSessionParams = @{
                        Organization     = $using:Domain
                        Credential       = $credential
                        PSSessionOption  = $remotePSSessionOption
                        CommandName      = $commands
                        ShowBanner       = $false
                        ShowProgress     = $false
                        TrackPerformance = $false
                        ErrorAction      = 'Stop'
                    }
                    $exchangeSession = Connect-ExchangeOnline @exchangeSessionParams

                    [Void]$informationLogs.Add("Successfully connected to Exchange Online")
                }
                else {
                    [Void]$verboseLogs.Add("Already connected to Exchange Online")
                }
            }
            catch {
                if (-Not [string]::IsNullOrEmpty($_.Exception.InnerExceptions)) {
                    $errorMessage = "$($_.Exception.InnerExceptions)"
                }
                else {
                    $errorMessage = "$($_.Exception.Message) $($_.ScriptStackTrace)"
                }
                [Void]$warningLogs.Add($errorMessage)
                [Void]$errorLogs.Add("Could not connect to Exchange Online, error: $_")
            }
            finally {
                $returnobject = @{
                    verboseLogs     = $verboseLogs
                    informationLogs = $informationLogs
                    warningLogs     = $warningLogs
                    errorLogs       = $errorLogs
                }
                Remove-Variable ("verboseLogs", "informationLogs", "warningLogs", "errorLogs")     
                Write-Output $returnobject 
            }
        }

        # Log the data from logging arrarys (since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands)
        $verboseLogs = $createSessionResult.verboseLogs
        foreach ($verboseLog in $verboseLogs) { Write-Verbose $verboseLog }
        $informationLogs = $createSessionResult.informationLogs
        foreach ($informationLog in $informationLogs) { Write-Information $informationLog }
        $warningLogs = $createSessionResult.warningLogs
        foreach ($warningLog in $warningLogs) { Write-Warning $warningLog }
        $errorLogs = $createSessionResult.errorLogs
        foreach ($errorLog in $errorLogs) { Write-Error $errorLog }
        if ($errorLogs.Count -ge 1) { throw }

        # Grant Exchange Online Mailbox permission
        $addExoMailboxPermission = Invoke-Command -Session $remoteSession -ScriptBlock {
            try {
                $aRef = $using:aRef
                $pRef = $using:pRef

                $success = $false
                $auditLogs = [Collections.Generic.List[PSCustomObject]]::new()

                # Create array for logging since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands
                $verboseLogs = [System.Collections.ArrayList]::new()
                $informationLogs = [System.Collections.ArrayList]::new()
                $warningLogs = [System.Collections.ArrayList]::new()
                $errorLogs = [System.Collections.ArrayList]::new()

                foreach ($permission in $pRef.Permissions) {
                    try {
                        switch ($permission) {
                            "Full Access" {
                                [Void][Void]$verboseLogs.Add("Granting permission FullAccess to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))")
                                # No error is thrown when user already has permission
                                $addFAPermission = Add-MailboxPermission -Identity $pRef.id -AccessRights FullAccess -InheritanceType All -AutoMapping:$AutoMapping -User $aRef.Guid -ErrorAction Stop
                                [Void]$verboseLogs.Add("Successfully granted permission FullAccess to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))")

                                $success = $true
                                $auditLogs.Add([PSCustomObject]@{
                                        Action  = "GrantPermission"
                                        Message = "Successfully granted permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                                        IsError = $false
                                    }
                                )
                            }
                            "Send As" {
                                [Void]$verboseLogs.Add("Granting permission SendAs to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))")
                                # No error is thrown when user already has permission
                                $addSAPermission = Add-RecipientPermission -Identity $pRef.id -AccessRights SendAs -Confirm:$false -Trustee $aRef.Guid -ErrorAction Stop
                                [Void]$verboseLogs.Add("Successfully granted permission SendAs to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))")

                                $success = $true
                                $auditLogs.Add([PSCustomObject]@{
                                        Action  = "GrantPermission"
                                        Message = "Successfully granted permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                                        IsError = $false
                                    }
                                )
                            }
                            "Send on Behalf" {
                                [Void]$verboseLogs.Add("Granting permission SendonBehalf to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))")
                                # No error is thrown when user already has permission
                                # Can only be assigned to mailbox (so just  a user account isn't sufficient, there has to be a mailbox for the user)
                                $addSoBPermission = Set-Mailbox -Identity $pRef.id -GrantSendOnBehalfTo @{add = "$($aRef.Guid)" } -Confirm:$false -ErrorAction Stop
                                [Void]$verboseLogs.Add("Successfully granted permission SendonBehalf to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))")

                                $success = $true
                                $auditLogs.Add([PSCustomObject]@{
                                        Action  = "GrantPermission"
                                        Message = "Successfully granted permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                                        IsError = $false
                                    }
                                )
                            }
                        }
                    }
                    catch {
                        if ($_ -like "*object '$($pRef.id)' couldn't be found*") {
                            [Void]$warningLogs.Add("Mailbox $($pRef.Name) ($($pRef.id)) couldn't be found. Possibly no longer exists. Skipping action")
                            $success = $true
                            $auditLogs.Add([PSCustomObject]@{
                                    Action  = "GrantPermission"
                                    Message = "Successfully granted permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                                    IsError = $false
                                }
                            )
                        }
                        elseif ($_ -like "*User or group ""$($aRef.Guid)"" wasn't found*") {
                            [Void]$warningLogs.Add("User $($aRef.UserPrincipalName) ($($aRef.Guid)) couldn't be found. Possibly no longer exists. Skipping action")
                            $success = $true
                            $auditLogs.Add([PSCustomObject]@{
                                    Action  = "GrantPermission"
                                    Message = "Successfully granted permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                                    IsError = $false
                                }
                            )
                        }
                        else {
                            # Log error for further analysis.  Contact Tools4ever Support to further troubleshoot
                            [Void]$warningLogs.Add("Error Granting permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid)). Error: $_")
                            $success = $false
                            $auditLogs.Add([PSCustomObject]@{
                                    Action  = "GrantPermission"
                                    Message = "Failed to grant permission $($permission) to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                                    IsError = $true
                                }
                            )
                        }
                    }
                }
            }
            catch {
                # Log error for further analysis.  Contact Tools4ever Support to further troubleshoot
                [Void]$warningLogs.Add("Error Granting permission $($pRef.Permissions -join ",") to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid)). Error: $_")
                $success = $false
                $auditLogs.Add([PSCustomObject]@{
                        Action  = "GrantPermission"
                        Message = "Failed to grant permission $($pRef.Permissions -join ",") to mailbox $($pRef.Name) ($($pRef.id)) for user $($aRef.UserPrincipalName) ($($aRef.Guid))"
                        IsError = $true
                    }
                )
            }
            finally {
                $returnobject = @{
                    success         = $success
                    auditLogs       = $auditLogs
                    verboseLogs     = $verboseLogs
                    informationLogs = $informationLogs
                    warningLogs     = $warningLogs
                    errorLogs       = $errorLogs
                }
                Remove-Variable ("aRef", "pRef", "success", "auditLogs", "verboseLogs", "informationLogs", "warningLogs", "errorLogs")     
                Write-Output $returnobject 
            }
        }
    }
    $success = $addExoMailboxPermission.success
    $auditLogs = $addExoMailboxPermission.auditLogs

    # Log the data from logging arrarys (since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands)
    $verboseLogs = $addExoMailboxPermission.verboseLogs
    foreach ($verboseLog in $verboseLogs) { Write-Verbose $verboseLog }
    $informationLogs = $addExoMailboxPermission.informationLogs
    foreach ($informationLog in $informationLogs) { Write-Information $informationLog }
    $warningLogs = $addExoMailboxPermission.warningLogs
    foreach ($warningLog in $warningLogs) { Write-Warning $warningLog }
    $errorLogs = $addExoMailboxPermission.errorLogs
    foreach ($errorLog in $errorLogs) { Write-Error $errorLog }
    if ($errorLogs.Count -ge 1) { throw }
}
catch {
    $auditLogs.Add([PSCustomObject]@{
            Action  = "GrantPermission"
            Message = "Failed to grant permission:  $_"
            IsError = $True
        })
    $success = $false
    Write-Warning $_
}
finally {
    Start-Sleep 1
    if ($null -ne $remoteSession) {           
        Disconnect-PSSession $remoteSession -WarningAction SilentlyContinue | out-null   # Suppress Warning: PSSession Connection was created using the EnableNetworkAccess parameter and can only be reconnected from the local computer. # to fix the warning the session must be created with a elevated prompt
        Write-Verbose "Remote Powershell Session '$($remoteSession.Name)' State: '$($remoteSession.State)' Availability: '$($remoteSession.Availability)'"
    }      
}


#build up result
$result = [PSCustomObject]@{ 
    Success   = $success
    AuditLogs = $auditLogs
    # Account   = [PSCustomObject]@{ }
}

Write-Output $result | ConvertTo-Json -Depth 10