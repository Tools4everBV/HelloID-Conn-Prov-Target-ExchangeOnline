$c = $configuration | ConvertFrom-Json

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Used to connect to Exchange Online in an unattended scripting scenario using a certificate.
# Follow the Microsoft Docs on how to set up the Azure App Registration: https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
$AADOrganization = $c.AzureADOrganization
$AADAppID = $c.AzureADAppId
$AADCertificateThumbprint = $c.AzureADCertificateThumbprint # Certificate has to be locally installed

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
    $remoteSession = Set-PSSession -PSSessionName 'HelloID_Prov_Exchange_Online'
    Connect-PSSession $remoteSession | out-null                                                                            

    # if it does not exist create new session to exchange online in remote session     
    $createSessionResult = Invoke-Command -Session $remoteSession -ScriptBlock {
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

    # Get Exchange Online groups
    $getExoGroups = Invoke-Command -Session $remoteSession -ScriptBlock {
        try {
            # Create array for logging since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands
            $verboseLogs = [System.Collections.ArrayList]::new()
            $informationLogs = [System.Collections.ArrayList]::new()
            $warningLogs = [System.Collections.ArrayList]::new()
            $errorLogs = [System.Collections.ArrayList]::new()

            [Void]$verboseLogs.Add("Searching for Exchange groups..")
            # Only get Exchange Groups (Mail-enabled Security Group of Distribution Group)
            # Do not get all groups using "Get-Group", since we cannot manage Microsoft 365 or Security Groups (they have to be managed from Azure AD) 
            # Filter Cloud-Only groups (IsDirSynced -eq 'False')
            $groups = Get-DistributionGroup -Filter "IsDirSynced -eq 'False'" -ResultSize Unlimited
            [Void]$informationLogs.Add("Finished searching for Exchange Groups. Found [$($groups.id.Count) groups]")
        }
        catch {
            throw "Could not gather Exchange groups. Error: $_"
        }
        finally {
            $returnobject = @{
                groups          = $groups
                verboseLogs     = $verboseLogs
                informationLogs = $informationLogs
                warningLogs     = $warningLogs
                errorLogs       = $errorLogs
            }
            Remove-Variable ("groups", "verboseLogs", "informationLogs", "warningLogs", "errorLogs")     
            Write-Output $returnobject 
        }
    }

    # Log the data from logging arrarys (since the "normal" Write-Information isn't sent to HelloID as another PS session performs the commands)
    $verboseLogs = $getExoGroups.verboseLogs
    foreach ($verboseLog in $verboseLogs) { Write-Verbose $verboseLog }
    $informationLogs = $getExoGroups.informationLogs
    foreach ($informationLog in $informationLogs) { Write-Information $informationLog }
    $warningLogs = $getExoGroups.warningLogs
    foreach ($warningLog in $warningLogs) { Write-Warning $warningLog }
    $errorLogs = $getExoGroups.errorLogs
    foreach ($errorLog in $errorLogs) { Write-Error $errorLog }
    if ($errorLogs.Count -ge 1) { throw }
}
catch {
    throw "Could not gather Exchange groups. Error: $_"
}
finally {
    Start-Sleep 1
    if ($null -ne $remoteSession) {           
        Disconnect-PSSession $remoteSession -WarningAction SilentlyContinue | out-null   # Suppress Warning: PSSession Connection was created using the EnableNetworkAccess parameter and can only be reconnected from the local computer. # to fix the warning the session must be created with a elevated prompt
        Write-Verbose "Remote Powershell Session '$($remoteSession.Name)' State: '$($remoteSession.State)' Availability: '$($remoteSession.Availability)'"
    }      
}


# Send results
$groups = $getExoGroups.groups
foreach ($group in $groups) {
    switch ($group.RecipientType) {
        "MailUniversalSecurityGroup" {
            $groupType = "Mail-enabled Security Group"
        }
        "MailUniversalDistributionGroup" {
            $groupType = "Distribution Group"
        }
    }
    
    $permission = @{
        DisplayName    = "$($groupType) - $($group.DisplayName)"
        Identification = @{
            Id   = $group.Guid
            Name = $group.DisplayName
        }
    }

    Write-output $permission | ConvertTo-Json -Depth 10    
}