<#
.SYNOPSIS
Export-RecipientPermissions XXXVersionStringXXX
Document, filter and compare Exchange permissions: Mailbox access rights, mailbox folder permissions, public folder permissions, send as, send on behalf, managed by, linked master accounts, forwarders, management role groups
.DESCRIPTION
Document, filter and compare Exchange permissions:
- mailbox access rights
- mailbox folder permissions
- public folder permissions
- send as
- send on behalf
- managed by
- linked master accounts
- forwarders
- management role groups

Easens the move to the cloud, as permission dependencies beyond the supported cross-premises permissions (https://docs.microsoft.com/en-us/Exchange/permissions) can easily be identified and even be represented graphically (sample code included).

Compare exports from different times to detect permission changes (sample code included).
.LINK
Github: https://github.com/GruberMarkus/Export-RecipientPermissions
.PARAMETER ExportFromOnPrem
Export from On-Prem or from Exchange Online
$true for export from on-prem, $false for export from Exchange Online
Default: $false
.PARAMETER ExchangeConnectionUriList
Server URIs to connect to
For on-prem installations, list all Exchange Server Remote PowerShell URIs the script can use
For Exchange Online, this parameter is ignored use 'https://outlook.office365.com/powershell-liveid/', or the URI specific to your cloud environment
.PARAMETER ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential
Credentials for Exchange connection
Username and password are stored as encrypted secure strings, if UseDefaultCredential is not enabled
.PARAMETER ExchangeOnlineConnectionParameters
This hashtable will be passed as parameter to Connect-ExchangeOnline
Allowed values: AppId, AzureADAuthorizationEndpointUri, BypassMailboxAnchoring, Certificate, CertificateFilePath, CertificatePassword, CertificateThumbprint, Credential, DelegatedOrganization, EnableErrorReporting, ExchangeEnvironmentName, LogDirectoryPath, LogLevel, Organization, PageSize, TrackPerformance, UseMultithreading, UserPrincipalName
Values not in the allow list are removed or replaced with values determined by the script
.PARAMETER ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal
Maximum Exchange, AD and local sessions/jobs running in parallel
Watch CPU and RAM usage, and your Exchange throttling policy
.PARAMETER RecipientProperties
Recipient properties to import.
Be aware that these properties are not queried with a simple '`Get-Recipient`', but with '`Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Recipient -ResultSize Unlimited | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties)`'.
  This way, some properties have sub-values. For example, the property .PrimarySmtpAddress has .Local, .Domain and .Address as sub-values.
These properties are available for GrantorFilter and TrusteeFilter.
Properties that are always included: 'Identity', 'DistinguishedName', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'UserFriendlyName', 'LinkedMasterAccount'
.PARAMETER GrantorFilter
Only check grantors where the filter criteria matches $true.
The variable $Grantor has all attributes defined by '`RecipientProperties`. For example:
  .DistinguishedName
  .RecipientType, .RecipientTypeDetails
  .DisplayName
  .PrimarySmtpAddress: .Local, .Domain, .Address
  .EmailAddresses: .PrefixString, .IsPrimaryAddress, .SmtpAddress, .ProxyAddressString
    This attribute is an array. Code example:
      $GrantorFilter = "foreach (`$XXXSingleSmtpAddressXXX in `$Grantor.EmailAddresses.SmtpAddress) { if (`$XXXSingleSmtpAddressXXX -iin @(
                      'addressA@example.com',
                      'addressB@example.com'
      )) { `$true; break } }"
  .UserFriendlyName: User account holding the mailbox in the "<NetBIOS domain name>\<sAMAccountName>" format
  .ManagedBy: .Rdn, .Parent, .DistinguishedName, .DomainId, .Name
    This attribute is an array. Code example:
      $GrantorFilter = "foreach (`$XXXSingleManagedByXXX in `$Grantor.ManagedBy) { if (`$XXXSingleManagedByXXX -iin @(
                          'example.com/OU1/OU2/ObjectA',
                          'example.com/OU3/OU4/ObjectB',
      )) { `$true; break } }"
  On-prem only:
    .Identity: .tostring() (CN), .DomainId, .Parent (parent CN)
    .LinkedMasterAccount: Linked Master Account in the "<NetBIOS domain name>\<sAMAccountName>" format
Set to $null or '' to define all recipients as grantors to consider
Example: "`$Grantor.primarysmtpaddress.domain -ieq 'example.com'"
Default: $null
.PARAMETER TrusteeFilter
Only report trustees where the filter criteria matches $true.
If the trustee matches a recipient, the available attributes are the same as for GrantorFilter, only the reference variable is $Trustee instead of $Grantor.
If the trustee does not match a recipient (because it no longer exists, for exampe), $Trustee is just a string. In this case, the export shows the following:
  Column "Trustee Original Identity" contains the trustee description string as reported by Exchange
  Columns "Trustee Primary SMTP" and "Trustee Display Name" are empty
Example: "`$Trustee.primarysmtpaddress.domain -ieq 'example.com'"
Default: $null
.PARAMETER ExportMailboxAccessRights
Rights set on the mailbox itself, such as "FullAccess" and "ReadAccess"
Default: $true
.PARAMETER ExportMailboxAccessRightsSelf
Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST" in German, etc.)
Default: $false
.PARAMETER ExportMailboxAccessRightsInherited
Report inherited mailbox access rights (only works on-prem)
Default: $false
.PARAMETER ExportMailboxFolderPermissions
This part of the report can take very long
Default: $false
.PARAMETER ExportMailboxFolderPermissionsAnonymous
Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)
Default: $true
.PARAMETER ExportMailboxFolderPermissionsDefault
Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)
Default: $true
.PARAMETER ExportMailboxFolderPermissionsOwnerAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.
Default: $false
.PARAMETER ExportMailboxFolderPermissionsMemberAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.
Default: $false
.PARAMETER ExportMailboxFolderPermissionsExcludeFoldertype
List of Foldertypes to ignore.
Some known folder types are: Audits, Calendar, CalendarLogging, CommunicatorHistory, Conflicts, Contacts, ConversationActions, DeletedItems, Drafts, ExternalContacts, Files, GalContacts, ImContactList, Inbox, Journal, JunkEmail, LocalFailures, Notes, Outbox, QuickContacts, RecipientCache, RecoverableItemsDeletions, RecoverableItemsPurges, RecoverableItemsRoot, RecoverableItemsVersions, Root, RssSubscription, SentItems, ServerFailures, SyncIssues, Tasks, WorkingSet, YammerFeeds, YammerInbound, YammerOutbound, YammerRoot
Default: 'audits'
.PARAMETER ExportSendAs
Export Send As permissions
Default: $true
.PARAMETER ExportSendAsSelf
Export Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST" in German, etc.)
Default: $false
.PARAMETER ExportSendOnBehalf
Export Send On Behalf permissions
Default: $true
.PARAMETER ExportManagedBy
Only for distribution groups, and not to be confused with the "Manager" attribute
Default: $true
.PARAMETER ExportLinkedMasterAccount
Export Linked Master Account
Only works on-prem
Default: $true
.PARAMETER ExportPublicFolderPermissions
Export Public Folder Permissions
This part of the report can take very long
GrantorFilter refers to the public folder content mailbox
Default: $true
.PARAMETER ExportPublicFolderPermissionsAnonymous
Report public folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)
Default: $true
.PARAMETER ExportPublicFolderPermissionsDefault
Report public folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)
Default: $true
.PARAMETER ExportPublicFolderPermissionsExcludeFoldertype
List of Foldertypes to ignore.
Some known folder types are: IPF.Appointment, IPF.Contact, IPF.Note, IPF.Task
Default: ''
.PARAMETER ExportManagementRoleGroupMembers
Export members of management role groups
GrantorFilter does not apply to the export of management role groups, but TrusteeFilter does
Default: $true
.PARAMETER ExportForwarders
Export forwarders configured on recipients
Default: $true
.PARAMETER ExportTrustees
Include all trustees in permission report file, only valid or only invalid ones
Valid trustees are trustees which can be resolved to an Exchange recipient
Valid values: 'All', 'OnlyValid', 'OnlyInvalid'
Default: 'All'
.PARAMETER ExportFile
Name (and path) of the permission report file
Default: '.\export\Export-RecipientPermissions_Result.csv'
.PARAMETER ErrorFile
Name (and path) of the error log file
Set to $null or '' to disable debugging
Default: '.\export\Export-RecipientPermissions_Error.csv',
.PARAMETER DebugFile
Name (and path) of the debug log file
Set to $null or '' to disable debugging
Default: ''
.PARAMETER UpdateInverval
Interval to update the job progress
Updates are based von recipients done, not on duration
Number must be 1 or higher, lower numbers mean bigger debug files
Default: 100
.INPUTS
None. You cannot pipe objects to Export-RecipientPermissions.
.OUTPUTS
Export-RecipientPermissions writes the current activities, warnings and error messages to the standard output stream.
.EXAMPLE
Run Export-RecipientPermissions with default values (export from Exchange Online)
PS> .\Export-RecipientPermissions.ps1
.EXAMPLE
Run Export-RecipientPermissions with default values (export from Exchange On-Prem and use credential of currently logged-on user)
PS> .\Export-RecipientPermissions.ps1 -ExportFromOnPrem $true -UseDefaultCredential $true
.NOTES
Script : Export-RecipientPermissions
Version: XXXVersionStringXXX
Web    : https://github.com/GruberMarkus/Export-RecipientPermissions
License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)
#>

[CmdletBinding(PositionalBinding = $false)]


Param(
    # Export from On-Prem or from Exchange Online
    # $true for export from on-prem
    # $false for export from Exchange Online
    [boolean]$ExportFromOnPrem = $false,


    # Server URIs to connect to
    # For on-prem installations, list all Exchange Server Remote PowerShell URIs the script can use
    # For Exchange Online use 'https://outlook.office365.com/powershell-liveid/', or the URI specific to your cloud environment
    [uri[]]$ExchangeConnectionUriList = ('https://outlook.office365.com/powershell-liveid'),


    # Credentials for Exchange connection
    #
    # Use default credential
    #   Does not store encrypted credentials in file system, but will not work with 'ExportFromOnPrem = $false'
    [boolean]$UseDefaultCredential = $false,
    #
    # Username and password are stored as encrypted secure strings
    # These credentials will replace the value of the 'Credential' entry of the XXX parameter hashtable
    [string]$ExchangeCredentialUsernameFile = '.\Export-RecipientPermissions_CredentialUsername.txt',
    [string]$ExchangeCredentialPasswordFile = '.\Export-RecipientPermissions_CredentialPassword.txt',


    # Exchange Online connection parameters
    # This hashtable will be passed as parameter to Connect-ExchangeOnline
    # Allowed values: AppId, AzureADAuthorizationEndpointUri, BypassMailboxAnchoring, Certificate, CertificateFilePath, CertificatePassword, CertificateThumbprint, Credential, DelegatedOrganization, EnableErrorReporting, ExchangeEnvironmentName, LogDirectoryPath, LogLevel, Organization, PageSize, TrackPerformance, UseMultithreading, UserPrincipalName
    # Values not in the allow list are removed or replaced with values determined by the script
    [hashtable]$ExchangeOnlineConnectionParameters = @{ Credential = $true },


    # Maximum Exchange, AD and local sessions/jobs running in parallel
    # Watch CPU and RAM usage, and your Exchange throttling policy
    [int]$ParallelJobsExchange = $ExchangeConnectionUriList.count * 3,
    [int]$ParallelJobsAD = 50,
    [int]$ParallelJobsLocal = 50,


    # Recipient properties to import
    [string[]]$RecipientProperties = @(),

    # Grantors to consider
    [string]$GrantorFilter = $null, # "`$Grantor.primarysmtpaddress.domain -ieq 'example.com'"


    # Trustees to consider
    [string]$TrusteeFilter = $null, # "`$Trustee.primarysmtpaddress.domain -ieq 'example.com'"


    # Permissions to report
    #
    # Mailbox Access Rights
    # Rights set on the mailbox itself, such as "FullAccess" and "ReadAccess"
    [boolean]$ExportMailboxAccessRights = $true,
    [boolean]$ExportMailboxAccessRightsSelf = $false, # Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST" in German, etc.)
    [boolean]$ExportMailboxAccessRightsInherited = $false, # Report inherited mailbox access rights (only works on-prem)
    #
    # Mailbox Folder Permissions
    # This part of the report can take very long
    [boolean]$ExportMailboxFolderPermissions = $false,
    [boolean]$ExportMailboxFolderPermissionsAnonymous = $true, # Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)
    [boolean]$ExportMailboxFolderPermissionsDefault = $true, # Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)
    [boolean]$ExportMailboxFolderPermissionsOwnerAtLocal = $false, # Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.
    [boolean]$ExportMailboxFolderPermissionsMemberAtLocal = $false, # Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.
    [string[]]$ExportMailboxFolderPermissionsExcludeFoldertype = ('audits'), # List of Foldertypes to ignore. Some known folder types are: Audits, Calendar, CalendarLogging, CommunicatorHistory, Conflicts, Contacts, ConversationActions, DeletedItems, Drafts, ExternalContacts, Files, GalContacts, ImContactList, Inbox, Journal, JunkEmail, LocalFailures, Notes, Outbox, QuickContacts, RecipientCache, RecoverableItemsDeletions, RecoverableItemsPurges, RecoverableItemsRoot, RecoverableItemsVersions, Root, RssSubscription, SentItems, ServerFailures, SyncIssues, Tasks, WorkingSet, YammerFeeds, YammerInbound, YammerOutbound, YammerRoot
    #
    # Send As
    [boolean]$ExportSendAs = $true,
    [boolean]$ExportSendAsSelf = $false, # Report Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST" in German, etc.)
    #
    # Send On Behalf
    [boolean]$ExportSendOnBehalf = $true,
    #
    # Managed By
    # Only for distribution groups, and not to be confused with the "Manager" attribute
    [boolean]$ExportManagedBy = $true,
    #
    # Linked Master Account
    # Only works on-prem
    [boolean]$ExportLinkedMasterAccount = $true,
    #
    # Public Folder Permissions
    # This part of the report can take very long
    [boolean]$ExportPublicFolderPermissions = $true,
    [boolean]$ExportPublicFolderPermissionsAnonymous = $true, # Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)
    [boolean]$ExportPublicFolderPermissionsDefault = $true, # Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)
    [string[]]$ExportPublicFolderPermissionsExcludeFoldertype = (''), # List of Foldertypes to ignore. Some known folder types are: IPF.Appointment, IPF.Contact, IPF.Note, IPF.Task
    #
    # Management Role Groups
    [boolean]$ExportManagementRoleGroupMembers = $true,
    #
    # Forwarders
    [boolean]$ExportForwarders = $true,


    # Include all trustees in permission report file, only valid or only invalid ones
    # Valid trustees are trustees which can be resolved to an Exchange recipient
    [ValidateSet('All', 'OnlyValid', 'OnlyInvalid')]
    $ExportTrustees = 'All',


    # Name (and path) of the permission report file
    [string]$ExportFile = '.\export\Export-RecipientPermissions_Result.csv',


    # Name (and path) of the error file
    # Set to $null or '' to disable debugging
    [string]$ErrorFile = '.\export\Export-RecipientPermissions_Error.csv',


    # Name (and path) of the debug log file
    # Set to $null or '' to disable debugging
    [string]$DebugFile = '',


    # Interval to update the job progress
    # Updates are based von recipients done, not on duration
    # Number must be 1 or higher, lower numbers mean bigger debug files
    [int][ValidateRange(1, [int]::MaxValue)]$UpdateInterval = 100
)


#
# Do not change anything from here on.
#


$ConnectExchange = {
    $Stoploop = $false
    [int]$Retrycount = 1

    if (-not $connectionUri) {
        $connectionUri = $tempConnectionUriQueue.dequeue()
    }

    Write-Verbose "Connection URI: '$connectionUri'"

    while ($Stoploop -ne $true) {
        try {
            if ($Retrycount -gt 1) {
                Write-Host "Try $($Retrycount), via '$($connectionUri)'."
            }

            if ($ExchangeSession) {
                $null = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-SecurityPrincipal -ResultSize 1 -WarningAction SilentlyContinue } -ErrorAction Stop

                Write-Verbose "Exchange session established and working on try $($RetryCount)."

                $Stoploop = $true
            } else {
                throw
            }
        } catch {
            try {
                Write-Verbose "Exchange session not established or working on try $($RetryCount)."

                if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                    Disconnect-ExchangeOnline -Confirm:$false
                    Remove-Module ExchangeOnlineManagement
                }

                if ($ExchangeSession) {
                    Remove-PSSession -Session $ExchangeSession
                }

                if ($ExportFromOnPrem -eq $true) {
                    if ($UseDefaultCredential) {
                        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication Kerberos -AllowRedirection -Name 'ExchangeSession' -ErrorAction Stop
                    } else {
                        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $ExchangeCredential -Authentication Kerberos -AllowRedirection -Name 'ExchangeSession' -ErrorAction Stop
                    }

                    $null = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True } -ErrorAction Stop
                } else {
                    if ($ExchangeOnlineConnectionParameters.ContainsKey('Credential')) {
                        $ExchangeOnlineConnectionParameters['Credential'] = $ExchangeCredential
                    }

                    $ExchangeOnlineConnectionParameters['ConnectionUri'] = $connectionUri
                    $ExchangeOnlineConnectionParameters['CommandName'] = ('Get-DistributionGroup', 'Get-DynamicDistributionGroup', 'Get-Mailbox', 'Get-MailboxFolderPermission', 'Get-MailboxFolderStatistics', 'Get-MailboxPermission', 'Get-MailPublicFolder', 'Get-Publicfolder', 'Get-PublicFolderClientPermission', 'Get-Recipient', 'Get-RecipientPermission', 'Get-SecurityPrincipal', 'Get-UnifiedGroup')

                    Import-Module '.\bin\ExchangeOnlineManagement' -Force -ErrorAction Stop
                    Connect-ExchangeOnline @ExchangeOnlineConnectionParameters
                    $ExchangeSession = Get-PSSession | Sort-Object -Property Id | Select-Object -Last 1
                }

                $null = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-SecurityPrincipal -ResultSize 1 -WarningAction SilentlyContinue } -ErrorAction Stop

                Write-Verbose "Exchange session established and working on try $($RetryCount)."

                $Stoploop = $true
            } catch {
                if ($Retrycount -lt 3) {
                    Write-Host "Exchange session could not be established in a working state to '$($connectionUri)' on try $($Retrycount)."
                    Write-Host $error[0]

                    if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                        Disconnect-ExchangeOnline -Confirm:$false
                        Remove-Module ExchangeOnlineManagement
                    }

                    if ($ExchangeSession) {
                        Remove-PSSession -Session $ExchangeSession
                    }

                    $connectionUri = $tempConnectionUriQueue.dequeue()

                    $SleepTime = (30 * $RetryCount * ($RetryCount / 2)) + 15

                    Write-Host "Trying again in $($SleepTime) seconds via '$connectionUri'."

                    Start-Sleep -Seconds $SleepTime
                    $Retrycount++
                } else {
                    throw "Exchange session could not be established in a working state on $($Retrycount) retries. Giving up."
                    $Stoploop = $true
                }
            }
        }
    }
}


try {
    Set-Location $PSScriptRoot

    if ($PSVersionTable.PSEdition -ieq 'desktop') {
        $UTF8Encoding = 'UTF8'
    } else {
        $UTF8Encoding = 'UTF8BOM'
    }
    if ($ExportFile) {
        $ExportFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportFile)
    }

    if ($ErrorFile) {
        $ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
    }

    if ($DebugFile) {
        $DebugFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DebugFile)
    }

    $ExchangeCredentialUsernameFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExchangeCredentialUsernameFile)
    $ExchangeCredentialPasswordFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExchangeCredentialPasswordFile)

    if ($DebugFile) {
        try {
            Stop-Transcript
        } catch {
        }
        $null = Start-Transcript -Path $DebugFile -Force
    }


    Clear-Host
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"


    Write-Host
    Write-Host "Script notes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host '  Script : Export-RecipientPermissions'
    Write-Host '  Version: XXXVersionStringXXX'
    Write-Host '  Web    : https://github.com/GruberMarkus/Export-RecipientPermissions'
    Write-Host "  License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)"


    Write-Host
    Write-Host "Script environment and parameters @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "  PowerShell: '$((($($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition), $($PSVersionTable.Platform), $($PSVersionTable.OS)) | Where-Object {$_}) -join "', '")'"
    Write-Host "  PowerShell bitness: $(if ([Environment]::Is64BitProcess -eq $false) {'Non-'})64-bit process on a $(if ([Environment]::Is64OperatingSystem -eq $false) {'Non-'})64-bit operating system"
    Write-Host "  Script path: '$PSCommandPath'"
    Write-Host "  PowerShell invocation: '$(($MyInvocation.Line).trimend([environment]::NewLine))'"
    Write-Host '  Parameters'
    foreach ($parameter in (Get-Command -Name $PSCommandPath).Parameters.keys) {
        Write-Host "    $($parameter): " -NoNewline

        if ((Get-Variable -Name $parameter -EA SilentlyContinue -ValueOnly) -is [hashtable]) {
            Write-Host "'$(@((Get-Variable -Name $parameter -ValueOnly).GetEnumerator() | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ', ')'"
        } else {
            Write-Host "'$((Get-Variable -Name $parameter -EA SilentlyContinue -ValueOnly) -join ', ')'"
        }
    }


    if ($ErrorFile) {
        if (Test-Path $ErrorFile) {
            Remove-Item $ErrorFile -Force
        }
        try {
            foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
                Remove-Item $JobErrorFile -Force
            }
        } catch {}
        $null = New-Item $ErrorFile -Force
        '"Timestamp";"Task";"TaskDetail";"Error"' | Out-File $ErrorFile -Encoding $UTF8Encoding -Force
    }


    if ($DebugFile) {
        foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
            Remove-Item $JobDebugFile -Force
        }
    }


    if ($ExportFile) {
        if (Test-Path $ExportFile) {
            Remove-Item $ExportFile -Force
        }
        try {
            foreach ($RecipientFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
                Remove-Item $Recipientfile -Force
            }
        } catch {}
        $null = New-Item $ExportFile -Force
        '"Grantor Primary SMTP";"Grantor Display Name";"Grantor Recipient Type";"Grantor Environment";"Folder";"Permission";"Allow/Deny";"Inherited";"InheritanceType";"Trustee Original Identity";"Trustee Primary SMTP";"Trustee Display Name";"Trustee Recipient Type";"Trustee Environment"' | Out-File $ExportFile -Encoding $UTF8Encoding -Force
    }


    $tempConnectionUriQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new(10000))
    while ($tempConnectionUriQueue.count -le 10000) {
        foreach ($ExchangeConnectionUri in $ExchangeConnectionUriList) {
            $tempConnectionUriQueue.Enqueue($ExchangeConnectionUri.AbsoluteUri)
        }
    }


    if ($ExchangeOnlineConnectionParameters) {
        $ExchangeOnlineConnectionParametersFiltered = @{}

        $ExchangeOnlineConnectionParameters.GetEnumerator() | Where-Object { $_.name -iin (
                'AppId',
                'AzureADAuthorizationEndpointUri',
                'BypassMailboxAnchoring',
                'Certificate',
                'CertificateFilePath',
                'CertificatePassword',
                'CertificateThumbprint',
                'Credential',
                'DelegatedOrganization',
                'EnableErrorReporting',
                'ExchangeEnvironmentName',
                'LogDirectoryPath',
                'LogLevel',
                'Organization',
                'PageSize',
                'TrackPerformance',
                'UseMultithreading',
                'UserPrincipalName'
            ) } | ForEach-Object {
            $ExchangeOnlineConnectionParametersFiltered.add($_.name, $_.value)
        }

        $ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParametersFiltered
        $ExchangeOnlineConnectionParametersFiltered = $null

        $ExchangeOnlineConnectionParameters.add('UseRPSSession', $true)
        $ExchangeOnlineConnectionParameters.add('ShowBanner', $false)
        $ExchangeOnlineConnectionParameters.add('ShowProgress', $false)


        if ($RecipientProperties -contains '*') {
            $RecipientProperties = @('*')
        } else {
            @('Identity', 'DistinguishedName', 'ExchangeGuid', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy') | ForEach-Object {
                if ($RecipientProperties -inotcontains $_) {
                    $RecipientProperties += $_
                }
            }
        }

        @('UserFriendlyName', 'LinkedMasterAccount', 'FileLock') | ForEach-Object {
            if ($RecipientProperties -inotcontains $_) {
                $RecipientProperties += $_
            }
        }

        if ($ExportForwarders) {
            @('ExternalEmailAddress', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward') | ForEach-Object {
                if ($RecipientProperties -inotcontains $_) {
                    $RecipientProperties += $_
                }
            }
        }

        $RecipientProperties = @($RecipientProperties | Select-Object -Unique)
    }


    # Credentials
    Write-Host
    Write-Host "Exchange credentials @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ((-not $UseDefaultCredential) -or (($ExportFromOnPrem -eq $false) -and ($ExchangeOnlineConnectionParameters.ContainsKey('Credential')))) {
        if (-not ((Test-Path $ExchangeCredentialUsernameFile) -and (Test-Path $ExchangeCredentialPasswordFile))) {
            Write-Host '  No stored credential found'
            Write-Host '    Username and password are stored as encrypted secure strings'
            Read-Host -Prompt '    Please enter username for later use (characters are masked)' -AsSecureString | ConvertFrom-SecureString | Out-File $ExchangeCredentialUsernameFile -Force -Encoding $UTF8Encoding
            Read-Host -Prompt '    Please enter password for later use (characters are masked)' -AsSecureString | ConvertFrom-SecureString | Out-File $ExchangeCredentialPasswordFile -Force -Encoding $UTF8Encoding
        }

        Write-Host '  Loading credentials encrypted as secure strings'
        Write-Host "    Username file: '$ExchangeCredentialUsernameFile'"
        Write-Host "    Password file: '$ExchangeCredentialPasswordFile'"
        Write-Host '  To change username and/or password, delete one or all of the files mentioned above and run the script again'
        $ExchangeCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ([PSCredential]::new('X', (Get-Content $ExchangeCredentialUsernameFile -Encoding $UTF8Encoding | ConvertTo-SecureString)).GetNetworkCredential().Password), (Get-Content $ExchangeCredentialPasswordFile -Encoding $UTF8Encoding | ConvertTo-SecureString)
    } else {
        Write-Host '  Use current credential'
        $ExchangeCredential = $null
    }


    # Connect to Exchange
    Write-Host
    Write-Host "Connect to Exchange for single-thread operations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host '  Single-thread Exchange operation'

    . ([scriptblock]::Create($ConnectExchange))


    # Import recipients
    Write-Host
    Write-Host "Import recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host '  Single-thread Exchange operation'

    $AllRecipients = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

    $x = @()

    try {
        $x += @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties) -ErrorAction Stop))
    } catch {
        . ([scriptblock]::Create($ConnectExchange))
        $x += @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties) -ErrorAction Stop))
    }

    if ($ExportPublicFolderPermissions) {
        try {
            $x += @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -PublicFolder -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties) -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $x += @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -PublicFolder -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties) -ErrorAction Stop))
        }
    }

    $AllRecipients.AddRange(@($x | Sort-Object @{Expression = { $_.PrimarySmtpAddress.Address } }))
    $AllRecipients.TrimToSize()
    $x = $null

    Write-Host ('  {0:0000000} recipients found' -f $($AllRecipients.count))


    # Import recipient permissions (SendAs)
    Write-Host
    Write-Host "Import Send As permissions from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (($ExportFromOnPrem -eq $false) -and ($ExportSendAs -eq $true)) {
        Write-Host '  Single-thread Exchange operation'
        $AllRecipientsSendas = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count * 2))

        try {
            $AllRecipientsSendas.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-RecipientPermission -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, trustee, trusteesidstring, accessrights, accesscontroltype, isinherited, inheritancetype } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendas.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-RecipientPermission -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, trustee, trusteesidstring, accessrights, accesscontroltype, isinherited, inheritancetype } -ErrorAction Stop))
        }

        $AllRecipientsSendas.TrimToSize()
        Write-Host ('  {0:0000000} Send As permissions found' -f $($AllRecipientsSendas.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import Send On Behalf from cloud
    Write-Host
    Write-Host "Import Send On Behalf permissions from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (($ExportFromOnPrem -eq $false) -and ($ExportSendOnBehalf -eq $true)) {
        Write-Host '  Single-thread Exchange operation'
        $AllRecipientsSendonbehalf = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count * 2))

        Write-Host "  Mailboxes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailbox -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailbox -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Distribution groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-distributiongroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-distributiongroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Dynamic Distribution Groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-dynamicdistributiongroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-dynamicdistributiongroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Unified Groups (Microsoft 365 Groups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-unifiedgroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-unifiedgroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Mail-enabled Public Folders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicfolder -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicfolder -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        $AllRecipientsSendonbehalf.TrimToSize()
        Write-Host ('  {0:0000000} Send On Behalf permissions found' -f $($AllRecipientsSendonbehalf.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import mailbox databases
    Write-Host
    Write-Host "Import mailbox databases @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportFromOnPrem) {
        Write-Host '  Single-thread Exchange operation'

        $AllMailboxDatabases = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

        try {
            $AllMailboxDatabases.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxDatabase -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property Guid, ProhibitSendQuota } -ErrorAction Stop) | Sort-Object { $_.DisplayName }))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllMailboxDatabases.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxDatabase -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property Guid, ProhibitSendQuota } -ErrorAction Stop) | Sort-Object { $_.DisplayName }))
        }

        $AllMailboxDatabases.TrimToSize()
        Write-Host ('  {0:0000000} mailbox databases found' -f $($AllMailboxDatabases.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import Public Folders
    Write-Host
    Write-Host "Import Public Folders and their content mailboxes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportPublicFolderPermissions) {
        Write-Host '  Single-thread Exchange operation'

        $AllPublicFolders = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

        try {
            $AllPublicFolders.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-PublicFolder -recurse -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property EntryId, ContentMailboxGuid, MailEnabled, MailRecipientGuid, FolderClass, FolderPath } -ErrorAction Stop) | Sort-Object { $_.FolderPath }))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllPublicFolders.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-PublicFolder -recurse -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property EntryId, ContentMailboxGuid, MailEnabled, MailRecipientGuid, FolderClass, FolderPath } -ErrorAction Stop) | Sort-Object { $_.FolderPath }))
        }

        $AllPublicFolders.TrimToSize()
        Write-Host ('  {0:0000000} Public Folders found' -f $($AllPublicFolders.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import Management Role Groups and their members
    Write-Host
    Write-Host "Import Management Role Groups and their members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportManagementRoleGroupMembers) {
        Write-Host '  Single-thread Exchange operation'

        $AllManagementRoleGroupMembers = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

        try {
            $AllRoleGroups = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-rolegroup -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Name, Guid } -ErrorAction Stop) | Sort-Object -Property name
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRoleGroups = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-rolegroup -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Name, Guid } -ErrorAction Stop) | Sort-Object -Property name
        }

        foreach ($RoleGroup in $AllRoleGroups) {
            $RoleGroupMembers = @()

            try {
                $RoleGroupMembers += @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-rolegroupmember $args[0] -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object Identity, PrimarySmtpAddress } -ArgumentList $RoleGroup.guid.guid -ErrorAction Stop)
            } catch {
                . ([scriptblock]::Create($ConnectExchange))
                $RoleGroupMembers += @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-rolegroupmember $args[0] -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object Identity, PrimarySmtpAddress } -ArgumentList $RoleGroup.guid.guid -ErrorAction Stop)
            }

            foreach ($RoleGroupMember in $RoleGroupMembers) {
                $AllManagementRoleGroupMembers.AddRange(@($RoleGroupMember | Select-Object @{name = 'RoleGroup'; expression = { $RoleGroup.Name } }, @{name = 'TrusteeOriginalIdentity'; expression = { $_.Identity.name } }, @{name = 'TrusteePrimarySmtpAddress'; expression = { $_.PrimarySmtpAddress.address } } | Sort-Object -Property RoleGroup, TrusteeOriginalIdentity, TrusteePrimarySmtpAddress))
            }
        }

        $AllRoleGroups = $null
        $RoleGroupMembers = $null
        $AllManagementRoleGroupMembers.TrimToSize()
        Write-Host ('  {0:0000000} Role Group memberships found' -f $($AllManagementRoleGroupMembers.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import forwarding addresses
    Write-Host
    Write-Host "Import forwarding addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportForwarders) {
        Write-Host '  Single-thread Exchange operation'

        foreach ($Recipient in $AllRecipients) {
            $Recipient.ExternalEmailAddress = $Recipient.ExternalEmailAddress.SmtpAddress
        }

        $AllForwardingAddresses = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

        try {
            $AllForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -filter '(ForwardingAddress -ne $null) -or (ForwardingSmtpAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -filter '(ForwardingAddress -ne $null) -or (ForwardingSmtpAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        }

        try {
            $AllForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicFolder -filter 'ForwardingAddress -ne $null' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicFolder -filter 'ForwardingAddress -ne $null' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        }

        $AllForwardingAddresses.TrimToSize()
        Write-Host ('  {0:0000000} forwarding addresses found' -f $($AllForwardingAddresses.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Disconnect from Exchange
    Write-Host
    Write-Host "Single-thread operations completed, remove connection to Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
        Disconnect-ExchangeOnline -Confirm:$false
        Remove-Module ExchangeOnlineManagement
    }

    if ($ExchangeSession) {
        Remove-PSSession -Session $ExchangeSession
    }


    # Create lookup hashtables for GUID, DistinguishedName and PrimarySmtpAddress
    Write-Host
    Write-Host "Create lookup hashtables for GUID, DistinguishedName and PrimarySmtpAddress @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  DistinguishedName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsDnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipientsDnToIndex.ContainsKey($(($AllRecipients[$x]).distinguishedname))) {
            # Same DN defined multiple times - set index to $null
            Write-Verbose "    '$(($AllRecipients[$x]).distinguishedname)' is not unique."
            $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $null
        } else {
            $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $x
        }
    }

    Write-Host "  Identity GUID to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsIdentityGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipientsIdentityGuidToIndex.ContainsKey($(($AllRecipients[$x]).identity.objectguid.guid))) {
            # Same GUID defined multiple times - set index to $null
            Write-Verbose "    '$(($AllRecipients[$x]).identity.objectguid.guid)' is not unique."
            $AllRecipientsIdentityGuidToIndex[$(($AllRecipients[$x]).identity.objectguid.guid)] = $null
        } else {
            $AllRecipientsIdentityGuidToIndex[$(($AllRecipients[$x]).identity.objectguid.guid)] = $x
        }
    }

    Write-Host "  Exchange GUID to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsExchangeGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipientsExchangeGuidToIndex.ContainsKey($(($AllRecipients[$x]).ExchangeGuid.Guid))) {
            # Same GUID defined multiple times - set index to $null
            Write-Verbose "    '$(($AllRecipients[$x]).ExchangeGuid.Guid)' is not unique."
            $AllRecipientsExchangeGuidToIndex[$(($AllRecipients[$x]).ExchangeGuid.Guid)] = $null
        } else {
            $AllRecipientsExchangeGuidToIndex[$(($AllRecipients[$x]).ExchangeGuid.Guid)] = $x
        }
    }

    Write-Host "  EmailAddresses to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsSmtpToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.EmailAddresses.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].EmailAddresses) {
            foreach ($EmailAddress in $AllRecipients[$x].EmailAddresses.SmtpAddress) {
                if ($AllRecipientsSmtpToIndex.ContainsKey($EmailAddress)) {
                    # Same EmailAddress defined multiple times - set index to $null
                    Write-Verbose "    '$($EmailAddress)' is not unique."
                    $AllRecipientsSmtpToIndex[$EmailAddress] = $null
                } else {
                    $AllRecipientsSmtpToIndex[$EmailAddress] = $x
                }
            }
        }
    }


    # Import LinkedMasterAccount
    Write-Host
    Write-Host "Import LinkedMasterAccount of each mailbox by database @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($ExportFromOnPrem) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllMailboxDatabases.count))
        for ($x = 0; $x -lt $AllMailboxDatabases.count; $x++) {
            $tempQueue.enqueue($AllMailboxDatabases[$x].guid.guid)
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $AllRecipientsIdentityGuidToIndex,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $ErrorFile,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchange,
                            $ExchangeOnlineConnectionParameters,
                            $ExchangeCredential,
                            $UseDefaultCredential,
                            $ScriptPath,
                            $VerbosePreference,
                            $DebugPreference,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Import LinkedMasterAccount @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $MailboxDatabaseGuid = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "MailboxDatabaseGuid $($MailboxDatabaseGuid) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    try {
                                        $mailboxes = @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -database $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property Identity, LinkedMasterAccount } -ArgumentList $MailboxDatabaseGuid -ErrorAction Stop))
                                    } catch {
                                        . ([scriptblock]::Create($ConnectExchange))
                                        $mailboxes = @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -database $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property Identity, LinkedMasterAccount } -ArgumentList $MailboxDatabaseGuid -ErrorAction Stop))
                                    }


                                    foreach ($mailbox in $mailboxes) {
                                        if ($mailbox.LinkedMasterAccount) {
                                            try {
                                                ($AllRecipients[$($AllRecipientsIdentityGuidToIndex[$($mailbox.identity.objectguid.guid)])]).LinkedMasterAccount = $mailbox.LinkedMasterAccount
                                            } catch {
                                                """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Import LinkedMasterAccount"";""Mailbox GUID $($mailbox.identity.objectguid.guid)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Import LinkedMasterAccount"";""Database GUID $($MailboxDatabaseGuid)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Import LinkedMasterAccount"";"""";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                                Disconnect-ExchangeOnline -Confirm:$false
                                Remove-Module ExchangeOnlineManagement
                            }

                            if ($ExchangeSession) {
                                Remove-PSSession -Session $ExchangeSession
                            }

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityGuidToIndex   = $AllRecipientsIdentityGuidToIndex
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ConnectExchange                    = $ConnectExchange
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExchangeCredential                 = $ExchangeCredential
                        UseDefaultCredential               = $UseDefaultCredential
                        ScriptPath                         = $PSScriptRoot
                        VerbosePreference                  = $VerbosePreference
                        DebugPreference                    = $DebugPreference
                        UTF8Encoding                       = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} databases to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all databases have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # UserFriendlyNames
    Write-Host
    Write-Host "Find each recipient's UserFriendlyNames @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxAccessRights -or $ExportSendAs) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        for ($RecipientID = 0; $RecipientID -lt $AllRecipients.count; $RecipientID++) {
            $tempqueue.enqueue($RecipientID)
        }

        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min([math]::ceiling($tempQueueCount / 100), $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param (
                            $AllRecipients,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $AllRecipientsIdentityGuidToIndex,
                            $DebugFile,
                            $ErrorFile,
                            $ExportFromOnPrem,
                            $ConnectExchange,
                            $ExchangeOnlineConnectionParameters,
                            $ExchangeCredential,
                            $UseDefaultCredential,
                            $ScriptPath,
                            $VerbosePreference,
                            $DebugPreference
                        )

                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Create connection between UserFriendlyNames and recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            while ($tempQueue.count -gt 0) {
                                Write-Host "Filter string @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                $dequeued = 0
                                $filterstring = ''

                                while (($dequeued -lt 100) -and ($tempQueue.count -gt 0)) {
                                    try {
                                        $x = $tempQueue.dequeue()
                                    } catch {
                                    }
                                    if ($x) {
                                        $filterstring += "(guid -eq '$($AllRecipients[$x].identity.objectguid.guid)') -or "
                                        $dequeued++
                                    }
                                }
                                $filterstring = $filterstring.trimend(' -or ')

                                Write-Host "  $filterstring"

                                if ($filterstring -ne '') {
                                    try {
                                        $securityprincipals = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-securityprincipal -filter "$($args[0])" -resultsize unlimited -WarningAction silentlycontinue | Select-Object userfriendlyname, guid } -ArgumentList $filterstring -ErrorAction Stop)
                                    } catch {
                                        . ([scriptblock]::Create($ConnectExchange))
                                        $securityprincipals = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-securityprincipal -filter "$($args[0])" -resultsize unlimited -WarningAction silentlycontinue | Select-Object userfriendlyname, guid } -ArgumentList $filterstring -ErrorAction Stop)
                                    }

                                    foreach ($securityprincipal in $securityprincipals) {
                                        try {
                                            Write-Host "  '$($securityprincipal.guid.guid)' = '$($securityprincipal.UserFriendlyName)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                            ($AllRecipients[$($AllRecipientsIdentityGuidToIndex[$($securityprincipal.guid.guid)])]).UserFriendlyName = $securityprincipal.UserFriendlyName
                                        } catch {
                                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Find each recipient's UserFriendlyNames"";""$($securityprincipcal.UserFriendlyName)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                        }
                                    }
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Find each recipient's UserFriendlyNames"";"""";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                                Disconnect-ExchangeOnline -Confirm:$false
                                Remove-Module ExchangeOnlineManagement
                            }

                            if ($ExchangeSession) {
                                Remove-PSSession -Session $ExchangeSession
                            }

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        AllRecipientsIdentityGuidToIndex   = $AllRecipientsIdentityGuidToIndex
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ExchangeCredential                 = $ExchangeCredential
                        UseDefaultCredential               = $UseDefaultCredential
                        ScriptPath                         = $PSScriptRoot
                        ConnectExchange                    = $ConnectExchange
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        VerbosePreference                  = $VerbosePreference
                        DebugPreference                    = $DebugPreference
                        UTF8Encoding                       = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - (($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count * 100))
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all recipients have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Create lookup hashtables for UserFriendlyName and LinkedMasterAccount
    Write-Host
    Write-Host "Create lookup hashtables for UserFriendlyName and LinkedMasterAccount @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  UserFriendlyName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsUfnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $Recipient = $AllRecipients[$x]
        if ($Recipient.userfriendlyname) {
            if ($AllRecipientsUfnToIndex.ContainsKey($($Recipient.userfriendlyname))) {
                # Same UserFriendlName defined multiple time - set index to $null
                if ($AllRecipientsUfnToIndex[$($Recipient.userfriendlyname)]) {
                    Write-Verbose "    '$($Recipient.userfriendlyname)' used not only once: '$($AllRecipients[$($AllRecipientsUfnToIndex[$($Recipient.userfriendlyname)])].primarysmtpaddress.address)'"
                }

                Write-Verbose "    '$($Recipient.userfriendlyname)' used not only once: '$($Recipient.primarysmtpaddress.address)'"

                $AllRecipientsUfnToIndex[$Recipient.userfriendlyname] = $null
            } else {
                $AllRecipientsUfnToIndex[$Recipient.userfriendlyname] = $x
            }
        }
    }

    Write-Host "  LinkedMasterAccount to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsLinkedmasteraccountToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if (($AllRecipients[$x]).LinkedMasterAccount) {
            if ($AllRecipientsLinkedmasteraccountToIndex.ContainsKey($(($AllRecipients[$x]).LinkedMasterAccount))) {
                # Same LinkedMasterAccount defined multiple time - set index to $null
                if ($AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)]) {
                    Write-Verbose "    '$(($AllRecipients[$x]).LinkedMasterAccount)' used not only once: '$($AllRecipients[$($AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)])].primarysmtpaddress.address)'"
                }

                Write-Verbose "    '$(($AllRecipients[$x]).LinkedMasterAccount)' used not only once: '$(($AllRecipients[$x]).primarysmtpaddress.address)'"

                $AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)] = $null
            } else {
                $AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)] = $x
            }
        }
    }


    # Grantors
    Write-Host
    Write-Host "Define grantors by filtering recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "  Filter: { $($GrantorFilter) }"
    $GrantorsToConsider = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $Grantor = $AllRecipients[$x]

        if (-not $GrantorFilter) {
            $null = $GrantorsToConsider.add($x)
        } else {
            if ((. ([scriptblock]::Create($GrantorFilter))) -eq $true) {
                $null = $GrantorsToConsider.add($x)
            }
        }
    }
    $GrantorsToConsider.TrimToSize()
    Write-Host ('  {0:0000000}/{1:0000000} recipients are considered as grantors' -f $($GrantorsToConsider.count), $($AllRecipients.count))


    # Mailbox Access Rights
    Write-Host
    Write-Host "Mailbox Access Rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxAccessRights) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.recipienttypedetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $ExportMailboxAccessRightsSelf,
                            $ExportMailboxAccessRightsInherited,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $AllRecipientsUfnToIndex,
                            $AllRecipientsLinkedMasterAccountToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchange,
                            $ExchangeOnlineConnectionParameters,
                            $ExchangeCredential,
                            $UseDefaultCredential,
                            $ScriptPath,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Mailbox access rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileResult.clear()
                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }
                                $Grantor = $AllRecipients[$RecipientID]
                                $Trustee = $null

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    foreach ($MailboxPermission in
                                        @($(
                                                if ($ExportFromOnPrem) {
                                                    try {
                                                        Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                    } catch {
                                                        . ([scriptblock]::Create($ConnectExchange))
                                                        Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                    }
                                                } else {
                                                    if ($GrantorRecipientTypeDetails -ine 'groupmailbox') {
                                                        try {
                                                            Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                        } catch {
                                                            . ([scriptblock]::Create($ConnectExchange))
                                                            Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                        }
                                                    }
                                                }
                                            ))
                                    ) {
                                        foreach ($TrusteeRight in @($MailboxPermission | Where-Object { if ($ExportMailboxAccessRightsSelf) { $true } else { $_.user.SecurityIdentifier -ine 'S-1-5-10' } } | Where-Object { if ($ExportMailboxAccessRightsInherited) { $true } else { $_.IsInherited -ne $true } } | Select-Object *, @{ name = 'trustee'; Expression = { $_.user.rawidentity } })) {
                                            $trustees = [system.collections.arraylist]::new(1000)


                                            try {
                                                $index = $null
                                                $index = ($AllRecipientsUfnToIndex[$($TrusteeRight.trustee)], $AllRecipientsLinkedmasteraccountToIndex[$($TrusteeRight.trustee)]) | Select-Object -First 1
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $trustees.add($AllRecipients[$index])
                                            } else {
                                                $trustees.add($TrusteeRight.trustee)
                                            }
                                            foreach ($Trustee in $Trustees) {
                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }
                                                foreach ($Accessright in ($TrusteeRight.Accessrights -split ', ')) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                        $ExportFileresult.add((('"' + ((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            $Accessright,
                                                                            $(if ($Trusteeright.deny) { 'Deny' } else { 'Allow' }),
                                                                            $Trusteeright.IsInherited,
                                                                            $Trusteeright.InheritanceType,
                                                                            $TrusteeRight.trustee,
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Mailbox Access Rights"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Mailbox Access Rights"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                                Disconnect-ExchangeOnline -Confirm:$false
                                Remove-Module ExchangeOnlineManagement
                            }

                            if ($ExchangeSession) {
                                Remove-PSSession -Session $ExchangeSession
                            }

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                           = $AllRecipients
                        tempConnectionUriQueue                  = $tempConnectionUriQueue
                        tempQueue                               = $tempQueue
                        ExportMailboxAccessRightsSelf           = $ExportMailboxAccessRightsSelf
                        ExportMailboxAccessRightsInherited      = $ExportMailboxAccessRightsInherited
                        ExportFile                              = $ExportFile
                        ExportTrustees                          = $ExportTrustees
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        AllRecipientsUfnToIndex                 = $AllRecipientsUfnToIndex
                        AllRecipientsLinkedMasterAccountToIndex = $AllRecipientsLinkedMasterAccountToIndex
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExchangeCredential                      = $ExchangeCredential
                        UseDefaultCredential                    = $UseDefaultCredential
                        ScriptPath                              = $PSScriptRoot
                        ConnectExchange                         = $ConnectExchange
                        ExchangeOnlineConnectionParameters      = $ExchangeOnlineConnectionParameters
                        VerbosePreference                       = $VerbosePreference
                        DebugPreference                         = $DebugPreference
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantor mailboxes to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all grantor mailboxes have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Mailbox Folder Permissions
    Write-Host
    Write-Host "Mailbox Folder Permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxFolderPermissions) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.recipienttypedetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $ExportMailboxFolderPermissions,
                            $ExportMailboxFolderPermissionsAnonymous,
                            $ExportMailboxFolderPermissionsDefault,
                            $ExportMailboxFolderPermissionsOwnerAtLocal,
                            $ExportMailboxFolderPermissionsMemberAtLocal,
                            $ExportMailboxFolderPermissionsExcludeFoldertype,
                            $ExportFile,
                            $ErrorFile,
                            $ExportTrustees,
                            $AllRecipientsSmtpToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchange,
                            $ExchangeOnlineConnectionParameters,
                            $ExchangeCredential,
                            $UseDefaultCredential,
                            $ScriptPath,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Mailbox folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileResult.Clear()
                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    $Folders = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderstatistics -identity $args[0] -ErrorAction Stop -WarningAction silentlycontinue | Select-Object folderid, folderpath, foldertype } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop
                                } catch {
                                    . ([scriptblock]::Create($ConnectExchange))
                                    $Folders = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderstatistics -identity $args[0] -ErrorAction Stop -WarningAction silentlycontinue | Select-Object folderid, folderpath, foldertype } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop
                                }

                                foreach ($Folder in $Folders) {
                                    try {
                                        if (-not $folder.foldertype) { $folder.foldertype = $null }

                                        if ($folder.foldertype -iin $ExportMailboxFolderPermissionsExcludeFoldertype) { continue }

                                        if ($Folder.foldertype -ieq 'root') { $Folder.folderpath = '/' }

                                        Write-Host "  Folder '$($folder.folderid)' ('$folder.folderpath)') @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                        foreach ($FolderPermissions in
                                            @($(
                                                    if ($ExportFromOnPrem) {
                                                        try {
                                                            (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                        } catch {
                                                            . ([scriptblock]::Create($ConnectExchange))
                                                            (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                        }
                                                    } else {
                                                        if ($GrantorRecipientTypeDetails -ieq 'groupmailbox') {
                                                            try {
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -groupmailbox -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            } catch {
                                                                . ([scriptblock]::Create($ConnectExchange))
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -groupmailbox -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            }
                                                        } else {
                                                            try {
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            } catch {
                                                                . ([scriptblock]::Create($ConnectExchange))
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            }
                                                        }
                                                    }
                                                ))
                                        ) {
                                            foreach ($FolderPermission in $FolderPermissions) {
                                                foreach ($AccessRight in ($FolderPermission.AccessRights)) {
                                                    if ($ExportMailboxFolderPermissionsDefault -eq $false) {
                                                        if ($FolderPermission.user.usertype.value -ieq 'default') { continue }
                                                    }

                                                    if ($ExportMailboxFolderPermissionsAnonymous -eq $false) {
                                                        if ($FolderPermission.user.usertype.value -ieq 'anonymous') { continue }
                                                    }

                                                    if ($ExportFromOnPrem) {
                                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $FolderPermission.user.adrecipient.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($FolderPermission.user.adrecipient.PrimarySmtpAddress))) {
                                                            $trustee = $null

                                                            try {
                                                                $index = $null
                                                                $index = $AllRecipientsSmtpToIndex[$($FolderPermission.user.adrecipient.primarysmtpaddress)]
                                                            } catch {
                                                            }

                                                            if ($index -ge 0) {
                                                                $trustee = $AllRecipients[$index]
                                                            } else {
                                                                $trustee = $($FolderPermission.user.displayname)
                                                            }

                                                            if ($TrusteeFilter) {
                                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                                    continue
                                                                }
                                                            }

                                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }

                                                            $ExportFileResult.Add((('"' + ((
                                                                                $GrantorPrimarySMTP,
                                                                                $GrantorDisplayName,
                                                                                ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                                $GrantorEnvironment,
                                                                                $($Folder.Folderpath),
                                                                                $($Accessright),
                                                                                'Allow',
                                                                                'False',
                                                                                'None',
                                                                                $($FolderPermission.user.displayname),
                                                                                $($Trustee.primarysmtpaddress.address),
                                                                                $($Trustee.displayname),
                                                                                ("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                                $TrusteeEnvironment
                                                                            ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                        }
                                                    } else {
                                                        if ($ExportMailboxFolderPermissionsOwnerAtLocal -eq $false) {
                                                            if ($FolderPermission.user.recipientprincipal.primarysmtpaddress -ieq 'owner@local') { continue }
                                                        }

                                                        if ($ExportMailboxFolderPermissionsMemberAtLocal -eq $false) {
                                                            if ($FolderPermission.user.recipientprincipal.primarysmtpaddress -ieq 'member@local') { continue }
                                                        }

                                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $FolderPermission.user.recipientprincipal.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($FolderPermission.user.recipientprincipal.PrimarySmtpAddress))) {
                                                            $trustee = $null

                                                            try {
                                                                $index = $null
                                                                $index = $AllRecipientsSmtpToIndex[$($FolderPermission.user.recipientprincipal.primarysmtpaddress)]
                                                            } catch {
                                                            }

                                                            if ($index -ge 0) {
                                                                $trustee = $AllRecipients[$index]
                                                            } else {
                                                                $trustee = $($FolderPermission.user.displayname)
                                                            }

                                                            if ($TrusteeFilter) {
                                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                                    continue
                                                                }
                                                            }

                                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }

                                                            $ExportFileResult.Add((('"' + ((
                                                                                $GrantorPrimarySMTP,
                                                                                $GrantorDisplayName,
                                                                                ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                                $GrantorEnvironment,
                                                                                $($Folder.Folderpath),
                                                                                $($Accessright),
                                                                                'Allow',
                                                                                'False',
                                                                                'None',
                                                                                $($FolderPermission.user.displayname),
                                                                                $($Trustee.primarysmtpaddress.address),
                                                                                $($Trustee.displayname),
                                                                                ("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                                $TrusteeEnvironment
                                                                            ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    } catch {
                                        """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Mailbox Folder permissions"";""$($GrantorPrimarySMTP):$($Folder.folderid) ($($Folder.folderpath))"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                    }
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Mailbox Folder permissions"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                                Disconnect-ExchangeOnline -Confirm:$false
                                Remove-Module ExchangeOnlineManagement
                            }

                            if ($ExchangeSession) {
                                Remove-PSSession -Session $ExchangeSession
                            }

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                                   = $AllRecipients
                        tempConnectionUriQueue                          = $tempConnectionUriQueue
                        tempQueue                                       = $tempQueue
                        ExportMailboxFolderPermissions                  = $ExportMailboxFolderPermissions
                        ExportMailboxFolderPermissionsAnonymous         = $ExportMailboxFolderPermissionsAnonymous
                        ExportMailboxFolderPermissionsDefault           = $ExportMailboxFolderPermissionsDefault
                        ExportMailboxFolderPermissionsOwnerAtLocal      = $ExportMailboxFolderPermissionsOwnerAtLocal
                        ExportMailboxFolderPermissionsMemberAtLocal     = $ExportMailboxFolderPermissionsMemberAtLocal
                        ExportMailboxFolderPermissionsExcludeFoldertype = $ExportMailboxFolderPermissionsExcludeFoldertype
                        ExportFile                                      = $ExportFile
                        ExportTrustees                                  = $ExportTrustees
                        AllRecipientsSmtpToIndex                        = $AllRecipientsSmtpToIndex
                        ErrorFile                                       = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                                       = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                                = $ExportFromOnPrem
                        ExchangeCredential                              = $ExchangeCredential
                        UseDefaultCredential                            = $UseDefaultCredential
                        ScriptPath                                      = $PSScriptRoot
                        ConnectExchange                                 = $ConnectExchange
                        ExchangeOnlineConnectionParameters              = $ExchangeOnlineConnectionParameters
                        VerbosePreference                               = $VerbosePreference
                        DebugPreference                                 = $DebugPreference
                        TrusteeFilter                                   = $TrusteeFilter
                        UTF8Encoding                                    = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantor mailboxes to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all grantor mailboxes have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Send As
    Write-Host
    Write-Host "Send As @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportSendAs) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            if ($x -in $GrantorsToConsider) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        if ($ExportFromOnPrem) {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsAD)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel AD jobs"
        } else {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"
        }
        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $AllRecipientsUfnToIndex,
                            $AllRecipientsLinkedMasterAccountToIndex,
                            $AllRecipientsSmtpToIndex,
                            $ExportSendAsSelf,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $AllRecipientsSendas,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Send As @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $result = [system.collections.arraylist]::new(1000)
                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $result.clear()
                                $ExportFileresult.clear()
                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    if ($ExportFromOnPrem) {
                                        foreach ($entry in (([adsi]"LDAP://<GUID=$($Grantor.identity.objectguid.guid)>").ObjectSecurity.Access)) {
                                            $trustee = $null

                                            if ($entry.ObjectType -eq 'ab721a54-1e2f-11d0-9819-00aa0040529b') {
                                                if (($entry.identityreference -ilike '*\*') -and ($ExportSendAsSelf -eq $false)) {
                                                    if ((([System.Security.Principal.NTAccount]::new($entry.identityreference)).Translate([System.Security.Principal.SecurityIdentifier])).value -ieq 'S-1-5-10') {
                                                        continue
                                                    }
                                                }

                                                try {
                                                    $index = $null
                                                    $index = ($AllRecipientsUfnToIndex[$($entry.identityreference.tostring())], $AllRecipientsLinkedmasteraccountToIndex[$($entry.identityreference.tostring())]) | Select-Object -First 1
                                                } catch {
                                                }

                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    $trustee = $entry.identityreference
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                    $ExportFileresult.add((('"' + ((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendAs',
                                                                        $entry.AccessControlType,
                                                                        $entry.IsInherited,
                                                                        $entry.InheritanceType,
                                                                        $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                }
                                            }
                                        }
                                    } else {
                                        foreach ($entry in $AllRecipientsSendas) {
                                            if ($entry.identity.objectguid.guid -eq $Grantor.identity.objectguid.guid) {
                                                if (($ExportSendAsSelf -eq $false) -and ($entry.trusteesidstring -ieq 'S-1-5-10')) {
                                                    continue
                                                }
                                                $trustee = $null

                                                if ($entry.trustee -ilike '*\*') {
                                                    try {
                                                        $index = $null
                                                        $index = ($AllRecipientsUfnToIndex[$($entry.trustee)], $AllRecipientsLinkedmasteraccountToIndex[$($entry.trustee)]) | Select-Object -First 1
                                                    } catch {
                                                    }
                                                } elseif ($entry.trustee -ilike '*@*') {
                                                    $index = $null
                                                    $index = $AllRecipientsSmtpToIndex[$($entry.trustee)]
                                                }

                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    $trustee = $entry.trustee
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                foreach ($AccessRight in $entry.AccessRights) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {

                                                        $ExportFileresult.add((('"' + ((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            $AccessRight,
                                                                            $entry.AccessControlType,
                                                                            $entry.IsInherited,
                                                                            $entry.InheritanceType,
                                                                            $(($Trustee.displayname, $entry.trustee) | Select-Object -First 1),
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""SendAs"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Send As"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                           = $AllRecipients
                        tempQueue                               = $tempQueue
                        ExportFile                              = $ExportFile
                        ExportTrustees                          = $ExportTrustees
                        AllRecipientsUfnToIndex                 = $AllRecipientsUfnToIndex
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
                        AllRecipientsLinkedMasterAccountToIndex = $AllRecipientsLinkedMasterAccountToIndex
                        ExportSendAsSelf                        = $ExportSendAsSelf
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ScriptPath                              = $PSScriptRoot
                        AllRecipientsSendas                     = $AllRecipientsSendas
                        VerbosePreference                       = $VerbosePreference
                        DebugPreference                         = $DebugPreference
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all grantors have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Send On Behalf
    Write-Host
    Write-Host "Send On Behalf @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportSendOnBehalf) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            if (($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        if ($ExportFromOnPrem) {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsAD)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel AD jobs"
        } else {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"
        }

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $AllRecipientsDnToIndex,
                            $AllRecipientsIdentityGuidToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $AllRecipientsSendonbehalf,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Send On Behalf @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)
                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()
                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    if ($ExportFromOnPrem) {
                                        $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher("(objectguid=$([System.String]::Join('', (([guid]$($Grantor.identity.objectguid.guid)).ToByteArray() | ForEach-Object { '\' + $_.ToString('x2') })).ToUpper()))")
                                        $directorySearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($Grantor.identity.domainid)")
                                        $directorySearcher.PropertiesToLoad.Add('publicDelegates')
                                        $directorySearcherResults = $directorySearcher.FindOne()

                                        foreach ($directorySearcherResult in $directorySearcherResults) {
                                            foreach ($delegateBindDN in $directorySearcherResult.properties.publicdelegates) {
                                                $index = $null
                                                $index = $AllRecipientsDnToIndex[$delegateBindDN]

                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    $trustee = $delegateBindDN
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                    $ExportFileresult.add((('"' + ((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendOnBehalf',
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                }
                                            }
                                        }
                                    } else {
                                        foreach ($entry in $AllRecipientsSendonbehalf) {
                                            if ($entry.identity.objectguid.guid -eq $Grantor.identity.objectguid.guid) {
                                                $trustee = $null
                                                foreach ($AccessRight in $entry.GrantSendOnBehalfTo) {
                                                    $index = $null
                                                    $index = $AllRecipientsIdentityGuidToIndex[$($AccessRight.objectguid.guid)]

                                                    if ($index -ge 0) {
                                                        $trustee = $AllRecipients[$index]
                                                    } else {
                                                        $trustee = $AccessRight.tostring()
                                                    }

                                                    if ($TrusteeFilter) {
                                                        if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                            continue
                                                        }
                                                    }

                                                    if ($ExportFromOnPrem) {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                    } else {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                    }

                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                        $ExportFileresult.add((('"' + ((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            'SendOnBehalf',
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Send On Behalf"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Send On Behalf"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                    = $AllRecipients
                        tempQueue                        = $tempQueue
                        ExportFile                       = $ExportFile
                        ExportTrustees                   = $ExportTrustees
                        AllRecipientsDnToIndex           = $AllRecipientsDnToIndex
                        AllRecipientsIdentityGuidToIndex = $AllRecipientsIdentityGuidToIndex
                        ErrorFile                        = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                        = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                 = $ExportFromOnPrem
                        ScriptPath                       = $PSScriptRoot
                        AllRecipientsSendonbehalf        = $AllRecipientsSendonbehalf
                        VerbosePreference                = $VerbosePreference
                        DebugPreference                  = $DebugPreference
                        TrusteeFilter                    = $TrusteeFilter
                        UTF8Encoding                     = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all grantors have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }


            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Managed By
    Write-Host
    Write-Host "Managed By @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportManagedBy) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $AllRecipientsIdentityGuidToIndex,
                            $DebugFile,
                            $ScriptPath,
                            $ExportFromOnPrem,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Managed By @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()
                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    foreach ($TrusteeRight in $Grantor.ManagedBy) {
                                        $trustees = [system.collections.arraylist]::new(1000)
                                        $index = $null
                                        $index = $AllRecipientsIdentityGuidToIndex[$($TrusteeRight.objectguid.guid)]

                                        if ($index -ge 0) {
                                            $trustees.add($AllRecipients[$index])
                                        } else {
                                            $trustees.add((($TrusteeRight.distinguishedname, $TrusteeRight.identity.objectguid.guid) | Select-Object -First 1))
                                        }

                                        foreach ($Trustee in $Trustees) {
                                            if ($TrusteeFilter) {
                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                    continue
                                                }
                                            }

                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                $ExportFileresult.add((('"' + ((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'ManagedBy',
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Managed By"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Managed By"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                    = $AllRecipients
                        tempQueue                        = $tempQueue
                        ExportFile                       = $ExportFile
                        ExportTrustees                   = $ExportTrustees
                        AllRecipientsIdentityGuidToIndex = $AllRecipientsIdentityGuidToIndex
                        ErrorFile                        = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                        = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                       = $PSScriptRoot
                        ExportFromOnPrem                 = $ExportFromOnPrem
                        VerbosePreference                = $VerbosePreference
                        DebugPreference                  = $DebugPreference
                        TrusteeFilter                    = $TrusteeFilter
                        UTF8Encoding                     = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all grantors have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Linked Master Account
    Write-Host
    Write-Host "Linked Master Account @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportLinkedMasterAccount -and $ExportFromOnPrem) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.recipienttypedetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $ExportLinkedMasterAccount,
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $AllRecipientsLinkedmasteraccountToIndex,
                            $DebugFile,
                            $ScriptPath,
                            $ExportFromOnPrem,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Linked Master Account @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    try {
                                        $index = $null
                                        $index = $AllRecipientsLinkedmasteraccountToIndex[$($Grantor.LinkedMasterAccount)]
                                    } catch {
                                    }

                                    if ($index -ge 0) {
                                        $Trustee = $AllRecipients[$index]
                                    } else {
                                        $Trustee = $Grantor.LinkedMasterAccount
                                    }

                                    if ($Grantor.LinkedMasterAccount) {
                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                $ExportFileresult.add((('"' + ((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'LinkedMasterAccount',
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $Grantor.LinkedMasterAccount,
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Linked Master Account"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Linked Master Account"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        ExportLinkedMasterAccount               = $ExportLinkedMasterAccount
                        AllRecipients                           = $AllRecipients
                        tempQueue                               = $tempQueue
                        ExportFile                              = $ExportFile
                        ExportTrustees                          = $ExportTrustees
                        AllRecipientsLinkedmasteraccountToIndex = $AllRecipientsLinkedmasteraccountToIndex
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                              = $PSScriptRoot
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        VerbosePreference                       = $VerbosePreference
                        DebugPreference                         = $DebugPreference
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all grantors have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Public Folder Permissions
    Write-Host
    Write-Host "Public Folder Permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportPublicFolderPermissions) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllPublicFolders.count))

        for ($x = 0; $x -lt $AllPublicFolders.count; $x++) {
            $tempQueue.enqueue($x)
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllPublicFolders,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $ExportPublicFolderPermissions,
                            $ExportPublicFolderPermissionsAnonymous,
                            $ExportPublicFolderPermissionsDefault,
                            $ExportPublicFolderPermissionsExcludeFoldertype,
                            $ExportFile,
                            $ErrorFile,
                            $ExportTrustees,
                            $GrantorFilter,
                            $AllRecipients,
                            $AllRecipientsSmtpToIndex,
                            $AllRecipientsIdentityGuidToIndex,
                            $AllRecipientsExchangeGuidToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchange,
                            $ExchangeOnlineConnectionParameters,
                            $ExchangeCredential,
                            $UseDefaultCredential,
                            $ScriptPath,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Public folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileResult.Clear()
                                try {
                                    $PublicFolderId = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $folder = $AllPublicFolders[$PublicFolderId]

                                try {
                                    $index = $null
                                    $index = $AllRecipientsExchangeGuidToIndex[$($folder.ContentMailboxGuid.Guid)]
                                } catch {
                                    Write-Host 'GUID not found in AllRecipientsExchangeGuidToIndex'
                                }

                                if ($index -ge 0) {
                                    $RecipientId = $index
                                    $Grantor = $AllRecipients[$RecipientId]
                                } else {
                                    continue
                                }

                                if ($GrantorFilter) {
                                    if ((. ([scriptblock]::Create($GrantorFilter))) -ne $true) {
                                        continue
                                    }
                                }

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    $folder.folderpath = '/' + ($folder.folderpath -join '/')

                                    Write-Host "  Folder '$($folder.EntryId)' ('$($folder.folderpath)') @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                    if (-not $folder.folderclass) { $folder.folderclass = $null }

                                    if ($folder.folderclass -iin $ExportPublicFolderPermissionsExcludeFoldertype) { continue }


                                    if ($folder.MailEnabled) {
                                        $trustee = $null

                                        try {
                                            $index = $null
                                            $index = $AllRecipientsIdentityGuidToIndex[$($folder.MailRecipientGuid.Guid)]
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $trustee = $AllRecipients[$index]
                                        } else {
                                            $trustee = $($folder.MailRecipientGuid.Guid)
                                        }

                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }
                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            $ExportFileResult.Add((('"' + ((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                $($Folder.Folderpath),
                                                                'MailEnabled',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $($Trustee.primarysmtpaddress.address),
                                                                $($Trustee.primarysmtpaddress.address),
                                                                $($Trustee.displayname),
                                                                $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))

                                        }
                                    }

                                    foreach ($FolderPermissions in
                                        @($(
                                                try {
                                                    Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-publicfolderclientpermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList $($Folder.EntryId) -ErrorAction Stop
                                                } catch {
                                                    . ([scriptblock]::Create($ConnectExchange))
                                                    Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-publicfolderclientpermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList $($Folder.EntryId) -ErrorAction Stop
                                                }
                                            ))
                                    ) {
                                        foreach ($FolderPermission in $FolderPermissions) {
                                            foreach ($AccessRight in ($FolderPermission.AccessRights)) {
                                                if ($ExportPublicFolderPermissionsDefault -eq $false) {
                                                    if ($FolderPermission.user.usertype.value -ieq 'default') { continue }
                                                }

                                                if ($ExportPublicFolderPermissionsAnonymous -eq $false) {
                                                    if ($FolderPermission.user.usertype.value -ieq 'anonymous') { continue }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $FolderPermission.user.adrecipient.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($FolderPermission.user.adrecipient.PrimarySmtpAddress))) {
                                                        $trustee = $null

                                                        try {
                                                            $index = $null
                                                            $index = $AllRecipientsSmtpToIndex[$($FolderPermission.user.adrecipient.primarysmtpaddress)]
                                                        } catch {
                                                        }

                                                        if ($index -ge 0) {
                                                            $trustee = $AllRecipients[$index]
                                                        } else {
                                                            $trustee = $($FolderPermission.user.displayname)
                                                        }

                                                        if ($TrusteeFilter) {
                                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                                continue
                                                            }
                                                        }

                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }

                                                        $ExportFileResult.Add((('"' + ((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            $($Folder.Folderpath),
                                                                            $($Accessright),
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $($FolderPermission.user.displayname),
                                                                            $($Trustee.primarysmtpaddress.address),
                                                                            $($Trustee.displayname),
                                                                            $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                    }
                                                } else {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $FolderPermission.user.recipientprincipal.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($FolderPermission.user.recipientprincipal.PrimarySmtpAddress))) {
                                                        $trustee = $null

                                                        try {
                                                            $index = $null
                                                            $index = $AllRecipientsSmtpToIndex[$($FolderPermission.user.recipientprincipal.primarysmtpaddress)]
                                                        } catch {
                                                        }

                                                        if ($index -ge 0) {
                                                            $trustee = $AllRecipients[$index]
                                                        } else {
                                                            $trustee = $($FolderPermission.user.displayname)
                                                        }

                                                        if ($TrusteeFilter) {
                                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                                continue
                                                            }
                                                        }

                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }

                                                        $ExportFileResult.Add((('"' + ((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            $($Folder.Folderpath),
                                                                            $($Accessright),
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $($FolderPermission.user.displayname),
                                                                            $($Trustee.primarysmtpaddress.address),
                                                                            $($Trustee.displayname),
                                                                            $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Public Folder permissions"";""$($GrantorPrimarySMTP):$($Folder.folderid) ($($Folder.folderpath))"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.PF{1:0000000}.txt' -f $RecipientId, $PublicFolderId))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Public Folder permissions"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
                                Disconnect-ExchangeOnline -Confirm:$false
                                Remove-Module ExchangeOnlineManagement
                            }

                            if ($ExchangeSession) {
                                Remove-PSSession -Session $ExchangeSession
                            }

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllPublicFolders                               = $AllPublicFolders
                        tempConnectionUriQueue                         = $tempConnectionUriQueue
                        tempQueue                                      = $tempQueue
                        ExportPublicFolderPermissions                  = $ExportPublicFolderPermissions
                        ExportPublicFolderPermissionsAnonymous         = $ExportPublicFolderPermissionsAnonymous
                        ExportPublicFolderPermissionsDefault           = $ExportPublicFolderPermissionsDefault
                        ExportPublicFolderPermissionsExcludeFoldertype = $ExportPublicFolderPermissionsExcludeFoldertype
                        ExportFile                                     = $ExportFile
                        ExportTrustees                                 = $ExportTrustees
                        GrantorFilter                                  = $GrantorFilter
                        AllRecipients                                  = $AllRecipients
                        AllRecipientsSmtpToIndex                       = $AllRecipientsSmtpToIndex
                        AllRecipientsIdentityGuidToIndex               = $AllRecipientsIdentityGuidToIndex
                        AllRecipientsExchangeGuidToIndex               = $AllRecipientsExchangeGuidToIndex
                        ErrorFile                                      = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                                      = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                               = $ExportFromOnPrem
                        ExchangeCredential                             = $ExchangeCredential
                        UseDefaultCredential                           = $UseDefaultCredential
                        ScriptPath                                     = $PSScriptRoot
                        ConnectExchange                                = $ConnectExchange
                        ExchangeOnlineConnectionParameters             = $ExchangeOnlineConnectionParameters
                        VerbosePreference                              = $VerbosePreference
                        DebugPreference                                = $DebugPreference
                        TrusteeFilter                                  = $TrusteeFilter
                        UTF8Encoding                                   = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} Public Folders to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all Public Folders have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            if ($ResultFile) {
                foreach ($JobResultFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ResultFile), ('TEMP.*.PF*.txt'))))) {
                    Get-Content $JobResultFile -Encoding $UTF8Encoding | Add-Content ($JobResultFile.fullname -replace '\.PF\d{7}.txt$', '.txt') -Encoding $UTF8Encoding -Force
                    Remove-Item $JobResultFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Management Role Group Members
    Write-Host
    Write-Host "Management Role Group Members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportManagementRoleGroupMembers) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllManagementRoleGroupMembers.count))

        foreach ($x in (0..($AllManagementRoleGroupMembers.count - 1))) {
            $tempQueue.enqueue($x)
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllManagementRoleGroupMembers,
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $AllRecipientsSmtpToIndex,
                            $DebugFile,
                            $ScriptPath,
                            $ExportFromOnPrem,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Linked Master Account @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()

                                try {
                                    $RoleGroupMemberId = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $RoleGroupMember = $AllManagementRoleGroupMembers[$RoleGroupMemberId]

                                $GrantorPrimarySMTP = 'Management Role Group'
                                $GrantorDisplayName = $null
                                $GrantorRecipientType = 'ManagementRoleGroup'
                                $GrantorRecipientTypeDetails = $null
                                if ($ExportFromOnPrem) {
                                    $GrantorEnvironment = 'On-Prem'
                                } else {
                                    $GrantorEnvironment = 'Cloud'
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($RoleGroupMember.RoleGroup), $($RoleGroupMember.TrusteeOriginalIdentity) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    try {
                                        $index = $null
                                        $index = $AllRecipientsSmtpToIndex[$($RoleGroupMember.TrusteePrimarySmtpAddress)]
                                    } catch {
                                    }

                                    if ($index -ge 0) {
                                        $Trustee = $AllRecipients[$index]
                                    } else {
                                        $Trustee = $RoleGroupMember.TrusteeOriginalIdentity
                                    }

                                    if ($RoleGroupMember.TrusteePrimarySmtpAddress) {
                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }
                                    }

                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                        if ($ExportFromOnPrem) {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                        } else {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                            $ExportFileresult.add((('"' + ((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $GrantorRecipientType, $GrantorEnvironment,
                                                                $RoleGroupMember.RoleGroup,
                                                                'Member',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $RoleGroupMember.TrusteeOriginalIdentity,
                                                                $Trustee.PrimarySmtpAddress.address,
                                                                $Trustee.DisplayName,
                                                                $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                        }
                                    }
                                } catch {
                                    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Management Role Group Members"";""$($($GrantorPrimarySMTP), $($RoleGroupMember.RoleGroup), $($RoleGroupMember.TrusteeOriginalIdentity))"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.MRG{0:0000000}.txt' -f $RoleGroupMemberId))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Management Role Group Members"";""$($($GrantorPrimarySMTP), $($RoleGroupMember.RoleGroup), $($RoleGroupMember.TrusteeOriginalIdentity))"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllManagementRoleGroupMembers = $AllManagementRoleGroupMembers
                        AllRecipients                 = $AllRecipients
                        tempQueue                     = $tempQueue
                        ExportFile                    = $ExportFile
                        ExportTrustees                = $ExportTrustees
                        AllRecipientsSmtpToIndex      = $AllRecipientsSmtpToIndex
                        ErrorFile                     = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                     = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                    = $PSScriptRoot
                        ExportFromOnPrem              = $ExportFromOnPrem
                        VerbosePreference             = $VerbosePreference
                        DebugPreference               = $DebugPreference
                        TrusteeFilter                 = $TrusteeFilter
                        UTF8Encoding                  = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} management role group members to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all management role group members have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Forwarders
    Write-Host
    Write-Host "Forwarders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportForwarders) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if ($x -in $GrantorsToConsider) {
                $tempQueue.enqueue($x)
            }
        }

        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

        Write-Host '  Prepare imported data'
        $AllForwardingAddresses | ForEach-Object {
            try {
                try {
                    $GrantorIndex = $null
                    $GrantorIndex = $AllRecipientsIdentityGuidToIndex[$($_.Identity.ObjectGuid.Guid)]
                } catch {
                }

                if ($GrantorIndex -ge 0) {
                    $Grantor = $AllRecipients[$GrantorIndex]

                    if ($_.ForwardingAddress.ObjectGuid.Guid) {
                        try {
                            $TrusteeIndex = $null
                            $TrusteeIndex = $AllRecipientsIdentityGuidToIndex[$($_.ForwardingAddress.ObjectGuid.Guid)]
                        } catch {
                        }

                        if ($TrusteeIndex -ge 0) {
                            $AllRecipients[$GrantorIndex].ForwardingAddress = $AllRecipients[$TrusteeIndex].PrimarySmtpAddress.Address
                        }
                    }

                    $Grantor.ForwardingSmtpAddress = $_.ForwardingSmtpAddress.SmtpAddress
                    $Grantor.DeliverToMailboxAndForward = $_.DeliverToMailboxAndForward
                }
            } catch {
            }
        }

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $AllRecipientsSmtpToIndex,
                            $tempQueue,
                            $ExportFile,
                            $ExportTrustees,
                            $ErrorFile,
                            $DebugFile,
                            $ScriptPath,
                            $ExportFromOnPrem,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Forwarders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                foreach ($ForwarderType in ('ExternalEmailAddress', 'ForwardingAddress', 'ForwardingSmtpAddress')) {
                                    try {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$($Grantor.$ForwarderType)]
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $Trustee = $AllRecipients[$index]
                                        } else {
                                            $Trustee = $Grantor.$ForwarderType
                                        }

                                        if ($Grantor.$ForwarderType) {
                                            if ($TrusteeFilter) {
                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                    continue
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                    $ExportFileresult.add((('"' + ((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        $('Forward_' + $ForwarderType + $(if ((-not $Grantor.DeliverToMailboxAndForward) -or ($ForwarderType -ieq 'ExternalEmailAddress')) { '_ForwardOnly' } else { '_DeliverAndForward' } )),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $Grantor.$ForwarderType,
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                }
                                            }
                                        }
                                    } catch {
                                        """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Forwarders"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                                    }
                                }

                                if ($ExportFileResult) {
                                    $ExportFileResult | Sort-Object -Unique | Add-Content ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";""Forwarders"";""$($GrantorPrimarySMTP)"";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients            = $AllRecipients
                        AllRecipientsSmtpToIndex = $AllRecipientsSmtpToIndex
                        tempQueue                = $tempQueue
                        ExportFile               = $ExportFile
                        ExportTrustees           = $ExportTrustees
                        ErrorFile                = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath               = $PSScriptRoot
                        ExportFromOnPrem         = $ExportFromOnPrem
                        VerbosePreference        = $VerbosePreference
                        DebugPreference          = $DebugPreference
                        TrusteeFilter            = $TrusteeFilter
                        UTF8Encoding             = $UTF8Encoding
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '  Not all recipients have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Add-Content $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobErrorFile -Encoding $UTF8Encoding | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }
} catch {
    Write-Host 'Unexpected error. Exiting.'
    """$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"";"""";"""";""$($_ | Out-String)""" -replace '(?<!;|^)"(?!;|$)', '""' | Add-Content -Path $ErrorFile -Encoding $UTF8Encoding -Force
} finally {
    Write-Host
    Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
        Disconnect-ExchangeOnline -Confirm:$false
        Remove-Module ExchangeOnlineManagement
    }

    if ($ExchangeSession) {
        Remove-PSSession -Session $ExchangeSession
    }

    Write-Host "  Runspaces and RunspacePool @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($runspaces) {
        foreach ($runspace in $runspaces) {
            $runspace.PowerShell.Dispose()
        }
    }
    if ($RunspacePool) {
        $RunspacePool.dispose()
    }

    Write-Host "  Temporary result files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $RecipientFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))

    if ($RecipientFiles.count -gt 0) {
        Write-Host ('    {0:0000000} files to check. Done (in steps of {1:0000000}):' -f $RecipientFiles.count, $UpdateInterval)
        Write-Host ('      {0:0000000} @{1}@' -f 0, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))

        $lastCount = 1

        foreach ($RecipientFile in $RecipientFiles) {
            if ($RecipientFile.length -gt 0) {
                Get-Content $RecipientFile -Encoding $UTF8Encoding | Sort-Object -Unique | Add-Content $ExportFile -Encoding $UTF8Encoding -Force
            }

            Remove-Item $RecipientFile -Force

            if (($lastCount % $UpdateInterval -eq 0) -or ($lastcount -eq $RecipientFiles.count)) {
                Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $lastcount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                if ($lastcount -eq $RecipientFiles.count) { Write-Host }
            }

            $lastCount++
        }
    }

    if ($ErrorFile) {
        Write-Host "  Temporary error files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $JobErrorFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))

        if ($JobErrorFiles.count -gt 0) {
            Write-Host ('    {0:0000000} files to check. Done (in steps of {1:0000000}):' -f $JobErrorFiles.count, $UpdateInterval)
            Write-Host ('      {0:0000000} @{1}@' -f 0, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))

            $x = @()

            $lastCount = 1

            foreach ($JobErrorFile in $JobErrorFiles) {
                if ($JobErrorFile.length -gt 0) {
                    $x += @(Get-Content $JobErrorFile -Encoding $UTF8Encoding)
                }

                Remove-Item $JobErrorFile -Force

                if (($lastCount % $UpdateInterval -eq 0) -or ($lastcount -eq $JobErrorFiles.count)) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $lastcount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    if ($lastcount -eq $JobErrorFiles.count) { Write-Host }
                }

                $lastCount++
            }

            $x | Sort-Object -Unique | Add-Content $ErrorFile -Encoding $UTF8Encoding -Force
        }
    }

    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($DebugFile) {
        $null = Stop-Transcript
        Start-Sleep -Seconds 1

        $JobDebugFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))

        if ($JobDebugFiles.count -gt 0) {
            foreach ($JobDebugFile in $JobDebugFiles) {
                if ($JobDebugFile.length -gt 0) {
                    Get-Content $JobDebugFile | Add-Content -Path $DebugFile -Encoding $UTF8Encoding -Force
                }

                Remove-Item $JobDebugFile -Force
            }
        }
    }

    Remove-Variable * -ErrorAction SilentlyContinue
    [System.GC]::Collect() # garbage collection
    Start-Sleep 1
}
