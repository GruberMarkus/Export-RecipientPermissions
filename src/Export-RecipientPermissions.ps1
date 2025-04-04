<#
.SYNOPSIS
Export-RecipientPermissions XXXVersionStringXXX
Document, filter and compare Exchange permissions: Mailbox access rights, mailbox folder permissions, public folder permissions, send as, send on behalf, managed by, linked master accounts, forwarders, management role groups, distribution group members
.DESCRIPTION
Document, filter and compare Exchange permissions:
- Mailbox access rights
- Mailbox folder permissions
- Public Folder permissions
- Send As
- Send On Behalf
- Managed By
- Moderated By
- Linked Master Accounts
- Forwarders
- Sender restrictions
- Resource delegates
- Group members
- Management Role group members

Easens the move to the cloud, as permission dependencies beyond the supported cross-premises permissions (https://docs.microsoft.com/en-us/Exchange/permissions) can easily be identified and even be represented graphically (sample code included).

Compare exports from different times to detect permission changes (sample code included).

.LINK
Github: https://github.com/GruberMarkus/Export-RecipientPermissions


.PARAMETER ExportFromOnPrem
Export from On-Prem or from Exchange Online
$true for export from on-prem, $false for export from Exchange Online
Default: $false


.PARAMETER ExchangeConnectionUriList
Exchange remote PowerShell URIs to connect to
For on-prem installations, list all Exchange Server Remote PowerShell URIs the script can use
For Exchange Online, use 'https://outlook.office365.com/powershell-liveid/' or the URI specific to your cloud environment
Default:
    If ExportFromOnPrem ist set to false: 'https://outlook.office365.com/powershell-liveid/'
    If ExportFromOnPrem ist set to true: 'http://<server>/powershell' for each Exchange server with the mailbox server role

.PARAMETER ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential
Credentials for Exchange connection
Username and password are stored as encrypted secure strings, if UseDefaultCredential is not enabled


.PARAMETER ExchangeOnlineConnectionParameters
This hashtable will be passed as parameter to Connect-ExchangeOnline
All values are allowed, but CommandName is set by the script. By default, ShowBanner and ShowProgress are set to $false; SkipLoadingFormatData to $true.


.PARAMETER ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal
Maximum Exchange, AD and local sessions/jobs running in parallel
Watch CPU and RAM usage, and your Exchange throttling policy - frequent connection errors indicate that the values are set too high
Default values:
    ParallelJobsExchange: $ExchangeConnectionUriList.count
    ParallelJobsAD: 50
    ParallelJobsLocal: 50


.PARAMETER RecipientProperties
Recipient properties to import.
Be aware that these properties are not queried with '`Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Recipient -ResultSize Unlimited | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties)`', but with a simple '`Get-Recipient`'.
These properties are available for GrantorFilter and TrusteeFilter.
Properties that are always included: 'Identity', 'DistinguishedName', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'Name', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'UserFriendlyName', 'LinkedMasterAccount'


.PARAMETER GrantorFilter
Only check grantors where the filter criteria matches $true.
The variable $Grantor has all attributes defined by '`RecipientProperties`. For example:
  .DistinguishedName
  .RecipientType, .RecipientTypeDetails
  .DisplayName
  .Identity
  .PrimarySmtpAddress
  .EmailAddresses
    This attribute is an array. Code example:
      $GrantorFilter = "if ((`$Grantor.EmailAddresses -ilike 'smtp:AddressA@example.com') -or (`$Grantor.EmailAddresses -ilike 'smtp:Test*@example.com')) { `$true } else { `$false }"
  .UserFriendlyName: User account holding the mailbox in the "<NetBIOS domain name>\<sAMAccountName>" format
  .ManagedBy
    This attribute is an array. Code example:
      $GrantorFilter = "foreach (`$XXXSingleManagedByXXX in `$Grantor.ManagedBy) { if (`$XXXSingleManagedByXXX -iin @(
                          'example.com/OU1/OU2/ObjectA',
                          'example.com/OU3/OU4/ObjectB',
      )) { `$true; break } }"
  On-prem only:
    .LinkedMasterAccount: Linked Master Account in the "<NetBIOS domain name>\<sAMAccountName>" format
Set to $null or '' to define all recipients as grantors to consider
Example: "`$Grantor.primarysmtpaddress -ilike '*@example.com'"
Default: $null


.PARAMETER TrusteeFilter
Only report trustees where the filter criteria matches $true.
If the trustee matches a recipient, the available attributes are the same as for GrantorFilter, only the reference variable is $Trustee instead of $Grantor.
If the trustee does not match a recipient (because it no longer exists, for exampe), $Trustee is just a string. In this case, the export shows the following:
  Column "Trustee Original Identity" contains the trustee description string as reported by Exchange
  Columns "Trustee Primary SMTP" and "Trustee Display Name" are empty
Example: "`$Trustee.primarysmtpaddress -ilike '*@example.com'"
Default: $null


.PARAMETER ExportFileFilter
Only report results where the filter criteria matches $true.
This filter works against every single row of the results found. ExportFile will only contain lines where this filter returns $true.
The $ExportFileLine variable contains an object with the header names from $ExportFile as string properties
    'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Folder', 'Permission', 'Allow/Deny', 'Inherited', 'InheritanceType', 'Trustee Original Identity', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment'
    When GUIDs are exported, additional attributes are available: 'Grantor Exchange GUID', 'Grantor AD ObjectGUID', 'Trustee Exchange GUID', 'Trustee AD ObjectGUID'
Example: "`$ExportFileLine.'Trustee Environment' -ieq 'On-Prem'"
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
Default: $false


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
The virtual right 'MemberRecurse' or 'MemberDirect' is used in the export file
GrantorFilter does not apply to the export of management role groups, but TrusteeFilter and ExportFileFilter do
Default: $true


.PARAMETER ExportForwarders
Export forwarders configured on recipients
Default: $true


.PARAMETER ExportModerators
Exports the virtual rights 'ModeratedBy' and 'ModeratedByBypass', listing all users and groups which are configured as moderators for a recipient or can bypass moderation.
Only works for recipients with moderation enabled.
Default: $true


.PARAMETER ExportRequireAllSendersAreAuthenticated
Exports the virtual right 'RequireAllSendersAreAuthenticated' with the trustee 'NT AUTHORITY\Authenticated Users' for each recipient which is configured to only receive messages from authenticated (internal) senders.
Default: $true


.PARAMETER ExportAcceptMessagesOnlyFrom
Exports the virtual right 'AcceptMessagesOnlyFrom' for each recipient which is configured to only receive messages from selected (internal) senders.
The attributes 'AcceptMessagesOnlyFrom' and 'AcceptMessagesOnlyFromDLMembers' are exported as the same virtual right 'AcceptMessagesOnlyFrom'.
Default: $true


.PARAMETER ExportResourceDelegates
Exports information about who is allowed or denied to book resources (rooms or equipment) and to accept or reject booking requests.
The following virtual rights are exported:
- ResourceDelegate
- ResourcePolicyDelegate_AllBookInPolicy
- ResourcePolicyDelegate_AllRequestInPolicy
- ResourcePolicyDelegate_AllRequestOutOfPolicy
- ResourcePolicyDelegate_BookInPolicy
- ResourcePolicyDelegate_RequestInPolicy
- ResourcePolicyDelegate_RequestOutOfPolicy
ResourcePolicyDelegate_AllBookInPolicy, ResourcePolicyDelegate_AllRequestinPolicy, ResourcePolicyDelegate_AllRequestOutOfPolicy: 'Everyone' is used as trustee.
ResourcePolicyDelegate_BookInPolicy, ResourcePolicyDelegate_RequestInPolicy, ResourcePolicyDelegate_RequestOutOfPolicy: Each of these virtual rights is reported even when the corresponding 'All'-right is enabled.
Default: $true


.PARAMETER ExportDistributionGroupMembers
Export distribution group members, including nested groups and dynamic groups
The parameter ExpandGroups can be used independently:
  ExpandGroups acts when a group is used as trustee: It adds every recurse member of the group as a separate trustee entry
  ExportDistributionGroupMembers exports the distribution group as grantor, which the recurse members as trustees
Valid values: 'None', 'All', 'OnlyTrustees'
  'None': Distribution group members are not exported Parameter ExpandGroups can still be used.
  'All': Members of all distribution groups are exported, parameter GrantorFilter is considerd
  'OnlyTrustees': Only export members of those distribution groups that are used as trustees, even when they are excluded via GrantorFilter
Default: 'None'


.PARAMETER ExportGroupMembersRecurse
When disabled, only direct members of groups are exported, and the virtual right 'MemberDirect' is used in the export file.
When enabled, recursive members of groups are exported, and the virtual right 'MemberRecurse' is used in the export file.
Default: $false


.PARAMETER ExportGuids
When enabled, the export contains the Exchange GUID and the AD ObjectGUID for each grantor and trustee
Default: $false

.PARAMETER ExportSids
When enabled, the export contains the SID (Security Identifier) for each grantor and trustee
Default: $false

.PARAMETER ExpandGroups
Expand trustee groups to their members, including nested groups and dynamic groups
This may drastically increase script run time and file size
This works for all groups, mail-enabled or not
The original permission is still documented, with one additional line for each member of the group used as trustee
  For each member of the group, 'Trustee Original Identity' is preserved, but the string '     [MemberRecurse] ' or '     [MemberDirect] ' (the leading whitespace consists of five spaces for sorting reasons) and the original identity of the recurse member
  The other trustee properties are the ones of the recurse member
TrusteeFilter is applied to trustee groups as well as to their finally expanded individual members
  Nested groups are expanded to individual members, but TrusteeFilter is not applied to the nested group
Default value: $false


.PARAMETER ExportGrantorsWithNoPermissions
Per default, Export-RecipientPermissions only exports grantors which have set at least one permission for at least one trustee.
If all grantors should be exported, set this parameter to $true.
If enabled, a grantor that that not grant any permission is included in the list with the following columns: "Grantor Primary SMTP", "Grantor Display Name", "Grantor Recipient Type", "Grantor Environment". The other columns for this recipient are empty.
Default value: $false


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
Default: '.\export\Export-RecipientPermissions_Error.csv'


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
    [boolean]$ExportFromOnPrem = $false,
    [uri[]]$ExchangeConnectionUriList = $(
        if ($ExportFromOnPrem) {
            try {
                $search = New-Object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$(([ADSI]'LDAP://RootDse').configurationNamingContext)")
                $search.Filter = '(&(objectClass=msExchExchangeServer)(msExchCurrentServerRoles:1.2.840.113556.1.4.803:=2))' # all Exchange servers with the mailbox role
                $search.PageSize = 1000
                [void]$search.PropertiesToLoad.Add('networkaddress')
                @((($search.FindAll().properties.networkaddress | Where-Object { $_ -ilike 'ncacn_ip_tcp:*' }) -ireplace '^ncacn_ip_tcp:', 'http://' -ireplace '$', '/powershell') | Sort-Object -Unique)
            } catch {
                @()
            }
        } else {
            'https://outlook.office365.com/powershell-liveid/'
        }
    ),
    [boolean]$UseDefaultCredential = $false,
    [string]$ExchangeCredentialUsernameFile = '.\Export-RecipientPermissions_CredentialUsername.txt',
    [string]$ExchangeCredentialPasswordFile = '.\Export-RecipientPermissions_CredentialPassword.txt',
    [hashtable]$ExchangeOnlineConnectionParameters = @{ Credential = $null },
    [int]$ParallelJobsExchange = $(if ($ExportFromOnPrem) { $ExchangeConnectionUriList.count } else { 3 }),
    [int]$ParallelJobsAD = 50,
    [int]$ParallelJobsLocal = 50,
    [string[]]$RecipientProperties = @(),
    [string]$GrantorFilter = $null,
    [string]$TrusteeFilter = $null,
    [string]$ExportFileFilter = $null,
    [boolean]$ExportMailboxAccessRights = $true,
    [boolean]$ExportMailboxAccessRightsSelf = $false,
    [boolean]$ExportMailboxAccessRightsInherited = $false,
    [boolean]$ExportMailboxFolderPermissions = $false,
    [boolean]$ExportMailboxFolderPermissionsAnonymous = $true,
    [boolean]$ExportMailboxFolderPermissionsDefault = $true,
    [boolean]$ExportMailboxFolderPermissionsOwnerAtLocal = $false,
    [boolean]$ExportMailboxFolderPermissionsMemberAtLocal = $false,
    [string[]]$ExportMailboxFolderPermissionsExcludeFoldertype = ('audits'),
    [boolean]$ExportSendAs = $true,
    [boolean]$ExportSendAsSelf = $false,
    [boolean]$ExportSendOnBehalf = $true,
    [boolean]$ExportManagedBy = $true,
    [boolean]$ExportLinkedMasterAccount = $true,
    [boolean]$ExportPublicFolderPermissions = $false,
    [boolean]$ExportPublicFolderPermissionsAnonymous = $true,
    [boolean]$ExportPublicFolderPermissionsDefault = $true,
    [string[]]$ExportPublicFolderPermissionsExcludeFoldertype = (''),
    [boolean]$ExportForwarders = $true,
    [boolean]$ExportModerators = $true,
    [boolean]$ExportRequireAllSendersAreAuthenticated = $true,
    [boolean]$ExportAcceptMessagesOnlyFrom = $true,
    [boolean]$ExportResourceDelegates = $true,
    [boolean]$ExportManagementRoleGroupMembers = $false,
    [ValidateSet('None', 'All', 'OnlyTrustees')]$ExportDistributionGroupMembers = 'None',
    [boolean]$ExportGroupMembersRecurse = $false,
    [boolean]$ExpandGroups = $false,
    [boolean]$ExportGuids = $false,
    [boolean]$ExportSids = $false,
    [boolean]$ExportGrantorsWithNoPermissions = $false,
    [ValidateSet('All', 'OnlyValid', 'OnlyInvalid')]$ExportTrustees = 'All',
    [parameter(dontshow = $true)][string]$ExportTimestamp = $((Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz') -ireplace ':', '' -ireplace '-', ''),
    [string]$ExportFile = ".\export\Export-RecipientPermissions_$($ExportTimestamp)_Result.csv",
    [string]$ErrorFile = ".\export\Export-RecipientPermissions_$($ExportTimestamp)_Error.csv",
    [string]$DebugFile = '',
    [int][ValidateRange(1, [int]::MaxValue)]$UpdateInterval = 100
)


#
# Do not change anything from here on.
#


$ConnectExchange = {
    param(
        [int]$RetryMaximum = 3,

        [scriptblock]$ScriptBlock = $(
            if ($ExportFromOnPrem) {
                { Get-SecurityPrincipal -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction Stop }
            } else {
                { Get-SecurityPrincipal -ResultSize 1 -WarningAction SilentlyContinue }
            }
        ),

        [scriptblock]$ScriptBlockAfter = $null,

        [switch]$NoReturnValue,

        [switch]$Disconnect,

        [string]$Indent = 2
    )

    [bool]$ConnectExchangeTempVariableStopLoop = $false
    [int]$ConnectExchangeTempVariableRetryCount = 0
    [int]$ConnectExchangeTempVariableSleepTime = 0
    [scriptblock]$ConnectExchangeTempVariableScriptBlockPre = { Set-AdServerSettings -ViewEntireForest 1 -ErrorAction Stop }

    $DisconnectExchange = {
        # Disconnect current session
        Write-Host "$(' '*$Indent)DisconnectExchange, remove existing connection"

        if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
            Disconnect-ExchangeOnline -Confirm:$false
            Remove-Module -Name 'ExchangeOnlineManagement' -Force
        }

        if (($ExportFromOnPrem -eq $true)) {
            if ($ConnectExchangeExchangeSessionModule) {
                $ConnectExchangeExchangeSessionModule | Remove-Module
            }

            Get-PSSession | Where-Object { $_.Name -eq 'ConnectExchangeExchangeSession' } | Remove-PSSession
        }
    }


    if ($Disconnect) {
        . ([scriptblock]::Create($DisconnectExchange))
        return
    }


    if (-not $ConnectExchangeTempVariableConnectionUri) {
        $ConnectExchangeTempVariableRetryCount++
    }

    while (($ConnectExchangeTempVariableStopLoop -eq $false) -and ($ConnectExchangeTempVariableRetryCount -le $RetryMaximum)) {
        try {
            # Calculate SleepTime in case of error
            $ConnectExchangeTempVariableSleepTime = (60 * $ConnectExchangeTempVariableRetryCount) + 15


            if ($ConnectExchangeTempVariableRetryCount -gt 0) {
                # Get (new) connection URI
                if ((-not $ConnectExchangeTempVariableConnectionUri) -or ($ConnectExchangeTempVariableRetryCount -gt 1)) {
                    $ConnectExchangeTempVariableConnectionUri = $tempConnectionUriQueue.dequeue()
                }


                # Create new session
                Write-Host "$(' '*$Indent)ConnectExchange, try $($ConnectExchangeTempVariableRetryCount)/$($RetryMaximum), start connecting to '$($ConnectExchangeTempVariableConnectionUri)'"

                if ($ExportFromOnPrem -eq $true) {
                    if ($UseDefaultCredential) {
                        $ConnectExchangeExchangeSession = New-PSSession -ConfigurationName 'Microsoft.Exchange' -ConnectionUri $ConnectExchangeTempVariableConnectionUri -Authentication Kerberos -AllowRedirection -Name 'ConnectExchangeExchangeSession' -ErrorAction Stop
                    } else {
                        $ConnectExchangeExchangeSession = New-PSSession -ConfigurationName 'Microsoft.Exchange' -ConnectionUri $ConnectExchangeTempVariableConnectionUri -Credential $ExchangeCredential -Authentication Kerberos -AllowRedirection -Name 'ConnectExchangeExchangeSession' -ErrorAction Stop
                    }
                } else {
                    if ($ExchangeOnlineConnectionParameters.ContainsKey('Credential')) {
                        $ExchangeOnlineConnectionParameters['Credential'] = $ExchangeCredential
                    }

                    if (-not $ExchangeOnlineConnectionParameters.ContainsKey('SkipLoadingFormatData')) {
                        $ExchangeOnlineConnectionParameters['SkipLoadingFormatData'] = $true
                    }

                    if (-not $ExchangeOnlineConnectionParameters.ContainsKey('SkipLoadingCmdletHelp')) {
                        $ExchangeOnlineConnectionParameters['SkipLoadingCmdletHelp'] = $true
                    }

                    if (-not $ExchangeOnlineConnectionParameters.ContainsKey('ShowBanner')) {
                        $ExchangeOnlineConnectionParameters['ShowBanner'] = $false
                    }

                    if (-not $ExchangeOnlineConnectionParameters.ContainsKey('ShowProgress')) {
                        $ExchangeOnlineConnectionParameters['ShowProgress'] = $false
                    }

                    $ExchangeOnlineConnectionParameters['ConnectionUri'] = $ConnectExchangeTempVariableConnectionUri

                    Import-Module 'ExchangeOnlineManagement' -Force -DisableNameChecking -ErrorAction Stop

                    Connect-ExchangeOnline @ExchangeOnlineConnectionParameters # -Prefix $ConnectExchangeTempVariableCmdletPrefix
                }


                # Execute $ScriptBlockPre to test new connection
                if ($ExportFromOnPrem) {
                    if ($ConnectExchangeTempVariableScriptBlockPre) {
                        $ConnectExchangeTempVariableScriptBlockExecute = $ConnectExchangeTempVariableScriptBlockPre
                        $null = Invoke-Command -Session $ConnectExchangeExchangeSession -HideComputerName -ScriptBlock ([scriptblock]::create($ExecutionContext.InvokeCommand.ExpandString($ConnectExchangeTempVariableScriptBlockExecute))) -ErrorAction Stop
                    }
                }
            }


            # Execute $ScriptBlock
            if ($ExportFromOnPrem) {
                if ($ScriptBlock) {
                    $ConnectExchangeTempVariableScriptBlockExecute = $ScriptBlock
                    $ConnectExchangeTempReturnValue = Invoke-Command -Session $ConnectExchangeExchangeSession -HideComputerName -ScriptBlock ([scriptblock]::create($ExecutionContext.InvokeCommand.ExpandString($ConnectExchangeTempVariableScriptBlockExecute))) -ErrorAction Stop
                }

                if ($ScriptBlockAfter) {
                    . ([scriptblock]::Create($ScriptBlockAfter))
                }
            } else {
                if ($ScriptBlock) {
                    $ConnectExchangeTempVariableScriptBlockExecute = $ScriptBlock
                    $ConnectExchangeTempReturnValue = . ([scriptblock]::create($ExecutionContext.InvokeCommand.ExpandString($ConnectExchangeTempVariableScriptBlockExecute)))
                }

                if ($ScriptBlockAfter) {
                    . ([scriptblock]::Create($ScriptBlockAfter))
                }
            }

            if ($NoReturnValue -eq $false) {
                return $ConnectExchangeTempReturnValue
            } else {
                $ConnectExchangeTempVariableStopLoop = $true
            }
        } catch {
            if ($ConnectExchangeTempVariableRetryCount -eq 0) {
                Write-Host "$(' '*$Indent)ConnectExchange, try $($ConnectExchangeTempVariableRetryCount)/$($RetryMaximum), failed, next try in $($ConnectExchangeTempVariableSleepTime) seconds"

                . ([scriptblock]::Create($DisconnectExchange))
                Start-Sleep -Seconds $ConnectExchangeTempVariableSleepTime

                $ConnectExchangeTempVariableRetryCount++
            } elseif ($ConnectExchangeTempVariableRetryCount -lt $RetryMaximum) {
                Write-Host "$(' '*$Indent)ConnectExchange, try $($ConnectExchangeTempVariableRetryCount)/$($RetryMaximum), failed, next try in $($ConnectExchangeTempVariableSleepTime) seconds`r`n    ScriptBlock: $($ConnectExchangeTempVariableScriptBlockExecute)`r`n    ScriptBlock expanded: $($ExecutionContext.InvokeCommand.ExpandString($ConnectExchangeTempVariableScriptBlockExecute))`r`n    Error: $($_ | Out-String)"

                . ([scriptblock]::Create($DisconnectExchange))
                Start-Sleep -Seconds $ConnectExchangeTempVariableSleepTime

                $ConnectExchangeTempVariableRetryCount++
            } else {
                $ConnectExchangeTempVariableError = $_

                . ([scriptblock]::Create($DisconnectExchange))

                throw "$(' '*$Indent)ConnectExchange, try $($ConnectExchangeTempVariableRetryCount)/$($RetryMaximum), failed, giving up`r`n    ScriptBlock: $($ConnectExchangeTempVariableScriptBlockExecute)`r`n    ScriptBlock expanded: $($ExecutionContext.InvokeCommand.ExpandString($ConnectExchangeTempVariableScriptBlockExecute))`r`n    Error: $($ConnectExchangeTempVariableError | Out-String)"
            }
        }
    }
}


$FilterGetMember = {
    filter GetMemberRecurse {
        param(
            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]$GroupToCheck,
            [switch]$DoNotResetGetMemberRecurseTempLoopProtection,
            [switch]$DirectMembersOnly
        )

        if (-not $DoNotResetGetMemberRecurseTempLoopProtection.IsPresent) {
            $script:GetMemberRecurseTempLoopProtection = @()
        }

        # Determine GroupToCheckType
        $GroupToCheckType = 'Unknown'

        if ($AllRecipientsIdentityToIndex.containskey($GroupToCheck)) {
            $AllRecipientsIndex = $AllRecipientsIdentityToIndex[$GroupToCheck]
        } else {
            $AllRecipientsIndex = $null
        }

        if ($AllGroupsIdentityToIndex.containskey($GroupToCheck)) {
            $AllGroupsIndex = $AllGroupsIdentityToIndex[$GroupToCheck]
        } else {
            $AllGroupsIndex = $null
        }

        if (($AllRecipientsIndex -ge 0) -and ($AllGroupsIndex -ge 0)) {
            If (($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails -ilike '*Group') -or ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails -ilike 'Group*')) {
                $GroupToCheckType = 'Group'
            } else {
                $GroupToCheckType = 'Unknown'
            }
        } elseif (($AllRecipientsIndex -ge 0) -and ($AllGroupsIndex -lt 0)) {
            if ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails -ilike 'DynamicDistributionGroup') {
                $GroupToCheckType = 'DynamicDistributionGroup'
            } elseif (($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails -inotlike '*Group') -and ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails -inotlike 'Group*')) {
                $GroupToCheckType = 'User'
            } else {
                $GroupToCheckType = 'Unknown'
            }
        } elseif (($AllRecipientsIndex -lt 0) -and ($AllGroupsIndex -ge 0)) {
            $GroupToCheckType = 'ManagementRoleGroup'
        } else {
            $GroupToCheckType = 'Unknown'
        }


        if ($GroupToCheckType -ieq 'User') {
            $AllRecipientsIndex
        } elseif (($GroupToCheckType -ieq 'Group') -or ($GroupToCheckType -ieq 'ManagementRoleGroup')) {
            foreach ($member in $AllGroups[$AllGroupsIndex].members) {
                if ($DirectMembersOnly.IsPresent) {
                    if ($AllRecipientsIdentityToIndex.ContainsKey($member)) {
                        $AllRecipientsIdentityToIndex[$member]
                    } else {
                        # $member is not known in $AllRecipients
                        "NotARecipient:$($member)"
                    }
                } else {
                    if (($AllGroupsIdentityToIndex.ContainsKey($member) -or $AllRecipientsIdentityToIndex.ContainsKey($member))) {
                        if ($member -notin $script:GetMemberRecurseTempLoopProtection) {
                            $script:GetMemberRecurseTempLoopProtection += $member
                            $member | GetMemberRecurse -DoNotResetGetMemberRecurseTempLoopProtection
                        }
                    } else {
                        # $member is neither known in $AllRecipients, nor in $AllGroups
                        "NotARecipient:$($member)"
                    }
                }
            }
        } elseif ($GroupToCheckType -ieq 'DynamicDistributionGroup') {
            if ($ExportFromOnPrem) {
                $DynamicGroup = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-DynamicDistributionGroup -identity $($AllRecipients[$AllRecipientsIndex].Guid.Guid) -WarningAction SilentlyContinue -ErrorAction Stop }
                $members = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Recipient -RecipientPreviewFilter $('''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent($($DynamicGroup.RecipientFilter)) + '''') -OrganizationalUnit $('''' + $($DynamicGroup.RecipientContainer) + '''') -WarningAction SilentlyContinue -ErrorAction Stop }
            } else {
                $members = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-DynamicDistributionGroupMember -identity $($AllRecipients[$AllRecipientsIndex].Guid.Guid) -WarningAction SilentlyContinue -ErrorAction Stop }
            }

            if ($members) {
                $members = @(@($members).Identity)
            }

            foreach ($member in @($members)) {
                if ($DirectMembersOnly.IsPresent) {
                    if ($AllRecipientsIdentityToIndex.ContainsKey($member)) {
                        $AllRecipientsIdentityToIndex[$member]
                    } else {
                        # $member is not known in $AllRecipients
                        "NotARecipient:$($member)"
                    }
                } else {
                    if (($AllGroupsIdentityToIndex.ContainsKey($member) -or $AllRecipientsIdentityToIndex.ContainsKey($member))) {
                        if ($member -notin $script:GetMemberRecurseTempLoopProtection) {
                            $script:GetMemberRecurseTempLoopProtection += $member
                            $member | GetMemberRecurse -DoNotResetGetMemberRecurseTempLoopProtection
                        }
                    } else {
                        # $member is neither known in $AllRecipients, nor in $AllGroups
                        "NotARecipient:$($member)"
                    }
                }
            }
        } else {
            if (($AllRecipientsIndex -ge 0) -and ($AllRecipients[$AllRecipientsIndex].UserFriendlyName)) {
                "NotARecipient:$($AllRecipients[$AllRecipientsIndex].UserFriendlyName)"
            } elseif (($AllGroupsIndex -ge 0) -and (($AllGroups[$AllGroupsIndex].DisplayName) -or ($AllGroups[$AllGroupsIndex].Name) -or ($AllGroups[$AllGroupsIndex].DistinguishedName))) {
                "NotARecipient:$(@(($AllGroups[$AllGroupsIndex].DistinguishedName), ($AllGroups[$AllGroupsIndex].Name), ($AllGroups[$AllGroupsIndex].DisplayName), 'Warning: No valid info found') | Where-Object { $_ } | Select-Object -First 1)"
            } else {
                "NotARecipient:$($GroupToCheck)"
            }
        }
    }
}


try {
    $OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

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
            $null = Stop-Transcript
        } catch {
        }

        $null = Start-Transcript -LiteralPath $DebugFile -Force
    }


    Clear-Host
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


    Write-Host
    Write-Host "Script notes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host '  Script : Export-RecipientPermissions'
    Write-Host '  Version: XXXVersionStringXXX'
    Write-Host '  Web    : https://github.com/GruberMarkus/Export-RecipientPermissions'
    Write-Host "  License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)"


    Write-Host
    Write-Host "Script environment and parameters @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
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
        $ErrorFileHeader = @(
            'Timestamp',
            'Task',
            'TaskDetail',
            'Error'
        )

        if (Test-Path $ErrorFile) {
            Remove-Item -LiteralPath $ErrorFile -Force -WarningAction SilentlyContinue -ErrorAction Stop
        }

        try {
            foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
                Remove-Item -LiteralPath $JobErrorFile -Force
            }
        } catch {
        }

        $null = New-Item -Path $ErrorFile -Force

        ('"' + ($ErrorFileHeader -join '";"') + '"') | Out-File $ErrorFile -Encoding $UTF8Encoding -Force
    }


    if ($DebugFile) {
        try {
            foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
                Remove-Item -LiteralPath $JobDebugFile -Force
            }
        } catch {
        }
    }


    if ($ExportFile) {
        if (Test-Path $ExportFile) {
            Remove-Item -LiteralPath $ExportFile -Force
        }

        foreach ($JobExportFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
            Remove-Item -LiteralPath $JobExportFile -Force
        }

        $ExportFileHeader = @(
            'Grantor Primary SMTP',
            'Grantor Display Name',
            $(if ($ExportGuids) { 'Grantor Exchange GUID' } else { '' }),
            $(if ($ExportGuids) { 'Grantor AD ObjectGUID' } else { '' }),
            $(if ($ExportSids) { 'Grantor SID' } else { '' }),
            'Grantor Recipient Type',
            'Grantor Environment',
            'Folder',
            'Permission',
            'Allow/Deny',
            'Inherited',
            'InheritanceType',
            'Trustee Original Identity',
            'Trustee Primary SMTP',
            'Trustee Display Name',
            $(if ($ExportGuids) { 'Trustee Exchange GUID' } else { '' }),
            $(if ($ExportGuids) { 'Trustee AD ObjectGUID' } else { '' }),
            $(if ($ExportSids) { 'Trustee SID' } else { '' }),
            'Trustee Recipient Type',
            'Trustee Environment'
        )

        $ExportFileHeaderIndexes = @(
            $(
                for ($i = 0; $i -lt $ExportFileHeader.Count; $i++) {
                    if ($ExportFileHeader[$i]) {
                        $i
                    }
                }
            )
        )

        $ExportFileHeader = $ExportFileHeader[$ExportFileHeaderIndexes]

        ('"' + ($ExportFileHeader -join '";"') + '"') | Out-File $ExportFile -Encoding $UTF8Encoding -Force
    }


    $tempConnectionUriQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new(10000))
    while ($tempConnectionUriQueue.count -le 100000) {
        foreach ($ExchangeConnectionUri in $ExchangeConnectionUriList) {
            $tempConnectionUriQueue.Enqueue($ExchangeConnectionUri.AbsoluteUri)
        }
    }

    if ($RecipientProperties -contains '*') {
        $RecipientProperties = @('*')
    } else {
        @('Identity', 'DistinguishedName', 'ExchangeGuid', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'Name', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'WhenSoftDeleted', 'Guid') | ForEach-Object {
            if ($RecipientProperties -inotcontains $_) {
                $RecipientProperties += $_
            }
        }

        if ($ExportForwarders) {
            @('ExternalEmailAddress') | ForEach-Object {
                if ($RecipientProperties -inotcontains $_) {
                    $RecipientProperties += $_
                }
            }
        }
    }

    $RecipientProperties = @($RecipientProperties | Sort-Object -Unique)


    # Not supported by Get-Recipient -Properties, but required for Select-Object
    $RecipientPropertiesExtended = $RecipientProperties

    @('UserFriendlyName', 'LinkedMasterAccount', 'IsTrustee') | ForEach-Object {
        if ($RecipientPropertiesExtended -inotcontains $_) {
            $RecipientPropertiesExtended += $_
        }
    }

    if ($ExportForwarders) {
        @('ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward') | ForEach-Object {
            if ($RecipientPropertiesExtended -inotcontains $_) {
                $RecipientPropertiesExtended += $_
            }
        }
    }

    if ($ExportModerators) {
        @('ModerationEnabled', 'ModeratedBy', 'ModeratedByBypass') | ForEach-Object {
            if ($RecipientPropertiesExtended -inotcontains $_) {
                $RecipientPropertiesExtended += $_
            }
        }
    }

    if ($ExportRequireAllSendersAreAuthenticated) {
        @('RequireAllSendersAreAuthenticated') | ForEach-Object {
            if ($RecipientPropertiesExtended -inotcontains $_) {
                $RecipientPropertiesExtended += $_
            }
        }
    }

    if ($ExportAcceptMessagesOnlyFrom) {
        @('AcceptMessagesOnlyFromSendersOrMembers') | ForEach-Object {
            if ($RecipientPropertiesExtended -inotcontains $_) {
                $RecipientPropertiesExtended += $_
            }
        }
    }

    if ($ExportResourceDelegates) {
        @('ResourceDelegates', 'AllBookInPolicy', 'BookInPolicy', 'AllRequestInPolicy', 'RequestInPolicy', 'AllRequestOutOfPolicy', 'RequestOutOfPolicy', 'LegacyExchangeDN') | ForEach-Object {
            if ($RecipientPropertiesExtended -inotcontains $_) {
                $RecipientPropertiesExtended += $_
            }
        }
    }

    if ($ExpandGroups -or $ExportManagementRoleGroupMembers -or ($ExportDistributionGroupMembers -ine 'None')) {
        @('RecipientFilter', 'RecipientContainer') | ForEach-Object {
            if ($RecipientPropertiesExtended -inotcontains $_) {
                $RecipientPropertiesExtended += $_
            }
        }
    }

    $RecipientPropertiesExtended = @($RecipientPropertiesExtended | Sort-Object -Unique)


    # Credentials
    Write-Host
    Write-Host "Exchange credentials @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if (
        (($ExportFromOnPrem -eq $true) -and ($UseDefaultCredential -eq $false)) -or
        (($ExportFromOnPrem -eq $false) -and ($ExchangeOnlineConnectionParameters.ContainsKey('Credential')))
    ) {
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
        $ExchangeCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ([PSCredential]::new('X', (Get-Content -LiteralPath $ExchangeCredentialUsernameFile -Encoding $UTF8Encoding | ConvertTo-SecureString)).GetNetworkCredential().Password), (Get-Content -LiteralPath $ExchangeCredentialPasswordFile -Encoding $UTF8Encoding | ConvertTo-SecureString)
    } else {
        Write-Host '  Use current credential'
        $ExchangeCredential = $null
    }

    # Connect to Exchange
    Write-Host
    Write-Host "Connect to Exchange for import operations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue


    # Import recipients
    Write-Host
    Write-Host "Import recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host '  Enumerate possible RecipientTypeDetails values'
    try {
        # Get-EXORecipient does not (yet) return allowed RecipientTypeDetails,
        #   so Get-Recipient is used for Exchange on-prem and Exchange Online
        if ($ExportFromOnPrem) {
            $null = Invoke-Command -Session $ConnectExchangeExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientTypeDetails '!!!Fail!!!' -resultsize 1 -ErrorAction Stop -WarningAction silentlycontinue } -ErrorAction Stop
        } else {
            $null = Get-Recipient -RecipientTypeDetails '!!!Fail!!!' -resultsize 1 -ErrorAction Stop -WarningAction silentlycontinue
        }
    } catch {
        $null = $error[0].exception -match '(?!.*: )(.*)(")$'
        $RecipientTypeDetailsListUnchecked = $matches[1].trim() -split ', ' | Where-Object { $_ } | Sort-Object -Unique
    }

    $RecipientTypeDetailsList = @()

    foreach ($RecipientTypeDetail in $RecipientTypeDetailsListUnchecked) {
        # Get-EXORecipient is extremly slow when querying for non-existing RecipienttypeDetails
        #   so Get-Recipient is used for Exchange on-prem and Exchange Online
        try {
            if ($ExportFromOnPrem) {
                $null = Invoke-Command -Session $ConnectExchangeExchangeSession -HideComputerName -ScriptBlock ([scriptblock]::create($ExecutionContext.InvokeCommand.ExpandString({ Get-Recipient -RecipientTypeDetails $($RecipientTypeDetail) -resultsize 1 -ErrorAction Stop -WarningAction silentlycontinue }))) -ErrorAction Stop
            } else {
                $null = Get-Recipient -RecipientTypeDetails $RecipientTypeDetail -resultsize 1 -ErrorAction Stop -WarningAction silentlycontinue
            }

            $RecipientTypeDetailsList += $RecipientTypeDetail
        } catch {
        }
    }

    Write-Host "  Default recipients, grouped by RecipientTypeDetails and first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $Filters = @()

    foreach ($tempChar in @([char[]](0..255) -clike '[A-Z0-9]')) {
        $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('(name -like ''' + $($tempChar) + '*'')') + ''''
    }

    $filters += '''' + ($filters.trimstart('''').trimend('''') -join ' -and ').replace('(name -like ''', '(name -notlike ''') + ''''

    $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

    foreach ($RecipientTypeDetail in $RecipientTypeDetailsList) {
        foreach ($Filter in $Filters) {
            $tempQueue.enqueue((, $RecipientTypeDetail, $Filter))
        }
    }

    $RecipientTypeDetailsList = $null
    $Filters = $null

    $AllRecipients = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

    $tempQueueCount = $tempQueue.count

    $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

    Write-Host "    Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

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
                        $ConnectExchange,
                        $DebugFile,
                        $DebugPreference,
                        $ErrorFile,
                        $ExchangeCredential,
                        $ExchangeOnlineConnectionParameters,
                        $ExportFromOnPrem,
                        $RecipientProperties,
                        $RecipientPropertiesExtended,
                        $ScriptPath,
                        $tempConnectionUriQueue,
                        $tempQueue,
                        $UseDefaultCredential,
                        $UTF8Encoding,
                        $VerbosePreference
                    )

                    try {
                        $DebugPreference = 'Continue'

                        Set-Location $ScriptPath

                        if ($DebugFile) {
                            $null = Start-Transcript -LiteralPath $DebugFile -Force
                        }

                        Write-Host "Import Recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                        . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                        while ($tempQueue.count -gt 0) {
                            try {
                                $QueueArray = $tempQueue.dequeue()
                            } catch {
                                continue
                            }

                            Write-Host "RecipientTypeDetails '$($QueueArray[0])', Filter '$($QueueArray[1])' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            try {
                                if ($ExportFromOnPrem) {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Recipient -RecipientTypeDetails $($QueueArray[0]) -Filter $($QueueArray[1]) -Properties $($RecipientProperties -join ', ') -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
                                } else {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-EXORecipient -RecipientTypeDetails $($QueueArray[0]) -Filter $($QueueArray[1]) -Properties $($RecipientProperties -join ', ') -ResultSize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
                                }

                                if ($x) {
                                    $x = @($x | Select-Object -Property $RecipientPropertiesExtended)
                                    $AllRecipients.AddRange(@($x))
                                    Write-Host "  $($x.count) recipients"
                                } else {
                                    Write-Host '  0 recipients'
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                'Import Recipients',
                                                "RecipientTypeDetails '$($QueueArray[0])', Filter '$($QueueArray[1])'",
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                            }
                        }
                    } catch {
                        (
                            '"' + (
                                @(
                                    (
                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                        'Import Recipients',
                                        '',
                                        $($_ | Out-String)
                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                            ) + '"'
                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    } finally {
                        . ([scriptblock]::create($ConnectExchange)) -Disconnect

                        if ($DebugFile) {
                            $null = Stop-Transcript
                            Start-Sleep -Seconds 1
                        }
                    }
                }
            ).AddParameters(
                @{
                    AllRecipients                      = $AllRecipients
                    ConnectExchange                    = $ConnectExchange
                    DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                    DebugPreference                    = $DebugPreference
                    ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                    ExchangeCredential                 = $ExchangeCredential
                    ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                    ExportFromOnPrem                   = $ExportFromOnPrem
                    RecipientProperties                = $RecipientProperties
                    RecipientPropertiesExtended        = $RecipientPropertiesExtended
                    ScriptPath                         = $PSScriptRoot
                    tempConnectionUriQueue             = $tempConnectionUriQueue
                    tempQueue                          = $tempQueue
                    UseDefaultCredential               = $UseDefaultCredential
                    UTF8Encoding                       = $UTF8Encoding
                    VerbosePreference                  = $VerbosePreference
                }
            )

            $Handle = $Powershell.BeginInvoke()

            $temp = '' | Select-Object PowerShell, Handle, Object
            $temp.PowerShell = $PowerShell
            $temp.Handle = $Handle
            [void]$runspaces.Add($Temp)
        }

        Write-Host ('    {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

        $lastCount = -1
        while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
            Start-Sleep -Seconds 1
            $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
            for ($x = $lastCount; $x -le $done; $x++) {
                if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                    Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                    if ($x -eq 0) {
                        Write-Host
                    }
                    $lastCount = $x
                }
            }
        }

        Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

        if ($tempQueue.count -ne 0) {
            Write-Host '      Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
        }

        foreach ($runspace in $runspaces) {
            # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
            # $runspace.PowerShell.Stop()
            $runspace.PowerShell.Dispose()
        }

        $RunspacePool.Close()
        $RunspacePool.Dispose()
        'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

        if ($DebugFile) {
            $null = Stop-Transcript
            Start-Sleep -Seconds 1
            foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                Remove-Item -LiteralPath $JobDebugFile -Force
            }

            $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
        }

        if ($ErrorFile) {
            foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                Remove-Item -LiteralPath $JobErrorFile -Force
            }
        }

        [GC]::Collect(); Start-Sleep -Seconds 1
    }

    Write-Host ('    {0:0000000} recipients found' -f $($AllRecipients.count))

    Write-Host "  Additional recipients of specific types @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host "    Single-thread Exchange operations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    Write-Host '      Migration mailboxes'
    # Get-EXOMailbox misses several options (such as -Migration), so Get-Mailbox is still used for Exchange Online sometimes
    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -Migration -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
    if ($x) {
        $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
    }

    if ($ExportFromOnPrem) {
        Write-Host '      Arbitration mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -Arbitration -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }

        Write-Host '      AuditLog mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -AuditLog -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }

        Write-Host '      AuxAuditLog mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -AuxAuditLog -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }

        Write-Host '      Monitoring mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -Monitoring -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }

        Write-Host '      RemoteArchive mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -RemoteArchive -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }
    } else {
        Write-Host '      Inactive mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-EXOMailbox -InactiveMailboxOnly -PropertySets All -ResultSize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }

        Write-Host '      Softdeleted mailboxes'
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-EXOMailbox -SoftDeletedMailbox -PropertySets All -ResultSize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipients.AddRange(@($x | Select-Object -Property $RecipientPropertiesExtended))
        }
    }

    Write-Host ('  {0:0000000} total recipients found' -f $($AllRecipients.count))

    Write-Host "  Sort list by PrimarySmtpAddress @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipients.TrimToSize()

    $x = @(@($AllRecipients) | Where-Object { $_.PrimarySmtpAddress } | Sort-Object -Property PrimarySmtpAddress)

    $AllRecipients.clear()
    $AllRecipients.AddRange(@($x))
    $x = $null
    $AllRecipients.TrimToSize()

    Write-Host '  Create lookup hashtables'
    Write-Host "    First character (lowercase) of name attribute for future wildcard searches @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $WildcardSearchStrings = @(@(for ($x = 0; $x -lt $AllRecipients.count; $x++) {
 (-join $AllRecipients[$x].Name[0]).ToLower()
            }) | Select-Object -Unique)

    Write-Host "    DistinguishedName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsDnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].distinguishedname) {
            if ($AllRecipientsDnToIndex.ContainsKey($(($AllRecipients[$x]).distinguishedname))) {
                Write-Host "      '$(($AllRecipients[$x]).distinguishedname)' is not unique" -ForegroundColor Yellow
                $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $null
            } else {
                $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $x
            }
        }
    }

    Write-Host "    Identity (CanonicalName) to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsIdentityToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].identity) {
            if ($AllRecipientsIdentityToIndex.ContainsKey($(($AllRecipients[$x]).identity))) {
                Write-Host "      '$(($AllRecipients[$x]).identity)' is not unique" -ForegroundColor Yellow
                $AllRecipientsIdentityToIndex[$(($AllRecipients[$x]).identity)] = $null
            } else {
                $AllRecipientsIdentityToIndex[$(($AllRecipients[$x]).identity)] = $x
            }
        }
    }

    Write-Host "    IdentityGuid to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsIdentityGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].Guid.Guid) {
            if ($AllRecipientsIdentityGuidToIndex.ContainsKey($(($AllRecipients[$x]).Guid.Guid))) {
                Write-Host "      '$(($AllRecipients[$x]).Guid.Guid)' is not unique" -ForegroundColor Yellow
                $AllRecipientsIdentityGuidToIndex[$(($AllRecipients[$x]).Guid.Guid)] = $null
            } else {
                $AllRecipientsIdentityGuidToIndex[$(($AllRecipients[$x]).Guid.Guid)] = $x
            }
        }
    }

    Write-Host "    ExchangeGuid to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsExchangeGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if (($AllRecipients[$x].ExchangeGuid.Guid) -and ($AllRecipients[$x].ExchangeGuid.Guid -ine '00000000-0000-0000-0000-000000000000')) {
            if ($AllRecipientsExchangeGuidToIndex.ContainsKey($(($AllRecipients[$x]).ExchangeGuid.Guid))) {
                Write-Host "      '$(($AllRecipients[$x]).ExchangeGuid.Guid)' is not unique" -ForegroundColor Yellow
                $AllRecipientsExchangeGuidToIndex[$(($AllRecipients[$x]).ExchangeGuid.Guid)] = $null
            } else {
                $AllRecipientsExchangeGuidToIndex[$(($AllRecipients[$x]).ExchangeGuid.Guid)] = $x
            }
        }
    }

    Write-Host "    EmailAddresses to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsSmtpToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.EmailAddresses.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].EmailAddresses) {
            foreach ($EmailAddress in (@(@($AllRecipients[$x].EmailAddresses | Where-Object { $_.StartsWith('smtp:', 'CurrentCultureIgnoreCase') }) | ForEach-Object { $_ -ireplace '^smtp:', '' }))) {
                if ($AllRecipientsSmtpToIndex.ContainsKey($EmailAddress)) {
                    Write-Host "      '$($EmailAddress)' is not unique" -ForegroundColor Yellow
                    $AllRecipientsSmtpToIndex[$EmailAddress] = $null
                } else {
                    $AllRecipientsSmtpToIndex[$EmailAddress] = $x
                }
            }
        }
    }

    Write-Host "    DisplayName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsDisplaynameToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].DisplayName) {
            if ($AllRecipientsDisplaynameToIndex.ContainsKey($AllRecipients[$x].DisplayName)) {
                Write-Verbose "      '$($AllRecipients[$x].DisplayName)' is not unique"
                $AllRecipientsDisplaynameToIndex[$AllRecipients[$x].DisplayName] = $null
            } else {
                $AllRecipientsDisplaynameToIndex[$AllRecipients[$x].DisplayName] = $x
            }
        }
    }


    # Import recipient permissions (SendAs)
    Write-Host
    Write-Host "Import Send As permissions from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (($ExportFromOnPrem -eq $false) -and ($ExportSendAs -eq $true)) {
        Write-Host '  Single-thread Exchange operation'
        $AllRecipientsSendas = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count * 2))

        if ($ExportFromOnPrem) {
            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-RecipientPermission -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        } else {
            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-RecipientPermission -ResultSize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        }

        if ($x) {
            $x = @($x | Select-Object identity, trustee, accessrights, accesscontroltype, isinherited, inheritancetype, trusteeSidString)

            if ($x) {
                $AllRecipientsSendas.AddRange(@($x))
            }
        }

        $AllRecipientsSendas.TrimToSize()
        Write-Host ('  {0:0000000} Send As permissions found' -f $($AllRecipientsSendas.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import Send On Behalf from cloud
    Write-Host
    Write-Host "Import Send On Behalf permissions from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (($ExportFromOnPrem -eq $false) -and ($ExportSendOnBehalf -eq $true)) {
        Write-Host '  Single-thread Exchange operation'
        $AllRecipientsSendonbehalf = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count * 2))

        Write-Host "  Mailboxes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        # Get-EXOMailbox does not support the GrantSendOnBehalfTo filter, so Get-Mailbox is used
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -filter 'GrantSendOnBehalfTo -ne `$null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipientsSendonbehalf.AddRange(@($x | Select-Object identity, grantsendonbehalfto))
        }

        Write-Host "  Distribution groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-DistributionGroup -filter 'GrantSendOnBehalfTo -ne `$null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipientsSendonbehalf.AddRange(@($x | Select-Object identity, grantsendonbehalfto))
        }

        Write-Host "  Dynamic Distribution Groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-DynamicDistributionGroup -filter 'GrantSendOnBehalfTo -ne `$null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipientsSendonbehalf.AddRange(@($x | Select-Object identity, grantsendonbehalfto))
        }

        Write-Host "  Unified Groups (Microsoft 365 Groups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-UnifiedGroup -filter 'GrantSendOnBehalfTo -ne `$null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipientsSendonbehalf.AddRange(@($x | Select-Object identity, grantsendonbehalfto))
        }

        Write-Host "  Mail-enabled Public Folders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailPublicfolder -filter 'GrantSendOnBehalfTo -ne `$null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllRecipientsSendonbehalf.AddRange(@($x | Select-Object identity, grantsendonbehalfto))
        }

        $AllRecipientsSendonbehalf.TrimToSize()
        Write-Host ('  {0:0000000} Send On Behalf permissions found' -f $($AllRecipientsSendonbehalf.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import mailbox databases
    Write-Host
    Write-Host "Import mailbox databases @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportFromOnPrem) {
        Write-Host '  Single-thread Exchange operation'

        $AllMailboxDatabases = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailboxDatabase -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllMailboxDatabases.AddRange(@($x | Select-Object -Property Guid, ProhibitSendQuota))
        }

        $AllMailboxDatabases.TrimToSize()
        Write-Host ('  {0:0000000} mailbox databases found' -f $($AllMailboxDatabases.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import Public Folders
    Write-Host
    Write-Host "Import Public Folders and their content mailboxes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportPublicFolderPermissions) {
        Write-Host '  Single-thread Exchange operation'

        $AllPublicFolders = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))

        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-PublicFolder -recurse -ErrorAction Stop -WarningAction silentlycontinue }
        if ($x) {
            $AllPublicFolders.AddRange(@($x | Select-Object -Property EntryId, ContentMailboxGuid, MailEnabled, MailRecipientGuid, FolderClass, FolderPath | Sort-Object -Property FolderPath ))
        }

        $AllPublicFolders.TrimToSize()
        Write-Host ('  {0:0000000} Public Folders found' -f $($AllPublicFolders.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import additional forwarding addresses
    Write-Host
    Write-Host "Import additional forwarding addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportForwarders) {
        Write-Host '  Single-thread Exchange operation'

        $AdditionalForwardingAddresses = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

        # Get-EXOMailbox does not support the ForwardingAddress filter, so Get-Mailbox is used
        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -filter '(ForwardingAddress -ne `$null) -or (ForwardingSmtpAddress -ne `$null)' -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue }
        if ($x) {
            $AdditionalForwardingAddresses.AddRange(@($x | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward))
        }

        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailPublicFolder -filter '(ForwardingAddress -ne `$null)' -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue }
        if ($x) {
            $AdditionalForwardingAddresses.AddRange(@($x | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward))
        }

        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailUser -filter '(ForwardingAddress -ne `$null)' -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue }
        if ($x) {
            $AdditionalForwardingAddresses.AddRange(@($x | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward))
        }

        if ($ExportFromOnPrem) {
            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-RemoteMailbox -filter '(ForwardingAddress -ne `$null)' -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue }
            if ($x) {
                $AdditionalForwardingAddresses.AddRange(@($x | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward))
            }
        }

        $AdditionalForwardingAddresses.TrimToSize()

        Write-Host ('  {0:0000000} additional forwarding addresses found' -f $($AdditionalForwardingAddresses.count))

        Write-Host '  Convert imported data'
        foreach ($Recipient in $AllRecipients) {
            if ($Recipient.ExternalEmailAddress) {
                if ($Recipient.RecipientTypeDetails -ieq 'PublicFolder') {
                    $Recipient.ExternalEmailAddress = $null
                } else {
                    $Recipient.ExternalEmailAddress = ($Recipient.ExternalEmailAddress -ireplace '^smtp:', '').ToLower()
                }
            }
        }

        $AdditionalForwardingAddresses | ForEach-Object {
            try {
                try {
                    $GrantorIndex = $null
                    $GrantorIndex = $AllRecipientsIdentityToIndex[$($_.Identity)]
                } catch {
                }

                if ($GrantorIndex -ge 0) {
                    $Grantor = $AllRecipients[$GrantorIndex]

                    if ($_.ForwardingAddress) {
                        try {
                            $TrusteeIndex = $null
                            $TrusteeIndex = $AllRecipientsIdentityToIndex[$($_.ForwardingAddress)]
                        } catch {
                        }

                        if ($TrusteeIndex -ge 0) {
                            $Grantor.ForwardingAddress = $AllRecipients[$TrusteeIndex].PrimarySmtpAddress
                        } else {
                            $Grantor.ForwardingAddress = $_.ForwardingAddress
                        }
                    }

                    if ($_.ForwardingSmtpAddress) {
                        $Grantor.ForwardingSmtpAddress = ($_.ForwardingSmtpAddress -ireplace '^smtp:', '').ToLower()
                    }

                    if ($_.DeliverToMailboxAndForward) {
                        $Grantor.DeliverToMailboxAndForward = $_.DeliverToMailboxAndForward
                    }
                }
            } catch {
            }
        }

        $AdditionalForwardingAddresses = $null
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Disconnect from Exchange
    Write-Host
    Write-Host "Single-thread Exchange operations completed, remove connection to Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    . ([scriptblock]::create($ConnectExchange)) -Disconnect

    [GC]::Collect(); Start-Sleep -Seconds 1


    # Import LinkedMasterAccounts
    Write-Host
    Write-Host "Import LinkedMasterAccounts of each mailbox by database @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

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
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import LinkedMasterAccounts @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $MailboxDatabaseGuid = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "MailboxDatabaseGuid $($MailboxDatabaseGuid) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $mailboxes = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Mailbox -database $($MailboxDatabaseGuid) -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }

                                    if ($mailboxes) {
                                        $mailboxes = @($mailboxes | Select-Object -Property Guid, LinkedMasterAccount)
                                    } else {
                                        $mailboxes = @()
                                    }

                                    foreach ($mailbox in $mailboxes) {
                                        if ($mailbox.LinkedMasterAccount) {
                                            try {
                                                ($AllRecipients[$($AllRecipientsIdentityGuidToIndex[$($mailbox.Guid.Guid)])]).LinkedMasterAccount = $mailbox.LinkedMasterAccount
                                            } catch {
                                                (
                                                    '"' + (
                                                        @(
                                                            (
                                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                                'Import LinkedMasterAccounts',
                                                                "Mailbox Identity GUID $($mailbox.Guid.Guid)",
                                                                $($_ | Out-String)
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                                    ) + '"'
                                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import LinkedMasterAccounts',
                                                    "Mailbox database GUID $(MailboxDatabaseGuid)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import LinkedMasterAccounts',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityGuidToIndex   = $AllRecipientsIdentityGuidToIndex
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} databases to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all databases have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        Write-Host ('  {0:0000000} Linked Master Accounts found' -f $(@($AllRecipients | Where-Object { $_.LinkedMasterAccount }).count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import security principals
    Write-Host
    Write-Host "Import security principals, grouped by first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if (
            ($ExportMailboxAccessRights) -or
            ($ExportSendAs) -or
            ($ExportLinkedMasterAccount -and $ExportFromOnPrem) -or
            ($ExportManagementRoleGroupMembers) -or
            ($ExportDistributionGroupMembers -ieq 'All') -or
            ($ExportDistributionGroupMembers -ieq 'OnlyTrustees') -or
            ($ExpandGroups) -or
            ($ExportGuids) -or
            ($ExportSids)
    ) {
        $AllSecurityPrincipals = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

        $Filters = @()

        foreach ($tempChar in @([char[]](0..255) -clike '[A-Z0-9]')) {
            $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('(name -like ''' + $($tempChar) + '*'')') + ''''
        }

        $filters += '''' + ($filters.trimstart('''').trimend('''') -join ' -and ').replace('(name -like ''', '(name -notlike ''') + ''''

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        foreach ($Filter in $Filters) {
            $tempQueue.enqueue($Filter)
        }

        $Filters = $null

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
                            $AllSecurityPrincipals,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import security principals @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $filter = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Filter '$($filter)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock $(
                                        if ($ExportFromOnPrem) {
                                            { Get-SecurityPrincipal -Filter $($filter) -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction stop }
                                        } else {
                                            { Get-SecurityPrincipal -Filter $($filter) -ResultSize Unlimited -WarningAction SilentlyContinue }
                                        }) -ScriptBlockAfter {
                                        if (@($ConnectExchangeTempReturnValue).count -ne @($ConnectExchangeTempReturnValue).guid.guid.count) {
                                            throw 'Error: Some security principals do not have a GUID, which must be a query error.'
                                        }
                                    }

                                    if ($x) {
                                        $x = @($x | Select-Object Sid, UserFriendlyName, Guid, DistinguishedName)

                                        $AllSecurityPrincipals.AddRange(@($x))

                                        Write-Host "  $($x.count) security principals"
                                    } else {
                                        Write-Host '  0 security principals'
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import security principals',
                                                    "Filter '$($filter)'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import security principals',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllSecurityPrincipals              = $AllSecurityPrincipals
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        $AllSecurityPrincipals.TrimToSize()
        Write-Host ('  {0:0000000} security principals found' -f $($AllSecurityPrincipals.count))

        Write-Host '  Add UserFriendlyName to AllRecipients'
        for ($x = 0; $x -lt $AllSecurityPrincipals.Count; $x++) {
            if ($AllRecipientsIdentityGuidToIndex.containskey($(($AllSecurityPrincipals[$x]).guid.guid))) {
                ($AllRecipients[$($AllRecipientsIdentityGuidToIndex[$(($AllSecurityPrincipals[$x]).guid.guid)])]).UserFriendlyName = ($AllSecurityPrincipals[$x]).UserFriendlyName
            }
        }

        Write-Host '  Create lookup hashtables'
        Write-Host "    SID to index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllSecurityPrincipalsSidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllSecurityPrincipals.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllSecurityPrincipals.Count; $x++) {
            if (($AllSecurityPrincipals[$x]).Sid) {
                if ($AllSecurityPrincipalsSidToIndex.ContainsKey(($AllSecurityPrincipals[$x]).Sid)) {
                    Write-Host "      '$(($AllSecurityPrincipals[$x]).Sid)' is not unique" -ForegroundColor Yellow
                    $AllSecurityPrincipalsSidToIndex[$(($AllSecurityPrincipals[$x]).Sid)] = $null
                } else {
                    $AllSecurityPrincipalsSidToIndex[$(($AllSecurityPrincipals[$x]).Sid)] = $x
                }
            }
        }

        Write-Host "    ObjectGuid to index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllSecurityPrincipalsObjectguidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllSecurityPrincipals.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllSecurityPrincipals.Count; $x++) {
            if (($AllSecurityPrincipals[$x]).Guid.Guid) {
                if ($AllSecurityPrincipalsObjectguidToIndex.ContainsKey(($AllSecurityPrincipals[$x]).Guid.Guid)) {
                    Write-Host "      '$(($AllSecurityPrincipals[$x]).Guid.Guid)' is not unique" -ForegroundColor Yellow
                    $AllSecurityPrincipalsObjectguidToIndex[$(($AllSecurityPrincipals[$x]).Guid.Guid)] = $null
                } else {
                    $AllSecurityPrincipalsObjectguidToIndex[$(($AllSecurityPrincipals[$x]).Guid.Guid)] = $x
                }
            }
        }

        Write-Host "    DistinguishedName to index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllSecurityPrincipalsDnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllSecurityPrincipals.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllSecurityPrincipals.Count; $x++) {
            if (($AllSecurityPrincipals[$x]).DistinguishedName) {
                if ($AllSecurityPrincipalsDnToIndex.ContainsKey(($AllSecurityPrincipals[$x]).DistinguishedName)) {
                    Write-Host "      '$(($AllSecurityPrincipals[$x]).DistinguishedName)' is not unique" -ForegroundColor Yellow
                    $AllSecurityPrincipalsDnToIndex[$(($AllSecurityPrincipals[$x]).DistinguishedName)] = $null
                } else {
                    $AllSecurityPrincipalsDnToIndex[$(($AllSecurityPrincipals[$x]).DistinguishedName)] = $x
                }
            }
        }

        Write-Host "    UserFriendlyName to index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllSecurityPrincipalsUfnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllSecurityPrincipals.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllSecurityPrincipals.Count; $x++) {
            if (($AllSecurityPrincipals[$x]).UserFriendlyName) {
                if ($AllSecurityPrincipalsUfnToIndex.ContainsKey(($AllSecurityPrincipals[$x]).UserFriendlyName)) {
                    Write-Host "      '$(($AllSecurityPrincipals[$x]).UserFriendlyName)' is not unique" -ForegroundColor Yellow
                    $AllSecurityPrincipalsUfnToIndex[$(($AllSecurityPrincipals[$x]).UserFriendlyName)] = $null
                } else {
                    $AllSecurityPrincipalsUfnToIndex[$(($AllSecurityPrincipals[$x]).UserFriendlyName)] = $x
                }
            }
        }

        Write-Host "    DisplayName to index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllSecurityPrincipalsDisplaynameToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllSecurityPrincipals.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllSecurityPrincipals.Count; $x++) {
            if (($AllSecurityPrincipals[$x]).DisplayName) {
                if ($AllSecurityPrincipalsDisplaynameToIndex.ContainsKey(($AllSecurityPrincipals[$x]).DisplayName)) {
                    Write-Verbose "      '$(($AllSecurityPrincipals[$x]).DisplayName)' is not unique"
                    $AllSecurityPrincipalsDisplaynameToIndex[$(($AllSecurityPrincipals[$x]).DisplayName)] = $null
                } else {
                    $AllSecurityPrincipalsDisplaynameToIndex[$(($AllSecurityPrincipals[$x]).DisplayName)] = $x
                }
            }
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Create lookup hashtables for UserFriendlyName and LinkedMasterAccount
    Write-Host
    Write-Host "Create lookup hashtables for UserFriendlyName and LinkedMasterAccount @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    Write-Host "  UserFriendlyName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsUfnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $Recipient = $AllRecipients[$x]
        if ($Recipient.userfriendlyname) {
            if ($AllRecipientsUfnToIndex.ContainsKey($($Recipient.userfriendlyname))) {
                if ($AllRecipientsUfnToIndex[$($Recipient.userfriendlyname)]) {
                    Write-Host "      '$($Recipient.userfriendlyname)' is not unique" -ForegroundColor Yellow
                }

                $AllRecipientsUfnToIndex[$Recipient.userfriendlyname] = $null
            } else {
                $AllRecipientsUfnToIndex[$Recipient.userfriendlyname] = $x
            }
        }
    }

    Write-Host "  LinkedMasterAccount to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $AllRecipientsLinkedmasteraccountToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if (($AllRecipients[$x]).LinkedMasterAccount) {
            if ($AllRecipientsLinkedmasteraccountToIndex.ContainsKey($(($AllRecipients[$x]).LinkedMasterAccount))) {
                # Same LinkedMasterAccount defined multiple time - set index to $null
                if ($AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)]) {
                    Write-Verbose "    '$(($AllRecipients[$x]).LinkedMasterAccount)' used not only once: '$($AllRecipients[$($AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)])].PrimarySmtpAddress)'"
                }

                Write-Verbose "    '$(($AllRecipients[$x]).LinkedMasterAccount)' used not only once: '$(($AllRecipients[$x]).PrimarySmtpAddress)'"

                $AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)] = $null
            } else {
                $AllRecipientsLinkedmasteraccountToIndex[$(($AllRecipients[$x]).LinkedMasterAccount)] = $x
            }
        }
    }


    # Import moderators
    Write-Host
    Write-Host "Import moderators, grouped by RecipientTypeDetails and first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportModerators) {
        $Filters = @()

        foreach ($WildcardSearchString in $WildcardSearchStrings) {
            $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('((name -like ''' + $WildcardSearchString + '*'') -and (ModerationEnabled -eq `$true))') + ''''
        }

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        foreach ($Cmdlet in @(@(
                    'Get-DistributionGroup',
                    'Get-DynamicDistributionGroup',
                    'Get-Mailbox', # Get-EXOMailbox can't yet handle the filter defined before
                    'Get-MailContact',
                    'Get-MailPublicFolder',
                    'Get-MailUser',
                    $(if ($ExportFromOnprem -eq $true) {
                            'Get-RemoteMailbox'
                        }), # available on-prem only
                    $(if ($ExportFromOnprem -eq $false) {
                            'Get-UnifiedGroup'
                        }) # Exchange Online only
                ) | Where-Object { $_ })) {
            foreach ($Filter in $Filters) {
                $tempQueue.enqueue(@($Cmdlet, $Filter))
            }
        }
        $Filters = $null

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
                            $AllRecipientsIdentityToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import moderators @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $QueueArray = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock ([scriptblock]::Create("$($QueueArray[0]) -Filter $($QueueArray[1]) -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue"))

                                    if ($x) {
                                        $x = @($x | Select-Object -Property Identity, ModerationEnabled, ModeratedBy, BypassModerationFromSendersOrMembers)

                                        Write-Host "  $($x.count) recipients"

                                        foreach ($ModeratedRecipient in @($x)) {
                                            try {
                                                $index = $null
                                                $index = $AllRecipientsIdentityToIndex[$($ModeratedRecipient.Identity)]
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $AllRecipients[$index].ModerationEnabled = $ModeratedRecipient.ModerationEnabled

                                                $AllRecipients[$index].ModeratedBy = $ModeratedRecipient.ModeratedBy
                                                if (($ModeratedRecipient.ModerationEnabled -eq $true) -and (-not $ModeratedRecipient.ModeratedBy)) {
                                                    $AllRecipients[$index].ModeratedBy = $AllRecipients[$index].ManagedBy
                                                }

                                                $AllRecipients[$index].ModeratedByBypass = $ModeratedRecipient.BypassModerationFromSendersOrMembers
                                            }
                                        }
                                    } else {
                                        Write-Host '  0 recipients'
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import moderators',
                                                    "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import moderators',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityToIndex       = $AllRecipientsIdentityToIndex
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        Write-Host ('  {0:0000000} recipients with moderation settings found' -f $(($AllRecipients | Where-Object { $_.ModerationEnabled -eq $true }).count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import RequireAllSendersAreAuthenticated
    Write-Host
    Write-Host "Import RequireAllSendersAreAuthenticated, grouped by RecipientTypeDetails and first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportRequireAllSendersAreAuthenticated) {
        $Filters = @()

        foreach ($WildcardSearchString in $WildcardSearchStrings) {
            $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('((name -like ''' + $WildcardSearchString + '*'') -and (RequireAllSendersAreAuthenticated -eq `$true))') + ''''
        }

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        foreach ($Cmdlet in @(@(
                    'Get-DistributionGroup',
                    'Get-DynamicDistributionGroup',
                    'Get-Mailbox', # Get-EXOMailbox can't yet handle the filter
                    'Get-MailContact',
                    'Get-MailPublicFolder',
                    'Get-MailUser',
                    $(if ($ExportFromOnprem -eq $true) {
                            'Get-RemoteMailbox'
                        }),
                    $(if ($ExportFromOnprem -eq $false) {
                            'Get-UnifiedGroup'
                        }),
                    'Get-SecurityPrincipal'
                ) | Where-Object { $_ })) {
            foreach ($Filter in $Filters) {
                $tempQueue.enqueue(@($Cmdlet, $Filter))
            }
        }
        $Filters = $null

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
                            $AllRecipientsIdentityToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import RequireAllSendersAreAuthenticated @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $QueueArray = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock ([scriptblock]::Create("$($QueueArray[0]) -Filter $($QueueArray[1]) -ResultSize Unlimited $(if (($($QueueArray[0]) -ieq 'Get-SecurityPrincipal') -and ($ExportFromOnPrem -eq $false)) { '' } else { ' -ErrorAction Stop' }) -WarningAction SilentlyContinue"))

                                    if ($x) {
                                        $x = @($x | Select-Object Identity)

                                        Write-Host "  $($x.count) recipients"

                                        foreach ($Recipient in @($x)) {
                                            try {
                                                $index = $null
                                                $index = $AllRecipientsIdentityToIndex[$($Recipient.Identity)]
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $AllRecipients[$index].RequireAllSendersAreAuthenticated = $true
                                            }
                                        }
                                    } else {
                                        Write-Host '  0 recipients'
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import RequireAllSendersAreAuthenticated',
                                                    "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import RequireAllSendersAreAuthenticated',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityToIndex       = $AllRecipientsIdentityToIndex
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        Write-Host ('  {0:0000000} recipients with RequireAllSendersAreAuthenticated found' -f $(($AllRecipients | Where-Object { $_.RequireAllSendersAreAuthenticated }).count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import AcceptMessagesOnlyFrom
    Write-Host
    Write-Host "Import AcceptMessagesOnlyFrom, grouped by RecipientTypeDetails and first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportAcceptMessagesOnlyFrom) {
        $Filters = @()

        foreach ($WildcardSearchString in $WildcardSearchStrings) {
            $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('((name -like ''' + $WildcardSearchString + '*'') -and ((AcceptMessagesOnlyFrom -ne `$null) -or (AcceptMessagesOnlyFromDLMembers -ne `$null)))') + ''''
        }

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        foreach ($Cmdlet in @(@(
                    'Get-DistributionGroup',
                    'Get-DynamicDistributionGroup',
                    'Get-Mailbox', # Get-EXOMailbox can't yet handle the filter
                    'Get-MailContact',
                    'Get-MailPublicFolder',
                    'Get-MailUser',
                    $(if ($ExportFromOnprem -eq $true) {
                            'Get-RemoteMailbox'
                        }),
                    $(if ($ExportFromOnprem -eq $false) {
                            'Get-UnifiedGroup'
                        })
                ) | Where-Object { $_ })) {
            foreach ($Filter in $Filters) {
                $tempQueue.enqueue(@($Cmdlet, $Filter))
            }
        }
        $Filters = $null

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
                            $AllRecipientsIdentityToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import AcceptMessagesOnlyFrom @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $QueueArray = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock ([scriptblock]::Create("$($QueueArray[0]) -Filter $($QueueArray[1]) -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue"))

                                    if ($x) {
                                        $x = @($x | Select-Object Identity, AcceptMessagesOnlyFromSendersOrMembers)

                                        Write-Host "  $($x.count) recipients"

                                        foreach ($Recipient in @($x)) {
                                            try {
                                                $index = $null
                                                $index = $AllRecipientsIdentityToIndex[$($Recipient.Identity)]
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $AllRecipients[$index].AcceptMessagesOnlyFromSendersOrMembers = $Recipient.AcceptMessagesOnlyFromSendersOrMembers
                                            }
                                        }
                                    } else {
                                        Write-Host '  0 recipients'
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import AcceptMessagesOnlyFrom',
                                                    "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import AcceptMessagesOnlyFrom',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityToIndex       = $AllRecipientsIdentityToIndex
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        Write-Host ('  {0:0000000} recipients with AcceptMessagesOnlyFrom found' -f $(($AllRecipients | Where-Object { $_.AcceptMessagesOnlyFromSendersOrMembers }).count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import ResourceDelegates
    Write-Host
    Write-Host "Import ResourceDelegates, grouped by RecipientTypeDetails @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportResourceDelegates) {
        $Filters = @()

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            if ($AllRecipients[$x].RecipientTypeDetails -iin @('RoomMailbox', 'EquipmentMailbox', 'RemoteRoomMailbox', 'RemoteEquipmentMailbox')) {
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
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import ResourceDelegates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Recipient = $AllRecipients[$RecipientID]

                                Write-Host "Recipient $($RecipientID) ($($Recipient.PrimarySmtpAddress)) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { get-calendarprocessing -identity $($Recipient.Guid.Guid) -ErrorAction stop -WarningAction SilentlyContinue }

                                    $x = $x | Select-Object ResourceDelegates, AllBookInPolicy, BookInPolicy, AllRequestInPolicy, RequestInPolicy, AllRequestOutOfPolicy, RequestOutOfPolicy
                                    if ($x) {
                                        $Recipient.ResourceDelegates = $x.ResourceDelegates
                                        $Recipient.AllBookInPolicy = $x.AllBookInPolicy
                                        $Recipient.BookInPolicy = $x.BookInPolicy
                                        $Recipient.AllRequestInPolicy = $x.AllRequestInPolicy
                                        $Recipient.RequestInPolicy = $x.RequestInPolicy
                                        $Recipient.AllRequestOutOfPolicy = $x.AllRequestOutOfPolicy
                                        $Recipient.RequestOutOfPolicy = $x.RequestOutOfPolicy
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import ResourceDelegates',
                                                    "Recipient $($RecipientID) ($($Recipient.PrimarySmtpAddress))",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import ResourceDelegates',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        Write-Host ('  {0:0000000} recipients with ResourceDelegates found' -f $(($AllRecipients | Where-Object { $_.ResourceDelegates -or $_.AllBookInPolicy -or $_.BookInPolicy -or $_.AllRequestInPolicy -or $_.RequestInPolicy -or $_.AllRequestOutOfPolicy -or $_.RequestOutOfPolicy }).count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import LegacyExchangeDN
    Write-Host
    Write-Host "Import LegacyExchangeDN, grouped by RecipientTypeDetails and first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportResourceDelegates) {
        $Filters = @()

        foreach ($WildcardSearchString in $WildcardSearchStrings) {
            $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('((name -like ''' + $WildcardSearchString + '*'') -and (LegacyExchangeDN -ne `$null))') + ''''
        }

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        foreach ($Cmdlet in @(@(
                    'Get-CASMailbox',
                    'Get-DistributionGroup',
                    'Get-DynamicDistributionGroup',
                    $(if ($ExportFromOnprem -eq $false) {
                            'Get-LinkedUser'
                        }),
                    'Get-Mailbox', # Get-EXOMailbox can't yet handle the filter
                    'Get-MailContact',
                    'Get-MailPublicFolder',
                    'Get-MailUser',
                    $(if ($ExportFromOnprem -eq $true) {
                            'Get-RemoteMailbox'
                        }),
                    'Get-User',
                    $(if ($ExportFromOnprem -eq $false) {
                            'Get-UnifiedGroup'
                        })
                ) | Where-Object { $_ })) {
            foreach ($Filter in $Filters) {
                $tempQueue.enqueue(@($Cmdlet, $Filter))
            }
        }
        $Filters = $null

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
                            $AllRecipientsIdentityToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import LegacyExchangeDN @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $QueueArray = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }


                                Write-Host "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock ([scriptblock]::Create("$($QueueArray[0]) -Filter $($QueueArray[1]) -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue"))

                                    if ($x) {
                                        $x = @($x | Select-Object Identity, LegacyExchangeDN)

                                        Write-Host "  $($x.count) recipients"

                                        foreach ($FoundRecipient in @($x)) {
                                            try {
                                                $index = $null
                                                $index = $AllRecipientsIdentityToIndex[$($FoundRecipient.Identity)]
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $AllRecipients[$index].LegacyExchangeDN = $FoundRecipient.LegacyExchangeDN
                                            }
                                        }
                                    } else {
                                        Write-Host '  0 recipients'
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import LegacyExchangeDN',
                                                    "Cmdlet '$($QueueArray[0])', Filter '$($QueueArray[1])'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import LegacyExchangeDN',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityToIndex       = $AllRecipientsIdentityToIndex
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        Write-Host ('  {0:0000000} recipients with LegacyExchangeDN found' -f $(($AllRecipients | Where-Object { $_.LegacyExchangeDN }).count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Create lookup hashtable for LegacyExchangeDN
    Write-Host
    Write-Host "Create lookup hashtable for LegacyExchangeDN @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportResourceDelegates) {
        Write-Host "  LegacyExchangeDN to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllRecipientsLegacyExchangeDnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]
            if ($Recipient.LegacyExchangeDn) {
                if ($AllRecipientsLegacyExchangeDnToIndex.ContainsKey($($Recipient.LegacyExchangeDn))) {
                    if ($AllRecipientsLegacyExchangeDnToIndex[$($Recipient.LegacyExchangeDn)]) {
                        Write-Host "      '$($Recipient.LegacyExchangeDn)' is not unique" -ForegroundColor Yellow
                    }

                    $AllRecipientsLegacyExchangeDnToIndex[$Recipient.LegacyExchangeDn] = $null
                } else {
                    $AllRecipientsLegacyExchangeDnToIndex[$Recipient.LegacyExchangeDn] = $x
                }
            }
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Define Grantors
    Write-Host
    Write-Host "Define grantors by filtering recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host "  Filter: { $($GrantorFilter) }"
    $GrantorsToConsider = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))


    if (-not $GrantorFilter) {
        $GrantorsToConsider.AddRange(@(0..$($AllRecipients.count - 1)))
    } else {
        for ($x = 0; $x -lt $AllRecipients.Count; $x++) {
            $Grantor = $AllRecipients[$x]

            if ((. ([scriptblock]::Create($GrantorFilter))) -eq $true) {
                $null = $GrantorsToConsider.add($x)
            }
        }
    }

    $GrantorsToConsider.TrimToSize()
    Write-Host ('  {0:0000000}/{1:0000000} recipients are considered as grantors' -f $($GrantorsToConsider.count), $($AllRecipients.count))


    # Import direct group membership
    Write-Host
    Write-Host "Import direct group membership, grouped by first character of name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportManagementRoleGroupMembers -or $ExpandGroups -or ($ExportDistributionGroupMembers -ieq 'All') -or ($ExportDistributionGroupMembers -ieq 'OnlyTrustees')) {
        $AllGroups = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

        $Filters = @()

        foreach ($tempChar in @([char[]](0..255) -clike '[A-Z0-9]')) {
            $Filters += '''' + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent('(name -like ''' + $tempChar + '*'')') + ''''
        }

        $filters += '''' + ($filters.trimstart('''').trimend('''') -join ' -and ').replace('(name -like ''', '(name -notlike ''') + ''''

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        foreach ($Filter in $Filters) {
            $tempQueue.enqueue($Filter)
        }

        $Filters = $null

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
                            $AllGroups,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import direct group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $filter = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Filter '$($filter)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-Group -Filter $($filter) -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue }

                                    if ($x) {
                                        $x = @($x | Select-Object Name, DisplayName, Identity, Guid, Members, RecipientType, RecipientTypeDetails)

                                        $AllGroups.AddRange(@($x | Sort-Object -Property @{expression = { ($_.DisplayName, $_.Name, 'Warning: No valid info found') | Where-Object { $_ } | Select-Object -First 1 } }))
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Import direct group membership',
                                                    "Filter '$($filter)'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Import direct group membership',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllGroups                          = $AllGroups
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all queries have been performed. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }

        $AllGroups.TrimToSize()
        Write-Host ('  {0:0000000} groups with direct members found' -f $($AllGroups.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export ManagedBy permissions
    Write-Host
    Write-Host "Get and export ManagedBy permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportManagedBy -or $ExportModerators) {
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
                            $AllRecipientsIdentityToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export ManagedBy permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $trustees = [system.collections.arraylist]::new(1000)

                                    foreach ($TrusteeRight in $Grantor.ManagedBy) {
                                        $index = $null
                                        $index = $AllRecipientsIdentityToIndex[$TrusteeRight]

                                        if ($index -ge 0) {
                                            $trustees.add($AllRecipients[$index])
                                        } else {
                                            $trustees.add($TrusteeRight)
                                        }
                                    }

                                    foreach ($Trustee in $Trustees) {
                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if ($ExportFromOnPrem) {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                $TrusteeEnvironment = 'Cloud'
                                            } else {
                                                $TrusteeEnvironment = 'On-Prem'
                                            }
                                        } else {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                $TrusteeEnvironment = 'On-Prem'
                                            } else {
                                                $TrusteeEnvironment = 'Cloud'
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            $ExportFileLines.add(
                                                               ('"' + (@((
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                            $GrantorEnvironment,
                                                            '',
                                                            'ManagedBy',
                                                            'Allow',
                                                            'False',
                                                            'None',
                                                            $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                            $Trustee.PrimarySmtpAddress,
                                                            $Trustee.DisplayName,
                                                            $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                            $TrusteeEnvironment
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                            )
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export ManagedBy permissions',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export ManagedBy permissions',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsIdentityToIndex           = $AllRecipientsIdentityToIndex
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all grantors have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Mailbox Access Rights
    Write-Host
    Write-Host "Get and export Mailbox Access Rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportMailboxAccessRights) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.RecipientTypeDetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
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
                            $AllRecipientsDisplaynameToIndex,
                            $AllRecipientsLinkedMasterAccountToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllRecipientsUfnToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsDisplaynameToIndex,
                            $AllSecurityPrincipalsDnToIndex,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $AllSecurityPrincipalsSidToIndex,
                            $AllSecurityPrincipalsUfnToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportMailboxAccessRightsInherited,
                            $ExportMailboxAccessRightsSelf,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $TrusteeFilter,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Mailbox Access Rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()

                                    $runsDone++
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]
                                $Trustee = $null

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "$($Grantor.ExchangeGuid.Guid), $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    foreach ($MailboxPermission in
                                        @($(
                                                if ($ExportFromOnPrem) {
                                                    $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailboxPermission -identity $($Grantor.Guid.Guid) -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue }

                                                    $UFNSelf = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-SecurityPrincipal -Types WellknownSecurityPrincipal -ErrorAction stop -WarningAction SilentlyContinue }
                                                    if ($UFNSelf) {
                                                        $UFNSelf = @($UFNSelf)
                                                        $UFNSelf = ($UFNSelf | Where-Object { $_.Sid -ieq 'S-1-5-10' }).UserFriendlyName
                                                    } else {
                                                        $UFNSelf = $null
                                                    }
                                                } else {
                                                    if ($GrantorRecipientTypeDetails -ine 'GroupMailbox') {
                                                        if ($Grantor.WhenSoftDeleted) {
                                                            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-EXOMailboxPermission -PrimarySmtpAddress $("'$($GrantorPrimarySMTP)'") -SoftDeletedMailbox -ResultSize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
                                                        } else {
                                                            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-EXOMailboxPermission -PrimarySmtpAddress $("'$($GrantorPrimarySMTP)'") -ResultSize unlimited -ErrorAction Stop -WarningAction silentlycontinue }
                                                        }

                                                        $UFNSelf = 'NT AUTHORITY\SELF'
                                                    }
                                                }

                                                if ($x) {
                                                    @($x | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType)
                                                } else {
                                                    @()
                                                }
                                            ))
                                    ) {
                                        foreach ($TrusteeRight in @($MailboxPermission | Where-Object { if ($ExportMailboxAccessRightsInherited) {
                                                        $true
                                                    } else {
                                                        $_.IsInherited -ne $true
                                                    } } | Select-Object *, @{ name = 'trustee'; Expression = { $_.user } })
                                        ) {
                                            if ((-not $ExportMailboxAccessRightsSelf) -and (($TrusteeRight.user -ieq 'S-1-5-10') -or ($TrusteeRight.user -ieq $UFNSelf))) {
                                                continue
                                            }

                                            $trustees = [system.collections.arraylist]::new(1000)

                                            $index = $null

                                            if (($TrusteeRight.user) -and ($TrusteeRight.user -ine 'S-1-5-10') -and ($TrusteeRight.user -ine $UFNSelf)) {
                                                foreach ($LookupTable in @($AllRecipientsUfnToIndex, $AllRecipientsLinkedmasteraccountToIndex, $AllRecipientsSmtpToIndex, $AllRecipientsDisplaynameToIndex)) {
                                                    if ($LookupTable.ContainsKey($($TrusteeRight.trustee))) {
                                                        $index = $LookupTable[$($TrusteeRight.trustee)]
                                                        break
                                                    }
                                                }
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
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                foreach ($Accessright in ($TrusteeRight.Accessrights -split ', ')) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                        $ExportFileLines.add(
                                                                ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        $Accessright,
                                                                        $(if ($Trusteeright.deny) {
                                                                                'Deny'
                                                                            } else {
                                                                                'Allow'
                                                                            }),
                                                                        $Trusteeright.IsInherited,
                                                                        $Trusteeright.InheritanceType,
                                                                        $TrusteeRight.trustee,
                                                                        $Trustee.PrimarySmtpAddress,
                                                                        $Trustee.DisplayName,
                                                                        $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) {
                                                                                $SecurityPrincipalsLookupSearchString = $(
                                                                                    if ($trustee.Guid.Guid) {
                                                                                        $trustee.Guid.Guid
                                                                                    } else {
                                                                                        $AllSecurityPrincipalsLookupSearchString = "$($TrusteeRight.User)"

                                                                                        $AllSecurityPrincipalsLookupResult = (
                                                                                            $AllSecurityPrincipalsDnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                            $AllSecurityPrincipalsObjectguidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                            $AllSecurityPrincipalsSidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                            $AllSecurityPrincipalsUfnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                            $AllSecurityPrincipalsDisplaynameToIndex[$AllSecurityPrincipalsLookupSearchString]
                                                                                        ) | Where-Object { $_ } | Select-Object -First 1

                                                                                        if ($AllSecurityPrincipalsLookupResult) {
                                                                                            if ($AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Sid.tostring().StartsWith('S-1-5-21-', 'CurrentCultureIgnoreCase')) {
                                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Guid.Guid
                                                                                            } else {
                                                                                                ''
                                                                                            }
                                                                                        } else {
                                                                                            try {
                                                                                                if ($ExportFromOnPrem) {
                                                                                                    # could be an object from a trust
                                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                                    $objNT = $objTrans.GetType()
                                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AllSecurityPrincipalsLookupSearchString)"))
                                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                                } else {
                                                                                                    ''
                                                                                                }
                                                                                            } catch {
                                                                                                ''
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                ); $SecurityPrincipalsLookupSearchString
                                                                            } else { '' }
                                                                        ),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                        $TrusteeEnvironment

                                                                    ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                        )

                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export Mailbox Access Rights',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Mailbox Access Rights',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                           = $AllRecipients
                        AllRecipientsDisplaynameToIndex         = $AllRecipientsDisplaynameToIndex
                        AllRecipientsLinkedMasterAccountToIndex = $AllRecipientsLinkedMasterAccountToIndex
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
                        AllRecipientsUfnToIndex                 = $AllRecipientsUfnToIndex
                        AllSecurityPrincipals                   = $AllSecurityPrincipals
                        AllSecurityPrincipalsDisplaynameToIndex = $AllSecurityPrincipalsDisplaynameToIndex
                        AllSecurityPrincipalsDnToIndex          = $AllSecurityPrincipalsDnToIndex
                        AllSecurityPrincipalsObjectguidToIndex  = $AllSecurityPrincipalsObjectguidToIndex
                        AllSecurityPrincipalsSidToIndex         = $AllSecurityPrincipalsSidToIndex
                        AllSecurityPrincipalsUfnToIndex         = $AllSecurityPrincipalsUfnToIndex
                        ConnectExchange                         = $ConnectExchange
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                         = $DebugPreference
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                      = $ExchangeCredential
                        ExchangeOnlineConnectionParameters      = $ExchangeOnlineConnectionParameters
                        ExportFile                              = $ExportFile
                        ExportFileFilter                        = $ExportFileFilter
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileHeaderIndexes                 = $ExportFileHeaderIndexes
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExportGuids                             = $ExportGuids
                        ExportSids                              = $exportSids
                        ExportMailboxAccessRightsInherited      = $ExportMailboxAccessRightsInherited
                        ExportMailboxAccessRightsSelf           = $ExportMailboxAccessRightsSelf
                        ExportTrustees                          = $ExportTrustees
                        ScriptPath                              = $PSScriptRoot
                        tempConnectionUriQueue                  = $tempConnectionUriQueue
                        tempQueue                               = $tempQueue
                        TrusteeFilter                           = $TrusteeFilter
                        UseDefaultCredential                    = $UseDefaultCredential
                        UTF8Encoding                            = $UTF8Encoding
                        VerbosePreference                       = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantor mailboxes to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all grantor mailboxes have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Mailbox Folder permissions
    Write-Host
    Write-Host "Get and export Mailbox Folder Permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportMailboxFolderPermissions) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.RecipientTypeDetails -ilike '*Mailbox') -and ($x -in $GrantorsToConsider) -and ($Recipient.RecipientTypeDetails -inotin @('PublicFolderMailbox', 'MonitoringMailbox')) -and (-not $Recipient.WhenSoftDeleted)) {
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
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportMailboxFolderPermissions,
                            $ExportMailboxFolderPermissionsAnonymous,
                            $ExportMailboxFolderPermissionsDefault,
                            $ExportMailboxFolderPermissionsExcludeFoldertype,
                            $ExportMailboxFolderPermissionsMemberAtLocal,
                            $ExportMailboxFolderPermissionsOwnerAtLocal,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $TrusteeFilter,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Mailbox Folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }
                                'xxx'

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                if ($ExportFromOnPrem) {
                                    $Folders = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailboxFolderStatistics -identity $($Grantor.Guid.Guid) -ErrorAction Stop -WarningAction silentlycontinue }
                                } else {
                                    $Folders = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-EXOMailboxFolderStatistics -PrimarySmtpAddress $("'$($GrantorPrimarySMTP)'") -ErrorAction Stop -WarningAction silentlycontinue }
                                }

                                if ($Folders) {
                                    $Folders = @($Folders | Select-Object folderid, folderpath, foldertype)
                                } else {
                                    $Folders = @()
                                }

                                foreach ($Folder in $Folders) {
                                    try {
                                        if (-not $folder.foldertype) {
                                            $folder.foldertype = $null
                                        }

                                        if ($folder.foldertype -iin $ExportMailboxFolderPermissionsExcludeFoldertype) {
                                            continue
                                        }

                                        if ($Folder.foldertype -ieq 'root') {
                                            $Folder.folderpath = '/'
                                        }

                                        Write-Host "  Folder '$($folder.folderid)' ('$($folder.folderpath)') @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                                        foreach ($FolderPermissions in
                                            @($(
                                                    if ($ExportFromOnPrem) {
                                                        $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailboxFolderPermission -identity $('' + $($Grantor.Guid.Guid) + ':' + $($Folder.folderid)) -ErrorAction stop -WarningAction silentlycontinue }
                                                    } else {
                                                        if ($GrantorRecipientTypeDetails -ieq 'groupmailbox') {
                                                            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailboxFolderPermission -Identity $("'$($GrantorPrimarySMTP):$($Folder.folderid)'") -GroupMailbox -ErrorAction stop -WarningAction silentlycontinue }
                                                        } else {
                                                            $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-MailboxFolderPermission -Identity $("'$($GrantorPrimarySMTP):$($Folder.folderid)'") -ErrorAction stop -WarningAction silentlycontinue }
                                                        }
                                                    }

                                                    if ($x) {
                                                        @($x | Select-Object user, accessrights)
                                                    } else {
                                                        @()
                                                    }
                                                ))
                                        ) {
                                            foreach ($FolderPermission in $FolderPermissions) {
                                                foreach ($AccessRight in ($FolderPermission.AccessRights)) {
                                                    if ($ExportMailboxFolderPermissionsDefault -eq $false) {
                                                        if ($FolderPermission.user.usertype.value -ieq 'default') {
                                                            continue
                                                        }
                                                    }

                                                    if ($ExportMailboxFolderPermissionsAnonymous -eq $false) {
                                                        if ($FolderPermission.user.usertype.value -ieq 'anonymous') {
                                                            continue
                                                        }
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

                                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                                $TrusteeEnvironment = 'Cloud'
                                                            } else {
                                                                $TrusteeEnvironment = 'On-Prem'
                                                            }

                                                            $ExportFileLines.Add(
                                                                    ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                            $(if ($ExportSids) {
                                                                                    try {
                                                                                        try {
                                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                        } catch {
                                                                                            if ($ExportFromOnPrem) {
                                                                                                # could be an object from a trust
                                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                                $objNT = $objTrans.GetType()
                                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                            } else {
                                                                                                ''
                                                                                            }
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                } else { '' }
                                                                            ),
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            $($Folder.Folderpath),
                                                                            $($Accessright),
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $($FolderPermission.user.displayname),
                                                                            $($Trustee.PrimarySmtpAddress),
                                                                            $($Trustee.displayname),
                                                                            $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                            $(if ($ExportGuids) { $(($Trustee.Guid.Guid, $FolderPermission.User.AdRecipient.Guid.Guid, '') | Select-Object -First 1) } else { '' }),
                                                                            $(if ($ExportSids) {
                                                                                    try {
                                                                                        try {
                                                                                            $SecurityPrincipalsLookupSearchString = $(($Trustee.Guid.Guid, $FolderPermission.User.AdRecipient.Guid.Guid, '') | Select-Object -First 1)
                                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                        } catch {
                                                                                            if ($ExportFromOnPrem) {
                                                                                                # could be an object from a trust
                                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                                $objNT = $objTrans.GetType()
                                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                            } else {
                                                                                                ''
                                                                                            }
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                } else { '' }
                                                                            ),
                                                                            $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                            )

                                                        }
                                                    } else {
                                                        if ($ExportMailboxFolderPermissionsOwnerAtLocal -eq $false) {
                                                            if ($FolderPermission.user.recipientprincipal.primarysmtpaddress -ieq 'owner@local') {
                                                                continue
                                                            }
                                                        }

                                                        if ($ExportMailboxFolderPermissionsMemberAtLocal -eq $false) {
                                                            if ($FolderPermission.user.recipientprincipal.primarysmtpaddress -ieq 'member@local') {
                                                                continue
                                                            }
                                                        }

                                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $FolderPermission.user.recipientprincipal)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($FolderPermission.user.recipientprincipal))) {
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

                                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                                $TrusteeEnvironment = 'On-Prem'
                                                            } else {
                                                                $TrusteeEnvironment = 'Cloud'
                                                            }

                                                            $ExportFileLines.Add(
                                                                    ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                            $(if ($ExportSids) {
                                                                                    try {
                                                                                        try {
                                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                        } catch {
                                                                                            if ($ExportFromOnPrem) {
                                                                                                # could be an object from a trust
                                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                                $objNT = $objTrans.GetType()
                                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                            } else {
                                                                                                ''
                                                                                            }
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                } else { '' }
                                                                            ),
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            $($Folder.Folderpath),
                                                                            $($Accessright),
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $($FolderPermission.user.displayname),
                                                                            $($Trustee.PrimarySmtpAddress),
                                                                            $($Trustee.displayname),
                                                                            $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                            $(if ($ExportGuids) { $(($Trustee.Guid.Guid, $FolderPermission.User.RecipientPrincipcal.Guid.Guid, '') | Select-Object -First 1) } else { '' }),
                                                                            $(if ($ExportSids) {
                                                                                    try {
                                                                                        try {
                                                                                            $SecurityPrincipalsLookupSearchString = $(($Trustee.Guid.Guid, $FolderPermission.User.RecipientPrincipcal.Guid.Guid, '') | Select-Object -First 1)
                                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                        } catch {
                                                                                            if ($ExportFromOnPrem) {
                                                                                                # could be an object from a trust
                                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                                $objNT = $objTrans.GetType()
                                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                            } else {
                                                                                                ''
                                                                                            }
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                } else { '' }
                                                                            ), $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                            )

                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Get and export Mailbox Folder permissions',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails), $($GrantorPrimarySMTP):$($Folder.folderid) ($($Folder.folderpath))",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Mailbox Folder permissions',
                                            "($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                                   = $AllRecipients
                        AllRecipientsSmtpToIndex                        = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                           = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex          = $AllSecurityPrincipalsObjectguidToIndex
                        ConnectExchange                                 = $ConnectExchange
                        DebugFile                                       = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                                 = $DebugPreference
                        ErrorFile                                       = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                              = $ExchangeCredential
                        ExchangeOnlineConnectionParameters              = $ExchangeOnlineConnectionParameters
                        ExportFile                                      = $ExportFile
                        ExportFileFilter                                = $ExportFileFilter
                        ExportFileHeader                                = $ExportFileHeader
                        ExportFileHeaderIndexes                         = $ExportFileHeaderIndexes
                        ExportFromOnPrem                                = $ExportFromOnPrem
                        ExportGuids                                     = $ExportGuids
                        ExportMailboxFolderPermissions                  = $ExportMailboxFolderPermissions
                        ExportMailboxFolderPermissionsAnonymous         = $ExportMailboxFolderPermissionsAnonymous
                        ExportMailboxFolderPermissionsDefault           = $ExportMailboxFolderPermissionsDefault
                        ExportMailboxFolderPermissionsExcludeFoldertype = $ExportMailboxFolderPermissionsExcludeFoldertype
                        ExportMailboxFolderPermissionsMemberAtLocal     = $ExportMailboxFolderPermissionsMemberAtLocal
                        ExportMailboxFolderPermissionsOwnerAtLocal      = $ExportMailboxFolderPermissionsOwnerAtLocal
                        ExportSids                                      = $ExportSids
                        ExportTrustees                                  = $ExportTrustees
                        ScriptPath                                      = $PSScriptRoot
                        tempConnectionUriQueue                          = $tempConnectionUriQueue
                        tempQueue                                       = $tempQueue
                        TrusteeFilter                                   = $TrusteeFilter
                        UseDefaultCredential                            = $UseDefaultCredential
                        UTF8Encoding                                    = $UTF8Encoding
                        VerbosePreference                               = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantor mailboxes to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all grantor mailboxes have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Send As permissions
    Write-Host
    Write-Host "Get and export Send As permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
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
                            $AllRecipientsDisplaynameToIndex,
                            $AllRecipientsLinkedMasterAccountToIndex,
                            $AllRecipientsSendas,
                            $AllRecipientsSmtpToIndex,
                            $AllRecipientsUfnToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsDisplaynameToIndex,
                            $AllSecurityPrincipalsDnToIndex,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $AllSecurityPrincipalsSidToIndex,
                            $AllSecurityPrincipalsUfnToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportSendAsSelf,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Send As permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    if ($ExportFromOnPrem) {
                                        try {
                                            $entries = @(([adsi]"LDAP://<GUID=$($Grantor.Guid.Guid)>").ObjectSecurity.Access)

                                            if (-not $entries) {
                                                throw 'retry'
                                            }
                                        } catch {
                                            Start-Sleep -Seconds 2

                                            $entries = @(([adsi]"LDAP://<GUID=$($Grantor.Guid.Guid)>").ObjectSecurity.Access)
                                        }

                                        foreach ($entry in $entries) {
                                            $trustee = $null

                                            if ($entry.ObjectType -eq 'ab721a54-1e2f-11d0-9819-00aa0040529b') {
                                                if (($ExportSendAsSelf -eq $false) -and ($entry.identityreference.value -ilike '*\*') -and ((([System.Security.Principal.NTAccount]::new($entry.identityreference.value)).Translate([System.Security.Principal.SecurityIdentifier])).value -ieq 'S-1-5-10')) {
                                                    continue
                                                } else {
                                                    try {
                                                        $index = $null
                                                        $index = ($AllRecipientsUfnToIndex[$($entry.identityreference.value)], $AllRecipientsLinkedmasteraccountToIndex[$($entry.identityreference.value)], '') | Select-Object -First 1
                                                    } catch {
                                                    }
                                                }

                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    $trustee = $entry.identityreference.value
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                    $ExportFileLines.add(
                                                            ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'SendAs',
                                                                    $entry.AccessControlType,
                                                                    $entry.IsInherited,
                                                                    $entry.InheritanceType,
                                                                    $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress,
                                                                    $Trustee.DisplayName,
                                                                    $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) {
                                                                            $SecurityPrincipalsLookupSearchString = $(
                                                                                if ($trustee.Guid.Guid) {
                                                                                    $trustee.Guid.Guid
                                                                                } else {
                                                                                    $AllSecurityPrincipalsLookupSearchString = "$($entry.identityreference.value)"

                                                                                    $AllSecurityPrincipalsLookupResult = (
                                                                                        $AllSecurityPrincipalsDnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                        $AllSecurityPrincipalsObjectguidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                        $AllSecurityPrincipalsSidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                        $AllSecurityPrincipalsUfnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                        $AllSecurityPrincipalsDisplaynameToIndex[$AllSecurityPrincipalsLookupSearchString]
                                                                                    ) | Where-Object { $_ } | Select-Object -First 1

                                                                                    if ($AllSecurityPrincipalsLookupResult) {
                                                                                        if ($AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Sid.tostring().StartsWith('S-1-5-21-', 'CurrentCultureIgnoreCase')) {
                                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Guid.Guid
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    } else {
                                                                                        try {
                                                                                            if ($ExportFromOnPrem) {
                                                                                                # could be an object from a trust
                                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                                $objNT = $objTrans.GetType()
                                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AllSecurityPrincipalsLookupSearchString)"))
                                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                            } else {
                                                                                                ''
                                                                                            }
                                                                                        } catch {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                }
                                                                            ); $SecurityPrincipalsLookupSearchString
                                                                        } else { '' }
                                                                    ),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                    )


                                                }
                                            }
                                        }
                                    } else {
                                        foreach ($entry in $AllRecipientsSendas) {
                                            if ($entry.Identity -eq $Grantor.Identity) {
                                                if (($ExportSendAsSelf -eq $false) -and ($entry.trustee -ieq 'NT AUTHORITY\SELF')) {
                                                    continue
                                                }

                                                $trustee = $null

                                                $index = $null

                                                if (($entry.trusteesidstring) -and ($entry.trusteesidstring -ine 'S-1-5-10')) {
                                                    try {
                                                        $index = $AllRecipientsUfnToIndex[$($($AllSecurityPrincipals[$($AllSecurityPrincipalsSidToIndex[$($entry.TrusteeSidString)])]).UserFriendlyName)]
                                                    } catch {
                                                    }
                                                }

                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    try {
                                                        $index = $AllRecipientsSmtpToIndex[$entry.trustee]

                                                        if ($index -ge 0) {
                                                            $trustee = $AllRecipients[$index]
                                                        } else {
                                                            $trustee = $entry.trustee
                                                        }
                                                    } catch {
                                                        $trustee = $entry.trustee
                                                    }
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                foreach ($AccessRight in $entry.AccessRights) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                        $ExportFileLines.add(
                                                                ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        $AccessRight,
                                                                        $entry.AccessControlType,
                                                                        $entry.IsInherited,
                                                                        $entry.InheritanceType,
                                                                        $(($Trustee.displayname, $entry.trustee, '') | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress,
                                                                        $Trustee.DisplayName,
                                                                        $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                        )

                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export Send As permissions',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Send As permissions',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
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
                        AllRecipientsDisplaynameToIndex         = $AllRecipientsDisplaynameToIndex
                        AllRecipientsLinkedMasterAccountToIndex = $AllRecipientsLinkedMasterAccountToIndex
                        AllRecipientsSendas                     = $AllRecipientsSendas
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
                        AllRecipientsUfnToIndex                 = $AllRecipientsUfnToIndex
                        AllSecurityPrincipals                   = $AllSecurityPrincipals
                        AllSecurityPrincipalsDisplaynameToIndex = $AllSecurityPrincipalsDisplaynameToIndex
                        AllSecurityPrincipalsDnToIndex          = $AllSecurityPrincipalsDnToIndex
                        AllSecurityPrincipalsObjectguidToIndex  = $AllSecurityPrincipalsObjectguidToIndex
                        AllSecurityPrincipalsSidToIndex         = $AllSecurityPrincipalsSidToIndex
                        AllSecurityPrincipalsUfnToIndex         = $AllSecurityPrincipalsUfnToIndex
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                         = $DebugPreference
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                              = $ExportFile
                        ExportFileFilter                        = $ExportFileFilter
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileHeaderIndexes                 = $ExportFileHeaderIndexes
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExportGuids                             = $ExportGuids
                        ExportSids                              = $ExportSids
                        ExportSendAsSelf                        = $ExportSendAsSelf
                        ExportTrustees                          = $ExportTrustees
                        ScriptPath                              = $PSScriptRoot
                        tempQueue                               = $tempQueue
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                        VerbosePreference                       = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all grantors have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Send On Behalf permissions
    Write-Host
    Write-Host "Get and export Send On Behalf permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
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
                            $AllRecipientsDnToIndex,
                            $AllRecipientsSendonbehalf,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Send On Behalf permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    if ($ExportFromOnPrem) {
                                        try {
                                            $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher("(objectguid=$([System.String]::Join('', (([guid]$($Grantor.Guid.Guid)).ToByteArray() | ForEach-Object { '\' + $_.ToString('x2') })).ToUpper()))")
                                            $directorySearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$(($Grantor.identity -split '/')[0])")
                                            $null = $directorySearcher.PropertiesToLoad.Add('publicDelegates')

                                            if ($ExportGuids) {
                                                $null = $directorySearcher.PropertiesToLoad.Add('objectGuid')
                                            }

                                            $directorySearcherResults = $directorySearcher.FindOne()

                                            if (-not $directorySearcherResults) {
                                                throw 'retry'
                                            }
                                        } catch {
                                            Start-Sleep -Seconds 2

                                            $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher("(objectguid=$([System.String]::Join('', (([guid]$($Grantor.Guid.Guid)).ToByteArray() | ForEach-Object { '\' + $_.ToString('x2') })).ToUpper()))")
                                            $directorySearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$(($Grantor.identity -split '/')[0])")
                                            $null = $directorySearcher.PropertiesToLoad.Add('publicDelegates')

                                            if ($ExportGuids) {
                                                $null = $directorySearcher.PropertiesToLoad.Add('objectGuid')
                                            }

                                            $directorySearcherResults = $directorySearcher.FindOne()
                                        }


                                        foreach ($directorySearcherResult in $directorySearcherResults) {
                                            foreach ($delegateBindDN in $directorySearcherResult.properties.publicdelegates) {
                                                $index = $null
                                                $index = $AllRecipientsDnToIndex[$delegateBindDN]

                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    $trustee = $delegateBindDn
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                    $ExportFileLines.add(
                                                            ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'SendOnBehalf',
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress,
                                                                    $Trustee.DisplayName,
                                                                    $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) {
                                                                            $SecurityPrincipalsLookupSearchString = $(
                                                                                if ($Trustee.Guid.Guid) {
                                                                                    $Trustee.Guid.Guid
                                                                                } else {
                                                                                    try {
                                                                                        [guid]::new($directorySearcherResult.properties.objectguid[0]).Guid
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            ); $SecurityPrincipalsLookupSearchString
                                                                        } else { '' }
                                                                    ),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                    )

                                                }
                                            }
                                        }
                                    } else {
                                        foreach ($entry in $AllRecipientsSendonbehalf) {
                                            if ($entry.Guid.Guid -eq $Grantor.Guid.Guid) {
                                                $trustee = $null
                                                foreach ($AccessRight in $entry.GrantSendOnBehalfTo) {
                                                    $index = $null
                                                    $index = $AllRecipientsIdentityToIndex[$AccessRight]

                                                    if ($index -ge 0) {
                                                        $trustee = $AllRecipients[$index]
                                                    } else {
                                                        $trustee = $AccessRight
                                                    }

                                                    if ($TrusteeFilter) {
                                                        if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                            continue
                                                        }
                                                    }

                                                    if ($ExportFromOnPrem) {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                            $TrusteeEnvironment = 'Cloud'
                                                        } else {
                                                            $TrusteeEnvironment = 'On-Prem'
                                                        }
                                                    } else {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                            $TrusteeEnvironment = 'On-Prem'
                                                        } else {
                                                            $TrusteeEnvironment = 'Cloud'
                                                        }
                                                    }

                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                        $ExportFileLines.add(
                                                                ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendOnBehalf',
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $(($Trustee.displayname, $Truste, '') | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress,
                                                                        $Trustee.DisplayName,
                                                                        $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                        )

                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export Send On Behalf permissions',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Send On Behalf permissions',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsDnToIndex                 = $AllRecipientsDnToIndex
                        AllRecipientsIdentityToIndex           = $AllRecipientsIdentityToIndex
                        AllRecipientsSendonbehalf              = $AllRecipientsSendonbehalf
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all grantors have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Linked Master Accounts
    Write-Host
    Write-Host "Get and export Linked Master Accounts @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportLinkedMasterAccount -and $ExportFromOnPrem) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.RecipientTypeDetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
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
                            $AllRecipientsLinkedmasteraccountToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsDisplaynameToIndex,
                            $AllSecurityPrincipalsDnToIndex,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $AllSecurityPrincipalsSidToIndex,
                            $AllSecurityPrincipalsUfnToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportLinkedMasterAccount,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Linked Master Accounts @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

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

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'Cloud'
                                                } else {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                } else {
                                                    $TrusteeEnvironment = 'Cloud'
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                $ExportFileLines.add(
                                                        ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                '',
                                                                'LinkedMasterAccount',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $Grantor.LinkedMasterAccount,
                                                                $Trustee.PrimarySmtpAddress,
                                                                $Trustee.DisplayName,
                                                                $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) {
                                                                        $SecurityPrincipalsLookupSearchString = $(
                                                                            if ($Trustee.Guid.Guid) {
                                                                                $Trustee.Guid.Guid
                                                                            } else {
                                                                                $AllSecurityPrincipalsLookupSearchString = "$($Trustee)"

                                                                                $AllSecurityPrincipalsLookupResult = (
                                                                                    $AllSecurityPrincipalsDnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsObjectguidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsSidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsUfnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsDisplaynameToIndex[$AllSecurityPrincipalsLookupSearchString]
                                                                                ) | Where-Object { $_ } | Select-Object -First 1

                                                                                if ($AllSecurityPrincipalsLookupResult) {
                                                                                    if ($AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Sid.tostring().StartsWith('S-1-5-21-', 'CurrentCultureIgnoreCase')) {
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Guid.Guid
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                } else {
                                                                                    try {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AllSecurityPrincipalsLookupSearchString)"))
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            }
                                                                        ); $SecurityPrincipalsLookupSearchString
                                                                    } else { '' }
                                                                ),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                )

                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export Linked Master Accounts',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Linked Master Accounts',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
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
                        AllRecipientsLinkedmasteraccountToIndex = $AllRecipientsLinkedmasteraccountToIndex
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                   = $AllSecurityPrincipals
                        AllSecurityPrincipalsDisplaynameToIndex = $AllSecurityPrincipalsDisplaynameToIndex
                        AllSecurityPrincipalsDnToIndex          = $AllSecurityPrincipalsDnToIndex
                        AllSecurityPrincipalsObjectguidToIndex  = $AllSecurityPrincipalsObjectguidToIndex
                        AllSecurityPrincipalsSidToIndex         = $AllSecurityPrincipalsSidToIndex
                        AllSecurityPrincipalsUfnToIndex         = $AllSecurityPrincipalsUfnToIndex
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                         = $DebugPreference
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                              = $ExportFile
                        ExportFileFilter                        = $ExportFileFilter
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileHeaderIndexes                 = $ExportFileHeaderIndexes
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExportGuids                             = $ExportGuids
                        ExportSids                              = $ExportSids
                        ExportLinkedMasterAccount               = $ExportLinkedMasterAccount
                        ExportTrustees                          = $ExportTrustees
                        ScriptPath                              = $PSScriptRoot
                        tempQueue                               = $tempQueue
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                        VerbosePreference                       = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} grantors to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all grantors have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Public Folder permissions
    Write-Host
    Write-Host "Get and export Public Folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportPublicFolderPermissions) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllPublicFolders.count))

        for ($x = 0; $x -lt $AllPublicFolders.count; $x++) {
            $folder = $AllPublicFolders[$x]

            try {
                $index = $null
                $index = $AllRecipientsExchangeGuidToIndex[$($folder.ContentMailboxGuid.Guid)]
            } catch {
            }

            if ($index -ge 0) {
                $Grantor = $AllRecipients[$index]

                if ($GrantorFilter) {
                    if ((. ([scriptblock]::Create($GrantorFilter))) -ne $true) {
                        continue
                    }
                }

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
                            $AllPublicFolders,
                            $AllRecipients,
                            $AllRecipientsExchangeGuidToIndex,
                            $AllRecipientsIdentityGuidToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportPublicFolderPermissions,
                            $ExportPublicFolderPermissionsAnonymous,
                            $ExportPublicFolderPermissionsDefault,
                            $ExportPublicFolderPermissionsExcludeFoldertype,
                            $ExportSids,
                            $ExportTrustees,
                            $GrantorFilter,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $TrusteeFilter,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Public folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $folder.folderpath = '/' + $($folder.folderpath -join '/')

                                    Write-Host "  Folder '$($folder.EntryId)' ('$($Folder.Folderpath)') @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                    if (-not $folder.folderclass) {
                                        $folder.folderclass = $null
                                    }

                                    if ($folder.folderclass -iin $ExportPublicFolderPermissionsExcludeFoldertype) {
                                        continue
                                    }

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
                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'Cloud'
                                                } else {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                } else {
                                                    $TrusteeEnvironment = 'Cloud'
                                                }
                                            }

                                            $ExportFileLines.Add(
                                                    ('"' + (@((
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                            $GrantorEnvironment,
                                                            $($Folder.Folderpath),
                                                            'MailEnabled',
                                                            'Allow',
                                                            'False',
                                                            'None',
                                                            $(($Trustee.PrimarySmtpAddress, $Trustee, '') | Select-Object -First 1),
                                                            $($Trustee.PrimarySmtpAddress),
                                                            $($Trustee.displayname),
                                                            $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                            $TrusteeEnvironment
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                            )

                                        }
                                    }

                                    foreach ($FolderPermissions in
                                        @($(
                                                $x = . ([scriptblock]::Create($ConnectExchange)) -ScriptBlock { Get-PublicFolderClientPermission -identity $($Folder.EntryId) -ErrorAction stop -WarningAction silentlycontinue }

                                                if ($x) {
                                                    @($x | Select-Object user, accessrights )
                                                } else {
                                                    @()
                                                }
                                            ))
                                    ) {
                                        foreach ($FolderPermission in $FolderPermissions) {
                                            foreach ($AccessRight in ($FolderPermission.AccessRights)) {
                                                if ($ExportPublicFolderPermissionsDefault -eq $false) {
                                                    if ($FolderPermission.user.usertype.value -ieq 'default') {
                                                        continue
                                                    }
                                                }

                                                if ($ExportPublicFolderPermissionsAnonymous -eq $false) {
                                                    if ($FolderPermission.user.usertype.value -ieq 'anonymous') {
                                                        continue
                                                    }
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

                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                            $TrusteeEnvironment = 'Cloud'
                                                        } else {
                                                            $TrusteeEnvironment = 'On-Prem'
                                                        }

                                                        $ExportFileLines.Add(
                                                                ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        $($Folder.Folderpath),
                                                                        $($Accessright),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $($FolderPermission.user.displayname),
                                                                        $($Trustee.PrimarySmtpAddress),
                                                                        $($Trustee.displayname),
                                                                        $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $(($Trustee.Guid.Guid, $FolderPermission.User.AdRecipient.Guid.Guid, '') | Select-Object -First 1) } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $(($Trustee.Guid.Guid, $FolderPermission.User.AdRecipient.Guid.Guid, '') | Select-Object -First 1)
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                        )

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

                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                            $TrusteeEnvironment = 'On-Prem'
                                                        } else {
                                                            $TrusteeEnvironment = 'Cloud'
                                                        }

                                                        $ExportFileLines.Add(
                                                                ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        $($Folder.Folderpath),
                                                                        $($Accessright),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $($FolderPermission.user.displayname),
                                                                        $($Trustee.PrimarySmtpAddress),
                                                                        $($Trustee.displayname),
                                                                        $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                        $(if ($ExportGuids) { $(($Trustee.Guid.Guid, $FolderPermission.User.RecipientPrincipal.Guid.Guid, '') | Select-Object -First 1) } else { '' }),
                                                                        $(if ($ExportSids) {
                                                                                try {
                                                                                    try {
                                                                                        $SecurityPrincipalsLookupSearchString = $(($Trustee.Guid.Guid, $FolderPermission.User.RecipientPrincipal.Guid.Guid, '') | Select-Object -First 1)
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                    } catch {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    }
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            } else { '' }
                                                                        ),
                                                                        $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                        )

                                                    }
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export Public Folder permissions',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails), $($GrantorPrimarySMTP):$($Folder.Entryd) ($($Folder.folderpath))",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.PF{1:0000000}.txt' -f $RecipientId, $PublicFolderId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Public Folder permissions',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllPublicFolders                               = $AllPublicFolders
                        AllRecipients                                  = $AllRecipients
                        AllRecipientsExchangeGuidToIndex               = $AllRecipientsExchangeGuidToIndex
                        AllRecipientsIdentityGuidToIndex               = $AllRecipientsIdentityGuidToIndex
                        AllRecipientsSmtpToIndex                       = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                          = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex         = $AllSecurityPrincipalsObjectguidToIndex
                        ConnectExchange                                = $ConnectExchange
                        DebugFile                                      = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                                = $DebugPreference
                        ErrorFile                                      = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                             = $ExchangeCredential
                        ExchangeOnlineConnectionParameters             = $ExchangeOnlineConnectionParameters
                        ExportFile                                     = $ExportFile
                        ExportFileFilter                               = $ExportFileFilter
                        ExportFileHeader                               = $ExportFileHeader
                        ExportFileHeaderIndexes                        = $ExportFileHeaderIndexes
                        ExportFromOnPrem                               = $ExportFromOnPrem
                        ExportGuids                                    = $ExportGuids
                        ExportPublicFolderPermissions                  = $ExportPublicFolderPermissions
                        ExportPublicFolderPermissionsAnonymous         = $ExportPublicFolderPermissionsAnonymous
                        ExportPublicFolderPermissionsDefault           = $ExportPublicFolderPermissionsDefault
                        ExportPublicFolderPermissionsExcludeFoldertype = $ExportPublicFolderPermissionsExcludeFoldertype
                        ExportSids                                     = $ExportSids
                        ExportTrustees                                 = $ExportTrustees
                        GrantorFilter                                  = $GrantorFilter
                        ScriptPath                                     = $PSScriptRoot
                        tempConnectionUriQueue                         = $tempConnectionUriQueue
                        tempQueue                                      = $tempQueue
                        TrusteeFilter                                  = $TrusteeFilter
                        UseDefaultCredential                           = $UseDefaultCredential
                        UTF8Encoding                                   = $UTF8Encoding
                        VerbosePreference                              = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} Public Folders to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all Public Folders have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    #Import-Csv $JobErrorFile -Encoding $UTF8Encoding -Delimiter ';' | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv $ErrorFile -Encoding $UTF8Encoding -Force -Append -NoTypeInformation -Delimiter ';'
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding -Force | Select-Object -Skip 1 | Sort-Object -Unique | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force

                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            if ($ResultFile) {
                foreach ($JobResultFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ResultFile), ('TEMP.*.PF*.txt'))))) {
                    Get-Content -LiteralPath $JobResultFile -Encoding $UTF8Encoding | Select-Object * -Skip 1 | Out-File -LiteralPath ($JobResultFile.fullname -ireplace '\.PF\d{7}.txt$', '.txt') -Append -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobResultFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Forwarders
    Write-Host
    Write-Host "Get and export Forwarders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportForwarders) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider) -and ($AllRecipients[$x].ExternalEmailAddress -or $AllRecipients[$x].ForwardingAddress -or $AllRecipients[$x].ForwardingSmtpAddress)) {
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
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Forwarders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                foreach ($ForwarderType in ('ExternalEmailAddress', 'ForwardingAddress', 'ForwardingSmtpAddress')) {
                                    try {
                                        if ($Grantor.$ForwarderType) {
                                            try {
                                                $index = $null
                                                $index = $AllRecipientsSmtpToIndex[$($Grantor.$ForwarderType -ireplace '^smtp:', '')]
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $Trustee = $AllRecipients[$index]
                                            } else {
                                                $Trustee = $Grantor.$ForwarderType
                                            }

                                            if ($TrusteeFilter) {
                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                    continue
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                $ExportFileLines.add(
                                                            ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                '',
                                                                $('Forward_' + $ForwarderType + $(if ((-not $Grantor.DeliverToMailboxAndForward) -or ($ForwarderType -ieq 'ExternalEmailAddress')) {
                                                                            '_ForwardOnly'
                                                                        } else {
                                                                            '_DeliverAndForward'
                                                                        } )),
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $($Grantor.$ForwarderType),
                                                                $Trustee.PrimarySmtpAddress,
                                                                $Trustee.DisplayName,
                                                                $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                )

                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Get and export Forwarders',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Forwarders',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all recipients have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export moderators
    Write-Host
    Write-Host "Get and export moderators @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportModerators) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider) -and ($null -ne $AllRecipients[$x].ModeratedBy)) {
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
                            $AllRecipientsIdentityToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export moderators @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                foreach ($ModeratorSetting in @('ModeratedBy', 'ModeratedByBypass')) {
                                    foreach ($Moderator in $($Grantor.$ModeratorSetting)) {
                                        try {
                                            try {
                                                $index = $null
                                                $index = $AllRecipientsIdentityToIndex[$($Moderator)]
                                            } catch {
                                            }

                                            if ($index -ge 0) {
                                                $Trustee = $AllRecipients[$index]
                                            } else {
                                                $Trustee = $Moderator
                                            }

                                            if ($TrusteeFilter) {
                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                    continue
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                    $ExportFileLines.add(
                                                            ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $ModeratorSetting,
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $Moderator,
                                                                    $Trustee.PrimarySmtpAddress,
                                                                    $Trustee.DisplayName,
                                                                    $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                    )

                                                }
                                            }
                                        } catch {
                                            (
                                                '"' + (
                                                    @(
                                                        (
                                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                            'Get and export moderators',
                                                            "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails), $($Moderator)",
                                                            $($_ | Out-String)
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                                ) + '"'
                                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                        }
                                    }
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export moderators',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsIdentityToIndex           = $AllRecipientsIdentityToIndex
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all recipients have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export AcceptMessagesOnlyFrom
    Write-Host
    Write-Host "Get and export AcceptMessagesOnlyFrom @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportAcceptMessagesOnlyFrom) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider) -and ($null -ne $AllRecipients[$x].AcceptMessagesOnlyFromSendersOrMembers)) {
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
                            $AllRecipientsIdentityToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export AcceptMessagesOnlyFrom @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                foreach ($AcceptedRecipient in $($Grantor.AcceptMessagesOnlyFromSendersOrMembers)) {
                                    try {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsIdentityToIndex[$($AcceptedRecipient)]
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $Trustee = $AllRecipients[$index]
                                        } else {
                                            $Trustee = $AcceptedRecipient
                                        }

                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'Cloud'
                                                } else {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                } else {
                                                    $TrusteeEnvironment = 'Cloud'
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                $ExportFileLines.add(
                                                            ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                '',
                                                                'AcceptMessagesOnlyFrom',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $AcceptedRecipient,
                                                                $Trustee.PrimarySmtpAddress,
                                                                $Trustee.DisplayName,
                                                                $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                )

                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Get and export AcceptMessagesOnlyFrom',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails), $($AcceptedRecipient)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export AcceptMessagesOnlyFrom',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsIdentityToIndex           = $AllRecipientsIdentityToIndex
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all recipients have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export ResourceDelegates
    Write-Host
    Write-Host "Get and export ResourceDelegates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportResourceDelegates) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider) -and (($null -ne $AllRecipients[$x].ResourceDelegates) -or ($null -ne $AllRecipients[$x].AllBookInPolicy) -or ($null -ne $AllRecipients[$x].BookInPolicy) -or ($null -ne $AllRecipients[$x].AllRequestInPolicy) -or ($null -ne $AllRecipients[$x].RequestInPolicy) -or ($null -ne $AllRecipients[$x].AllRequestOutOfPolicy) -or ($null -ne $AllRecipients[$x].RequestOutOfPolicy))) {
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
                            $AllRecipientsIdentityToIndex,
                            $AllRecipientsLegacyExchangeDnToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export ResourceDelegates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                foreach ($ResourceDelegatesSetting in @('ResourceDelegates', 'AllBookInPolicy', 'BookInPolicy', 'AllRequestInPolicy', 'RequestInPolicy', 'AllRequestOutOfPolicy', 'RequestOutOfPolicy')) {
                                    foreach ($AcceptedRecipient in $($Grantor.$ResourceDelegatesSetting)) {
                                        try {
                                            if ($AcceptedRecipient -is [boolean]) {
                                                $Trustee = 'Everyone'
                                            } else {
                                                try {
                                                    $index = $null
                                                    $index = $AllRecipientsIdentityToIndex[$($AcceptedRecipient)]
                                                } catch {
                                                }

                                                if ($index -ge 0) {
                                                    $Trustee = $AllRecipients[$index]
                                                } else {
                                                    try {
                                                        $index = $null
                                                        $index = $AllRecipientsLegacyExchangeDnToIndex[$($AcceptedRecipient)]
                                                    } catch {
                                                    }

                                                    if ($index -ge 0) {
                                                        $Trustee = $AllRecipients[$index]
                                                    } else {
                                                        $Trustee = $AcceptedRecipient
                                                    }
                                                }
                                            }

                                            if ($TrusteeFilter) {
                                                if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                    continue
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    } else {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                        $TrusteeEnvironment = 'On-Prem'
                                                    } else {
                                                        $TrusteeEnvironment = 'Cloud'
                                                    }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                    $ExportFileLines.add(
                                                                ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $(if ($ResourceDelegatesSetting -ieq 'ResourceDelegates') {
                                                                            'ResourceDelegate'
                                                                        } else {
                                                                            $("ResourcePolicyDelegate_$($ResourceDelegatesSetting)")
                                                                        }),
                                                                    $(if ($AcceptedRecipient -is [boolean]) {
                                                                            if ($AcceptedRecipient) {
                                                                                'Allow'
                                                                            } else {
                                                                                'Deny'
                                                                            }
                                                                        } else {
                                                                            'Allow'
                                                                        }),
                                                                    'False',
                                                                    'None',
                                                                    $(if ($AcceptedRecipient -is [boolean]) {
                                                                            $Trustee
                                                                        } else {
                                                                            $AcceptedRecipient
                                                                        }),
                                                                    $Trustee.PrimarySmtpAddress,
                                                                    $(@($Trustee.DisplayName, $Trustee, 'Warning: No valid info found') | Where-Object { $_ } | Select-Object -First 1),
                                                                    $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                    $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                                    $(if ($ExportSids) {
                                                                            try {
                                                                                try {
                                                                                    $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                                } catch {
                                                                                    if ($ExportFromOnPrem) {
                                                                                        # could be an object from a trust
                                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                        $objNT = $objTrans.GetType()
                                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        } else { '' }
                                                                    ),
                                                                    $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                    )

                                                }
                                            }
                                        } catch {
                                            (
                                                '"' + (
                                                    @(
                                                        (
                                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                            'Get and export ResourceDelegates',
                                                            "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails), $($AcceptedRecipient)",
                                                            $($_ | Out-String)
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                                ) + '"'
                                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                        }
                                    }
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export ResourceDelegates',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsIdentityToIndex           = $AllRecipientsIdentityToIndex
                        AllRecipientsLegacyExchangeDnToIndex   = $AllRecipientsLegacyExchangeDnToIndex
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all recipients have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export RequireAllSendersAreAuthenticated
    Write-Host
    Write-Host "Get and export RequireAllSendersAreAuthenticated @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportRequireAllSendersAreAuthenticated) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider) -and ($AllRecipients[$x].RequireAllSendersAreAuthenticated)) {
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
                            $AllRecipientsIdentityToIndex,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export RequireAllSendersAreAuthenticated @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $Trustee = 'NT AUTHORITY\Authenticated Users'

                                    if ($TrusteeFilter) {
                                        if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                            continue
                                        }
                                    }

                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                        if ($ExportFromOnPrem) {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                $TrusteeEnvironment = 'Cloud'
                                            } else {
                                                $TrusteeEnvironment = 'On-Prem'
                                            }
                                        } else {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                $TrusteeEnvironment = 'On-Prem'
                                            } else {
                                                $TrusteeEnvironment = 'Cloud'
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            $ExportFileLines.add(
                                                            ('"' + (@((
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                            $GrantorEnvironment,
                                                            '',
                                                            'RequireAllSendersAreAuthenticated',
                                                            'Allow',
                                                            'False',
                                                            'None',
                                                            $Trustee,
                                                            $Trustee.PrimarySmtpAddress,
                                                            $Trustee.DisplayName,
                                                            $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Trustee.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Trustee.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                            $TrusteeEnvironment
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                            )

                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export RequireAllSendersAreAuthenticated',
                                                    "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }


                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export RequireAllSendersAreAuthenticated',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                          = $AllRecipients
                        AllRecipientsIdentityToIndex           = $AllRecipientsIdentityToIndex
                        AllRecipientsSmtpToIndex               = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                  = $AllSecurityPrincipals
                        AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                        DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                        = $DebugPreference
                        ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                             = $ExportFile
                        ExportFileFilter                       = $ExportFileFilter
                        ExportFileHeader                       = $ExportFileHeader
                        ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                        ExportFromOnPrem                       = $ExportFromOnPrem
                        ExportGuids                            = $ExportGuids
                        ExportSids                             = $ExportSids
                        ExportTrustees                         = $ExportTrustees
                        ScriptPath                             = $PSScriptRoot
                        tempQueue                              = $tempQueue
                        TrusteeFilter                          = $TrusteeFilter
                        UTF8Encoding                           = $UTF8Encoding
                        VerbosePreference                      = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all recipients have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Calculate group membership
    Write-Host
    Write-Host "Calculate group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportManagementRoleGroupMembers -or $ExpandGroups -or ($ExportDistributionGroupMembers -ine 'None')) {
        Write-Host '  Create lookup hashtables'
        Write-Host "    GroupIdentity to group index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllGroupsIdentityToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllGroups.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllGroups.Count; $x++) {
            if ($AllGroups[$x].Identity) {
                $AllGroupsIdentityToIndex.Add($AllGroups[$x].Identity, $x)
            }
        }

        Write-Host "    GroupIdentity to recursive members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllGroupMembers = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllGroups.count, [StringComparer]::OrdinalIgnoreCase))

        # Normal distribution groups and management role groups
        for ($AllGroupsIndex = 0; $AllGroupsIndex -lt $AllGroups.count; $AllGroupsIndex++) {
            try {
                $AllRecipientsIndex = $AllRecipientsIdentityToIndex[$AllGroups[$AllGroupsIndex].Identity]
            } catch {
                $AllRecipientsIndex = $null
            }

            if (
                        ($ExportManagementRoleGroupMembers -and ($AllGroups[$AllGroupsIndex].RecipientTypeDetails -ieq 'RoleGroup')) -or
                        (($ExportDistributionGroupMembers -ieq 'All') -and ($AllRecipientsIndex -ge 0) -and ($AllRecipientsIndex -iin $GrantorsToConsider)) -or
                        ((($ExpandGroups) -or ($ExportDistributionGroupMembers -ieq 'OnlyTrustees')) -and ($AllRecipientsIndex -ge 0) -and ($AllRecipients[$AllRecipientsIndex].IsTrustee -eq $true))
            ) {
                if ($AllGroups[$AllGroupsIndex].Identity) {
                    $AllGroupMembers.Add($AllGroups[$AllGroupsIndex].Identity, @())
                }
            }
        }

        # Dynamic distribution groups
        for ($AllRecipientsIndex = 0; $AllRecipientsIndex -lt $AllRecipients.count; $AllRecipientsIndex++) {
            if ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails -ine 'DynamicDistributionGroup') {
                continue
            }

            if (
                        (($ExportDistributionGroupMembers -ieq 'All') -and ($AllRecipientsIndex -ge 0) -and ($AllRecipientsIndex -iin $GrantorsToConsider)) -or
                        ((($ExpandGroups) -or ($ExportDistributionGroupMembers -ieq 'OnlyTrustees')) -and ($AllRecipientsIndex -ge 0) -and ($AllRecipients[$AllRecipientsIndex].IsTrustee -eq $true))
            ) {
                $AllGroupMembers.Add($AllRecipients[$AllRecipientsIndex].Identity, @())
            }
        }

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllGroupMembers.count))

        foreach ($Enumerator in $AllGroupMembers.GetEnumerator()) {
            $tempQueue.enqueue($Enumerator.Name)
        }

        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        if ($ExportGroupMembersRecurse) {
            Write-Host '  Calculate recursive group membership'
        } else {
            Write-Host '  Calculate direct group membership'
        }

        Write-Host "    Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

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
                            $AllGroupMembers,
                            $AllGroups,
                            $AllGroupsIdentityToIndex,
                            $AllRecipients,
                            $AllRecipientsIdentityGuidToIndex,
                            $AllRecipientsIdentityToIndex,
                            $ConnectExchange,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExchangeCredential,
                            $ExchangeOnlineConnectionParameters,
                            $ExportFromOnPrem,
                            $ExportGroupMembersRecurse,
                            $FilterGetMember,
                            $ScriptPath,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $UseDefaultCredential,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Calculate group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            . ([scriptblock]::Create($ConnectExchange)) -NoReturnValue

                            . ([scriptblock]::Create($FilterGetMember))

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $GroupIdentity = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Group $($GroupIdentity) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    if ($ExportGroupMembersRecurse) {
                                        $AllGroupMembers[$GroupIdentity] = @($GroupIdentity | GetMemberRecurse | Sort-Object -Unique)
                                    } else {
                                        $AllGroupMembers[$GroupIdentity] = @($GroupIdentity | GetMemberRecurse -DirectMembersOnly | Sort-Object -Unique)
                                    }

                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Calculate recursive group membership',
                                                    "Group Identity $($GroupIdentity)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Calculate group membership',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            . ([scriptblock]::create($ConnectExchange)) -Disconnect

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllGroupMembers                    = $AllGroupMembers
                        AllGroups                          = $AllGroups
                        AllGroupsIdentityToIndex           = $AllGroupsIdentityToIndex
                        AllRecipients                      = $AllRecipients
                        AllRecipientsIdentityGuidToIndex   = $AllRecipientsIdentityGuidToIndex
                        AllRecipientsIdentityToIndex       = $AllRecipientsIdentityToIndex
                        ConnectExchange                    = $ConnectExchange
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                    = $DebugPreference
                        ErrorFile                          = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExchangeCredential                 = $ExchangeCredential
                        ExchangeOnlineConnectionParameters = $ExchangeOnlineConnectionParameters
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ExportGroupMembersRecurse          = $ExportGroupMembersRecurse
                        FilterGetMember                    = $FilterGetMember
                        ScriptPath                         = $PSScriptRoot
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        UseDefaultCredential               = $UseDefaultCredential
                        UTF8Encoding                       = $UTF8Encoding
                        VerbosePreference                  = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('    {0:0000000} groups to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all groups have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Management Role Group Members
    Write-Host
    Write-Host "Get and export Management Role Group members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExportManagementRoleGroupMembers) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllGroups.count))

        for ($x = 0; $x -lt $AllGroups.count; $x++) {
            if ($AllGroups[$x].RecipientTypeDetails -ieq 'RoleGroup') {
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
                            $AllGroupMembers,
                            $AllGroups,
                            $AllRecipients,
                            $AllRecipientsIdentityGuidToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsDisplaynameToIndex,
                            $AllSecurityPrincipalsDnToIndex,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $AllSecurityPrincipalsSidToIndex,
                            $AllSecurityPrincipalsUfnToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGroupMembersRecurse,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Management Role Group members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $AllGroupsId = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $RoleGroup = $AllGroups[$AllGroupsId]

                                if ($RoleGroup.identity -and $AllGroupMembers.containskey($RoleGroup.identity)) {
                                    $RoleGroupMembers = @($AllGroupMembers[$RoleGroup.Identity])
                                }

                                $GrantorPrimarySMTP = 'Management Role Group'
                                $GrantorDisplayName = $(($RoleGroup.DisplayName, $RoleGroup.Name, 'Warning: No valid info found') | Where-Object { $_ } | Select-Object -First 1)
                                $GrantorRecipientType = 'RoleGroup'

                                if ($ExportFromOnPrem) {
                                    $GrantorEnvironment = 'On-Prem'
                                } else {
                                    $GrantorEnvironment = 'Cloud'
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorDisplayName) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    foreach ($RoleGroupMember in $RoleGroupMembers) {
                                        if ($RoleGroupMember.tostring().startswith('NotARecipient:', 'CurrentCultureIgnoreCase')) {
                                            $Trustee = $RoleGroupMember -ireplace '^NotARecipient:', ''
                                        } else {
                                            try {
                                                $Trustee = $AllRecipients[$RoleGroupMember]
                                            } catch {
                                                $Trustee = $RoleGroupMember -ireplace '^NotARecipient:', ''
                                            }
                                        }

                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'Cloud'
                                                } else {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                } else {
                                                    $TrusteeEnvironment = 'Cloud'
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                $ExportFileLines.add(
                                                        ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $(if ($ExportGuids) { '' } else { '' }),
                                                                $(if ($ExportGuids) { $RoleGroup.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $GrantorRecipientType,
                                                                $GrantorEnvironment,
                                                                '',
                                                                $(if ($ExportGroupMembersRecurse) {
                                                                        'MemberRecurse'
                                                                    } else {
                                                                        'MemberDirect'
                                                                    }
                                                                ),
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $(($Trustee.PrimarySmtpAddress, $Trustee, '') | Select-Object -First 1),
                                                                $Trustee.PrimarySmtpAddress,
                                                                $Trustee.DisplayName,
                                                                $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) {
                                                                        $SecurityPrincipalsLookupSearchString = $(
                                                                            if ($Trustee.Guid.Guid) {
                                                                                $Trustee.Guid.Guid
                                                                            } else {
                                                                                $AllSecurityPrincipalsLookupSearchString = "$($Trustee)"

                                                                                $AllSecurityPrincipalsLookupResult = (
                                                                                    $AllSecurityPrincipalsDnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsObjectguidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsSidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsUfnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsDisplaynameToIndex[$AllSecurityPrincipalsLookupSearchString]
                                                                                ) | Where-Object { $_ } | Select-Object -First 1

                                                                                if ($AllSecurityPrincipalsLookupResult) {
                                                                                    if ($AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Sid.tostring().StartsWith('S-1-5-21-', 'CurrentCultureIgnoreCase')) {
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Guid.Guid
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                } else {
                                                                                    try {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AllSecurityPrincipalsLookupSearchString)"))
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            }
                                                                        ); $SecurityPrincipalsLookupSearchString
                                                                    } else { '' }
                                                                ),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                )

                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Get and export Management Role Group members',
                                                    "$($($GrantorPrimarySMTP), $($RoleGroupMember.RoleGroup), $($RoleGroupMember.TrusteeOriginalIdentity))",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    foreach ($ExportFileLine in $ExportFileLines) {
                                        try {
                                            $index = $null
                                            $index = $AllRecipientsSmtpToIndex[$ExportFileLine.'Trustee Primary SMTP']
                                        } catch {
                                        }

                                        if ($index -ge 0) {
                                            $AllRecipients[$index].IsTrustee = $true
                                        }
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.MRG{0:0000000}.txt' -f $AllGroupsId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Management Role Group members',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllGroupmembers                         = $AllGroupMembers
                        AllGroups                               = $AllGroups
                        AllRecipients                           = $AllRecipients
                        AllRecipientsIdentityGuidToIndex        = $AllRecipientsIdentityGuidToIndex
                        AllSecurityPrincipals                   = $AllSecurityPrincipals
                        AllSecurityPrincipalsDisplaynameToIndex = $AllSecurityPrincipalsDisplaynameToIndex
                        AllSecurityPrincipalsDnToIndex          = $AllSecurityPrincipalsDnToIndex
                        AllSecurityPrincipalsObjectguidToIndex  = $AllSecurityPrincipalsObjectguidToIndex
                        AllSecurityPrincipalsSidToIndex         = $AllSecurityPrincipalsSidToIndex
                        AllSecurityPrincipalsUfnToIndex         = $AllSecurityPrincipalsUfnToIndex
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                         = $DebugPreference
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                              = $ExportFile
                        ExportFileFilter                        = $ExportFileFilter
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileHeaderIndexes                 = $ExportFileHeaderIndexes
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExportGroupMembersRecurse               = $ExportGroupMembersRecurse
                        ExportGuids                             = $ExportGuids
                        ExportSids                              = $ExportSids
                        ExportTrustees                          = $ExportTrustees
                        ScriptPath                              = $PSScriptRoot
                        tempQueue                               = $tempQueue
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                        VerbosePreference                       = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} management role group members to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all management role group members have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Distribution Group Members
    # Must be the last export step because of '(($ExportDistributionGroupMembers -ieq 'OnlyTrustees') -and ($AllRecipients[$x].IsTrustee -eq $true))'
    Write-Host
    Write-Host "Get and export Distribution Group Members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (($ExportDistributionGroupMembers -ieq 'All') -or ($ExportDistributionGroupMembers -ieq 'OnlyTrustees')) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($AllRecipients[$x].RecipientTypeDetails -ilike 'Group*') -or ($AllRecipients[$x].RecipientTypeDetails -ilike '*Group')) {
                if ((($ExportDistributionGroupMembers -ieq 'All') -and ($x -in $GrantorsToConsider)) -or (($ExportDistributionGroupMembers -ieq 'OnlyTrustees') -and ($AllRecipients[$x].IsTrustee -eq $true))) {
                    if ($AllGroupMembers.ContainsKey($AllRecipients[$x].Identity)) {
                        $tempQueue.enqueue($x)
                    }

                    if (($ExportDistributionGroupMembers -ieq 'OnlyTrustees') -and (($x -notin $GrantorsToConsider))) {
                        $null = $GrantorsToConsider.add($x) # makes $ExportGrantorsWithNoPermissions work for these groups
                    }
                }
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
                            $AllGroupMembers,
                            $AllGroups,
                            $AllGroupsIdentityToIndex,
                            $AllRecipients,
                            $AllRecipientsIdentityToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsDisplaynameToIndex,
                            $AllSecurityPrincipalsDnToIndex,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $AllSecurityPrincipalsSidToIndex,
                            $AllSecurityPrincipalsUfnToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGroupMembersRecurse,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Distribution Group Members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $RecipientID = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                $GrantorRecipientType = $Grantor.RecipientType
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'Cloud'
                                    } else {
                                        $GrantorEnvironment = 'On-Prem'
                                    }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                        $GrantorEnvironment = 'On-Prem'
                                    } else {
                                        $GrantorEnvironment = 'Cloud'
                                    }
                                }

                                Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $GrantorMembers = @($AllGroupMembers[$Grantor.Identity])
                                } catch {
                                    continue
                                }

                                foreach ($index in $GrantorMembers) {
                                    if ($index.tostring().startswith('NotARecipient:', 'CurrentCultureIgnoreCase')) {
                                        $Trustee = $index -ireplace '^NotARecipient:', ''
                                    } else {
                                        $Trustee = $AllRecipients[$index]
                                    }

                                    try {
                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'Cloud'
                                                } else {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                    $TrusteeEnvironment = 'On-Prem'
                                                } else {
                                                    $TrusteeEnvironment = 'Cloud'
                                                }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                $ExportFileLines.add(
                                                        ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                '',
                                                                $(if ($ExportGroupMembersRecurse) {
                                                                        'MemberRecurse'
                                                                    } else {
                                                                        'MemberDirect'
                                                                    }),
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $(($Trustee.PrimarySmtpAddress, $Trustee, '') | Select-Object -First 1),
                                                                $Trustee.PrimarySmtpAddress,
                                                                $Trustee.DisplayName,
                                                                $(if ($ExportGuids) { $Trustee.ExchangeGuid.Guid } else { '' }),
                                                                $(if ($ExportGuids) {
                                                                        $SecurityPrincipalsLookupSearchString = $(
                                                                            if ($Trustee.Guid.Guid) {
                                                                                $Trustee.Guid.Guid
                                                                            } else {
                                                                                $AllSecurityPrincipalsLookupSearchString = "$($Trustee)"

                                                                                $AllSecurityPrincipalsLookupResult = (
                                                                                    $AllSecurityPrincipalsDnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsObjectguidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsSidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsUfnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                                    $AllSecurityPrincipalsDisplaynameToIndex[$AllSecurityPrincipalsLookupSearchString]
                                                                                ) | Where-Object { $_ } | Select-Object -First 1

                                                                                if ($AllSecurityPrincipalsLookupResult) {
                                                                                    if ($AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Sid.tostring().StartsWith('S-1-5-21-', 'CurrentCultureIgnoreCase')) {
                                                                                        $AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Guid.Guid
                                                                                    } else {
                                                                                        ''
                                                                                    }
                                                                                } else {
                                                                                    try {
                                                                                        if ($ExportFromOnPrem) {
                                                                                            # could be an object from a trust
                                                                                            # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                            $objNT = $objTrans.GetType()
                                                                                            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AllSecurityPrincipalsLookupSearchString)"))
                                                                                            $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                        } else {
                                                                                            ''
                                                                                        }
                                                                                    } catch {
                                                                                        ''
                                                                                    }
                                                                                }
                                                                            }
                                                                        ); $SecurityPrincipalsLookupSearchString
                                                                    } else { '' }
                                                                ),
                                                                $(if ($ExportSids) {
                                                                        try {
                                                                            try {
                                                                                $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                            } catch {
                                                                                if ($ExportFromOnPrem) {
                                                                                    # could be an object from a trust
                                                                                    # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } else {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    } else { '' }
                                                                ),
                                                                $("$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                                )

                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Get and export Distribution Group Members',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }

                                if ($ExportFileLines) {
                                    $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                    if ($ExportFileFilter) {
                                        $ExportFileLinesIndex = @()

                                        For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                            $ExportFileLine = $ExportFileLines[$x]
                                            if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                $ExportFileLinesIndex += $x
                                            }
                                        }

                                        $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                    }

                                    $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Get and export Distribution Group Members',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllGroupMembers                         = $AllGroupMembers
                        AllGroups                               = $AllGroups
                        AllGroupsIdentityToIndex                = $AllGroupsIdentityToIndex
                        AllRecipients                           = $AllRecipients
                        AllRecipientsIdentityToIndex            = $AllRecipientsIdentityToIndex
                        AllSecurityPrincipals                   = $AllSecurityPrincipals
                        AllSecurityPrincipalsDisplaynameToIndex = $AllSecurityPrincipalsDisplaynameToIndex
                        AllSecurityPrincipalsDnToIndex          = $AllSecurityPrincipalsDnToIndex
                        AllSecurityPrincipalsObjectguidToIndex  = $AllSecurityPrincipalsObjectguidToIndex
                        AllSecurityPrincipalsSidToIndex         = $AllSecurityPrincipalsSidToIndex
                        AllSecurityPrincipalsUfnToIndex         = $AllSecurityPrincipalsUfnToIndex
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                         = $DebugPreference
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                              = $ExportFile
                        ExportFileFilter                        = $ExportFileFilter
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileHeaderIndexes                 = $ExportFileHeaderIndexes
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExportGroupMembersRecurse               = $ExportGroupMembersRecurse
                        ExportGuids                             = $ExportGuids
                        ExportSids                              = $ExportSids
                        ExportTrustees                          = $ExportTrustees
                        ScriptPath                              = $PSScriptRoot
                        tempQueue                               = $tempQueue
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                        VerbosePreference                       = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} distribution groups to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all distribution groups have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Expand groups in temporary result files
    Write-Host
    Write-Host "Expand groups in temporary result files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if ($ExpandGroups) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($JobResultFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))) | Where-Object { $_.Length -gt 0 })) {
            $tempQueue.enqueue($JobResultFile.FullName)
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
                            $AllGroupMembers,
                            $AllGroups,
                            $AllGroupsIdentityToIndex,
                            $AllRecipients,
                            $AllRecipientsSmtpToIndex,
                            $AllSecurityPrincipals,
                            $AllSecurityPrincipalsDisplaynameToIndex,
                            $AllSecurityPrincipalsDnToIndex,
                            $AllSecurityPrincipalsObjectguidToIndex,
                            $AllSecurityPrincipalsSidToIndex,
                            $AllSecurityPrincipalsUfnToIndex,
                            $DebugFile,
                            $DebugPreference,
                            $ErrorFile,
                            $ExportFile,
                            $ExportFileFilter,
                            $ExportFileHeader,
                            $ExportFileHeaderIndexes,
                            $ExportFromOnPrem,
                            $ExportGroupMembersRecurse,
                            $ExportGuids,
                            $ExportSids,
                            $ExportTrustees,
                            $ScriptPath,
                            $tempQueue,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $VerbosePreference
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Expand groups in temporary result files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $JobResultFile = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "  $($JobResultFile) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                try {
                                    $ExportFileLines = [system.collections.arraylist]::new(1000)
                                    $ExportFileLinesOriginal = Import-Csv $JobResultFile -Encoding $UTF8Encoding -Delimiter ';'
                                    $ExportFileLinesExpanded = [system.collections.arraylist]::new(1000)

                                    foreach ($ExportFileLineOriginal in $ExportFileLinesOriginal) {
                                        if (($ExportFileLineOriginal.'Trustee Recipient Type' -ilike '*/Group*') -or ($ExportFileLineOriginal.'Trustee Recipient Type' -ilike '*Group')) {
                                            try {
                                                $Members = $null
                                                $Members = @($AllGroupMembers[$($AllRecipients[$($AllRecipientsSmtpToIndex[$($ExportFileLineOriginal.'Trustee Primary SMTP')])].Identity)])
                                            } catch {
                                                $Members = $null
                                            }

                                            if ($Members) {
                                                foreach ($Member in $Members) {
                                                    $ExportFileLineExpanded = $ExportFileLineOriginal.PSObject.Copy()

                                                    if ($Member.ToString().startswith('NotARecipient:', 'CurrentCultureIgnoreCase')) {
                                                        $Trustee = $Member -ireplace '^NotARecipient:', ''
                                                    } else {
                                                        $Trustee = $AllRecipients[$Member]
                                                    }

                                                    if ($ExportFromOnPrem) {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                            $TrusteeEnvironment = 'Cloud'
                                                        } else {
                                                            $TrusteeEnvironment = 'On-Prem'
                                                        }
                                                    } else {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') {
                                                            $TrusteeEnvironment = 'On-Prem'
                                                        } else {
                                                            $TrusteeEnvironment = 'Cloud'
                                                        }
                                                    }

                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress))) {
                                                        if ($ExportGroupMembersRecurse) {
                                                            $ExportFileLineExpanded.'Trustee Original Identity' = "$($ExportFileLineExpanded.'Trustee Original Identity')     [MemberRecurse] $(($Trustee.PrimarySmtpAddress, $Trustee.ToString(), '') | Select-Object -First 1)"
                                                        } else {
                                                            $ExportFileLineExpanded.'Trustee Original Identity' = "$($ExportFileLineExpanded.'Trustee Original Identity')     [MemberDirect] $(($Trustee.PrimarySmtpAddress, $Trustee.ToString(), '') | Select-Object -First 1)"
                                                        }
                                                        $ExportFileLineExpanded.'Trustee Primary SMTP' = $Trustee.PrimarySmtpAddress
                                                        $ExportFileLineExpanded.'Trustee Display Name' = $Trustee.DisplayName
                                                        if ($ExportGuids) {
                                                            $ExportFileLineExpanded.'Trustee Exchange GUID' = $Trustee.ExchangeGuid.Guid
                                                            $ExportFileLineExpanded.'Trustee AD ObjectGUID' = $(
                                                                if ($Trustee.Guid.Guid) {
                                                                    $Trustee.Guid.Guid
                                                                } else {
                                                                    $AllSecurityPrincipalsLookupSearchString = "$($Trustee)"

                                                                    $AllSecurityPrincipalsLookupResult = (
                                                                        $AllSecurityPrincipalsDnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                        $AllSecurityPrincipalsObjectguidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                        $AllSecurityPrincipalsSidToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                        $AllSecurityPrincipalsUfnToIndex[$AllSecurityPrincipalsLookupSearchString],
                                                                        $AllSecurityPrincipalsDisplaynameToIndex[$AllSecurityPrincipalsLookupSearchString]
                                                                    ) | Where-Object { $_ } | Select-Object -First 1

                                                                    if ($AllSecurityPrincipalsLookupResult) {
                                                                        if ($AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Sid.tostring().StartsWith('S-1-5-21-', 'CurrentCultureIgnoreCase')) {
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsLookupResult].Guid.Guid
                                                                        } else {
                                                                            ''
                                                                        }
                                                                    } else {
                                                                        try {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AllSecurityPrincipalsLookupSearchString)"))
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        } catch {
                                                                            ''
                                                                        }
                                                                    }
                                                                }
                                                            )
                                                        }
                                                        if ($ExportSids) {
                                                            $SecurityPrincipalsLookupSearchString = @($ExportFileLineExpanded.'Trustee AD ObjectGUID', $Trustee.Guid.Guid) | Where-Object { $_ } | Select-Object -First 1
                                                            $ExportFileLineExpanded.'Trustee SID' = $(
                                                                try {
                                                                    $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                } catch {
                                                                    if ($ExportFromOnPrem) {
                                                                        # could be an object from a trust
                                                                        # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                        $objNT = $objTrans.GetType()
                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                    } else {
                                                                        ''
                                                                    }
                                                                }
                                                            )
                                                        }
                                                        $ExportFileLineExpanded.'Trustee Recipient Type' = "$($Trustee.RecipientType)/$($Trustee.RecipientTypeDetails)" -ireplace '^/$', ''
                                                        $ExportFileLineExpanded.'Trustee Environment' = $TrusteeEnvironment
                                                    }

                                                    $ExportFileLinesExpanded.add($ExportFileLineExpanded)
                                                }
                                            }
                                        }
                                    }

                                    if ($ExportFileLinesExpanded) {
                                        $ExportFileLines = @(@($ExportFileLinesOriginal) + @($ExportFileLinesExpanded))
                                        $ExportFileLinesOriginal = $null
                                        $ExportFileLinesExpanded = $null

                                        if ($ExportFileFilter) {
                                            $ExportFileLinesIndex = @()

                                            For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                                $ExportFileLine = $ExportFileLines[$x]
                                                if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                    $ExportFileLinesIndex += $x
                                                }
                                            }
                                            $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                        }

                                        $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath $JobResultFile -Delimiter ';' -Encoding $UTF8Encoding -Force -NoTypeInformation
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                    'Expand groups in temporary result files',
                                                    "$($JobResultFile)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            'Expand groups in temporary result files',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllGroupMembers                         = $AllGroupMembers
                        AllGroups                               = $AllGroups
                        AllGroupsIdentityToIndex                = $AllGroupsIdentityToIndex
                        AllRecipients                           = $AllRecipients
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
                        AllSecurityPrincipals                   = $AllSecurityPrincipals
                        AllSecurityPrincipalsDisplaynameToIndex = $AllSecurityPrincipalsDisplaynameToIndex
                        AllSecurityPrincipalsDnToIndex          = $AllSecurityPrincipalsDnToIndex
                        AllSecurityPrincipalsObjectguidToIndex  = $AllSecurityPrincipalsObjectguidToIndex
                        AllSecurityPrincipalsSidToIndex         = $AllSecurityPrincipalsSidToIndex
                        AllSecurityPrincipalsUfnToIndex         = $AllSecurityPrincipalsUfnToIndex
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugPreference                         = $DebugPreference
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFile                              = $ExportFile
                        ExportFileFilter                        = $ExportFileFilter
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileHeaderIndexes                 = $ExportFileHeaderIndexes
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        ExportGroupMembersRecurse               = $ExportGroupMembersRecurse
                        ExportGuids                             = $ExportGuids
                        ExportSids                              = $ExportSids
                        ExportTrustees                          = $ExportTrustees
                        ScriptPath                              = $PSScriptRoot
                        tempQueue                               = $tempQueue
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                        VerbosePreference                       = $VerbosePreference
                    }
                )

                $Handle = $Powershell.BeginInvoke()

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                [void]$runspaces.Add($Temp)
            }

            Write-Host ('  {0:0000000} files to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                        if ($x -eq 0) {
                            Write-Host
                        }
                        $lastCount = $x
                    }
                }
            }

            Write-Host (("`r") + ('    {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

            if ($tempQueue.count -ne 0) {
                Write-Host '    Not all files have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                # $runspace.PowerShell.Stop()
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.Close()
            $RunspacePool.Dispose()
            'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }

                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep -Seconds 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Export grantors with no permissions
    Write-Host
    Write-Host "Export grantors with no permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($ExportGrantorsWithNoPermissions) {
        # Recipients
        if ($GrantorsToConsider) {
            Write-Host "  Recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

            foreach ($x in $GrantorsToConsider) {
                if (($AllRecipients[$x].RecipientTypeDetails -ilike 'Group*') -or ($AllRecipients[$x].RecipientTypeDetails -ilike '*Group')) {
                    if ($ExportDistributionGroupMembers -ieq 'OnlyTrustees') {
                        if ($AllRecipients[$x].IsTrustee -eq $true) {
                            $tempQueue.enqueue($x)
                        }
                    } else {
                        $tempQueue.enqueue($x)
                    }
                } else {
                    $tempQueue.enqueue($x)
                }
            }

            $tempQueueCount = $tempQueue.count

            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

            Write-Host "    Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

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
                                $AllSecurityPrincipals,
                                $AllSecurityPrincipalsObjectguidToIndex,
                                $DebugFile,
                                $DebugPreference,
                                $ErrorFile,
                                $ExportFile,
                                $ExportFileFilter,
                                $ExportFileHeader,
                                $ExportFileHeaderIndexes,
                                $ExportFromOnPrem,
                                $ExportGuids,
                                $ExportSids,
                                $ScriptPath,
                                $tempQueue,
                                $UTF8Encoding,
                                $VerbosePreference
                            )

                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Export grantors with no permissions (recipients) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                while ($tempQueue.count -gt 0) {
                                    try {
                                        $RecipientID = $tempQueue.dequeue()
                                    } catch {
                                        continue
                                    }

                                    try {
                                        $JobResultFile = ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID)))

                                        if (((Test-Path -LiteralPath $JobResultFile) -eq $false) -or ((Get-Item -LiteralPath $JobResultFile).Length -eq 0)) {
                                            $ExportFileLines = [system.collections.arraylist]::new(1)

                                            $Grantor = $AllRecipients[$RecipientID]

                                            $GrantorDisplayName = $Grantor.DisplayName
                                            $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                            $GrantorRecipientType = $Grantor.RecipientType
                                            $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                            if ($ExportFromOnPrem) {
                                                if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                                    $GrantorEnvironment = 'Cloud'
                                                } else {
                                                    $GrantorEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                                    $GrantorEnvironment = 'On-Prem'
                                                } else {
                                                    $GrantorEnvironment = 'Cloud'
                                                }
                                            }

                                            Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                            $ExportFileLines.add(
                                                    ('"' + (@((
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                            $GrantorEnvironment,
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            ''
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                            )


                                            if ($ExportFileLines) {
                                                $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                                if ($ExportFileFilter) {
                                                    $ExportFileLinesIndex = @()

                                                    For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                                        $ExportFileLine = $ExportFileLines[$x]
                                                        if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                            $ExportFileLinesIndex += $x
                                                        }
                                                    }

                                                    $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                                }

                                                $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Export grantors with no permissions (recipients)',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                'Export grantors with no permissions (recipients)',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            AllRecipients                          = $AllRecipients
                            AllSecurityPrincipals                  = $AllSecurityPrincipals
                            AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                            DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugPreference                        = $DebugPreference
                            ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ExportFile                             = $ExportFile
                            ExportFileFilter                       = $ExportFileFilter
                            ExportFileHeader                       = $ExportFileHeader
                            ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                            ExportFromOnPrem                       = $ExportFromOnPrem
                            ExportGuids                            = $ExportGuids
                            ExportSids                             = $ExportSids
                            ScriptPath                             = $PSScriptRoot
                            tempQueue                              = $tempQueue
                            UTF8Encoding                           = $UTF8Encoding
                            VerbosePreference                      = $VerbosePreference
                        }
                    )

                    $Handle = $Powershell.BeginInvoke()

                    $temp = '' | Select-Object PowerShell, Handle, Object
                    $temp.PowerShell = $PowerShell
                    $temp.Handle = $Handle
                    [void]$runspaces.Add($Temp)
                }

                Write-Host ('    {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                            if ($x -eq 0) {
                                Write-Host
                            }
                            $lastCount = $x
                        }
                    }
                }

                Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

                if ($tempQueue.count -ne 0) {
                    Write-Host '      Not all recipients have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
                }

                foreach ($runspace in $runspaces) {
                    # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                    # $runspace.PowerShell.Stop()
                    $runspace.PowerShell.Dispose()
                }

                $RunspacePool.Close()
                $RunspacePool.Dispose()
                'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

                if ($DebugFile) {
                    $null = Stop-Transcript
                    Start-Sleep -Seconds 1
                    foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }

                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep -Seconds 1
            }
        }


        # Public Folders
        if ($ExportPublicFolderPermissions) {
            Write-Host "  Public Folders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllPublicFolders.count))

            foreach ($x in (0..($AllPublicFolders.count - 1))) {
                $tempQueue.enqueue($x)
            }
            $tempQueueCount = $tempQueue.count

            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

            Write-Host "    Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

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
                                $AllRecipients,
                                $AllRecipientsExchangeGuidToIndex,
                                $AllSecurityPrincipals,
                                $AllSecurityPrincipalsObjectguidToIndex,
                                $DebugFile,
                                $DebugPreference,
                                $ErrorFile,
                                $ExportFile,
                                $ExportFileFilter,
                                $ExportFileHeader,
                                $ExportFileHeaderIndexes,
                                $ExportFromOnPrem,
                                $ExportGuids,
                                $ExportSids,
                                $ScriptPath,
                                $tempQueue,
                                $UTF8Encoding,
                                $VerbosePreference
                            )

                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Export grantors with no permissions (Public Folders) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                while ($tempQueue.count -gt 0) {
                                    try {
                                        $PublicFolderId = $tempQueue.dequeue()
                                    } catch {
                                        continue
                                    }

                                    $folder = $AllPublicFolders[$PublicFolderId]

                                    $folder.folderpath = '/' + $($folder.folderpath -join '/')

                                    try {
                                        $RecipientID = $null
                                        $RecipientID = $AllRecipientsExchangeGuidToIndex[$($folder.ContentMailboxGuid.Guid)]
                                    } catch {
                                        continue
                                    }

                                    try {
                                        $JobResultFile = ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.PF{1:0000000}.txt' -f $RecipientID, $PublicFolderId)))

                                        if (((Test-Path -LiteralPath $JobResultFile) -eq $false) -or ((Get-Item -LiteralPath $JobResultFile).Length -eq 0)) {
                                            $ExportFileLines = [system.collections.arraylist]::new(1)

                                            $Grantor = $AllRecipients[$RecipientID]

                                            $GrantorDisplayName = $Grantor.DisplayName
                                            $GrantorPrimarySMTP = $Grantor.PrimarySmtpAddress
                                            $GrantorRecipientType = $Grantor.RecipientType
                                            $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails

                                            if ($ExportFromOnPrem) {
                                                if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                                    $GrantorEnvironment = 'Cloud'
                                                } else {
                                                    $GrantorEnvironment = 'On-Prem'
                                                }
                                            } else {
                                                if ($Grantor.RecipientTypeDetails -ilike 'Remote*') {
                                                    $GrantorEnvironment = 'On-Prem'
                                                } else {
                                                    $GrantorEnvironment = 'Cloud'
                                                }
                                            }

                                            Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                            $ExportFileLines.add(
                                                    ('"' + (@((
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            $(if ($ExportGuids) { $Grantor.ExchangeGuid.Guid } else { '' }),
                                                            $(if ($ExportGuids) { $Grantor.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -ireplace '^/$', ''),
                                                            $GrantorEnvironment,
                                                            $($folder.folderpath),
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            ''
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                            )


                                            if ($ExportFileLines) {
                                                $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                                if ($ExportFileFilter) {
                                                    $ExportFileLinesIndex = @()

                                                    For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                                        $ExportFileLine = $ExportFileLines[$x]
                                                        if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                            $ExportFileLinesIndex += $x
                                                        }
                                                    }

                                                    $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                                }

                                                $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.PF{1:0000000}.txt' -f $RecipientId, $PublicFolderId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Export grantors with no permissions (Public Folders)',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                'Export grantors with no permissions (Public Folders)',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            AllPublicFolders                       = $AllPublicFolders
                            AllRecipients                          = $AllRecipients
                            AllRecipientsExchangeGuidToIndex       = $AllRecipientsExchangeGuidToIndex
                            AllSecurityPrincipals                  = $AllSecurityPrincipals
                            AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                            DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugPreference                        = $DebugPreference
                            ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ExportFile                             = $ExportFile
                            ExportFileFilter                       = $ExportFileFilter
                            ExportFileHeader                       = $ExportFileHeader
                            ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                            ExportFromOnPrem                       = $ExportFromOnPrem
                            ExportGuids                            = $ExportGuids
                            ExportSids                             = $ExportSids
                            ScriptPath                             = $PSScriptRoot
                            tempQueue                              = $tempQueue
                            UTF8Encoding                           = $UTF8Encoding
                            VerbosePreference                      = $VerbosePreference
                        }
                    )

                    $Handle = $Powershell.BeginInvoke()

                    $temp = '' | Select-Object PowerShell, Handle, Object
                    $temp.PowerShell = $PowerShell
                    $temp.Handle = $Handle
                    [void]$runspaces.Add($Temp)
                }

                Write-Host ('    {0:0000000} Public Folders to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                            if ($x -eq 0) {
                                Write-Host
                            }
                            $lastCount = $x
                        }
                    }
                }

                Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

                if ($tempQueue.count -ne 0) {
                    Write-Host '      Not all Public Folders have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
                }

                foreach ($runspace in $runspaces) {
                    # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                    # $runspace.PowerShell.Stop()
                    $runspace.PowerShell.Dispose()
                }

                $RunspacePool.Close()
                $RunspacePool.Dispose()
                'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

                if ($DebugFile) {
                    $null = Stop-Transcript
                    Start-Sleep -Seconds 1
                    foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }

                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                if ($ResultFile) {
                    foreach ($JobResultFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ResultFile), ('TEMP.*.PF*.txt'))))) {
                        Get-Content -LiteralPath $JobResultFile -Encoding $UTF8Encoding | Select-Object * -Skip 1 | Out-File -LiteralPath ($JobResultFile.fullname -ireplace '\.PF\d{7}.txt$', '.txt') -Append -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobResultFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep -Seconds 1
            }
        }


        # Management Role Groups
        if ($ExportManagementRoleGroupMembers) {
            Write-Host "  Management Role Groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllGroups.count))

            for ($AllGroupsId = 0; $AllGroupsId -lt $AllGroups.count; $AllGroupsId++) {
                if ($AllGroups[$AllGroupsId].RecipientTypeDetails -ieq 'RoleGroup') {
                    $tempQueue.enqueue($AllGroupsId)
                }
            }

            $tempQueueCount = $tempQueue.count

            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

            Write-Host "    Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

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
                                $AllGroups,
                                $AllSecurityPrincipals,
                                $AllSecurityPrincipalsObjectguidToIndex,
                                $DebugFile,
                                $DebugPreference,
                                $ErrorFile,
                                $ExportFile,
                                $ExportFileFilter,
                                $ExportFileHeader,
                                $ExportFileHeaderIndexes,
                                $ExportFromOnPrem,
                                $ExportGuids,
                                $ExportSids,
                                $ScriptPath,
                                $tempQueue,
                                $UTF8Encoding,
                                $VerbosePreference
                            )

                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Export grantors with no permissions (Management Role Groups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                while ($tempQueue.count -gt 0) {
                                    try {
                                        $AllGroupsId = $tempQueue.dequeue()
                                    } catch {
                                        continue
                                    }

                                    try {
                                        $JobResultFile = ([io.path]::ChangeExtension(($ExportFile), ('TEMP.MRG{0:0000000}.txt' -f $AllGroupsId)))

                                        if (((Test-Path -LiteralPath $JobResultFile) -eq $false) -or ((Get-Item -LiteralPath $JobResultFile).Length -eq 0)) {
                                            $ExportFileLines = [system.collections.arraylist]::new(1)

                                            $RoleGroup = $AllGroups[$AllGroupsId]

                                            $GrantorPrimarySMTP = 'Management Role Group'
                                            $GrantorDisplayName = $(($RoleGroup.DisplayName, $RoleGroup.Name, 'Warning: No valid info found') | Where-Object { $_ } | Select-Object -First 1)
                                            $GrantorRecipientType = 'RoleGroup'

                                            if ($ExportFromOnPrem) {
                                                $GrantorEnvironment = 'On-Prem'
                                            } else {
                                                $GrantorEnvironment = 'Cloud'
                                            }

                                            Write-Host "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                            $ExportFileLines.add(
                                                    ('"' + (@((
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            '',
                                                            $(if ($ExportGuids) { $RoleGroup.Guid.Guid } else { '' }),
                                                            $(if ($ExportSids) {
                                                                    try {
                                                                        try {
                                                                            $SecurityPrincipalsLookupSearchString = $Grantor.Guid.Guid
                                                                            $AllSecurityPrincipals[$AllSecurityPrincipalsObjectguidToIndex[$SecurityPrincipalsLookupSearchString]].sid.ToString()
                                                                        } catch {
                                                                            if ($ExportFromOnPrem) {
                                                                                # could be an object from a trust
                                                                                # No SID check required, as NameTranslate can only resolve Domain SIDs anyhow
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (7, "$($SecurityPrincipalsLookupSearchString)")) # 7 = GUID
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } else {
                                                                                ''
                                                                            }
                                                                        }
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                } else { '' }
                                                            ),
                                                            $GrantorRecipientType,
                                                            $GrantorEnvironment,
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '',
                                                            '' ,
                                                            '' ,
                                                            '' ,
                                                            '',
                                                            ''
                                                        ) | ForEach-Object { $_ -ireplace '"', '""' })[$ExportFileHeaderIndexes] -join '";"') + '"')
                                            )


                                            if ($ExportFileLines) {
                                                $ExportFileLines = @($ExportFileLines | ConvertFrom-Csv -Delimiter ';' -Header $ExportFileHeader)

                                                if ($ExportFileFilter) {
                                                    $ExportFileLinesIndex = @()

                                                    For ($x = 0; $x -lt $ExportFileLines.count; $x++) {
                                                        $ExportFileLine = $ExportFileLines[$x]
                                                        if ((. ([scriptblock]::Create($ExportFileFilter))) -eq $true) {
                                                            $ExportFileLinesIndex += $x
                                                        }
                                                    }

                                                    $ExportFileLines = @($ExportFileLines[$ExportFileLinesIndex])
                                                }

                                                $ExportFileLines | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv -LiteralPath([io.path]::ChangeExtension(($ExportFile), ('TEMP.MRG{0:0000000}.txt' -f $AllGroupsId))) -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                        'Export grantors with no permissions (Management Role Groups)',
                                                        "Exchange GUID $($Grantor.ExchangeGuid.Guid), Directory GUID $($Grantor.Guid.Guid), Primary SMTP $($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                    }
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                'Export grantors with no permissions (Management Role Groups)',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            AllGroups                              = $AllGroups
                            AllSecurityPrincipals                  = $AllSecurityPrincipals
                            AllSecurityPrincipalsObjectguidToIndex = $AllSecurityPrincipalsObjectguidToIndex
                            DebugFile                              = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugPreference                        = $DebugPreference
                            ErrorFile                              = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ExportFile                             = $ExportFile
                            ExportFileFilter                       = $ExportFileFilter
                            ExportFileHeader                       = $ExportFileHeader
                            ExportFileHeaderIndexes                = $ExportFileHeaderIndexes
                            ExportFromOnPrem                       = $ExportFromOnPrem
                            ExportGuids                            = $ExportGuids
                            ExportSids                             = $ExportSids
                            ScriptPath                             = $PSScriptRoot
                            tempQueue                              = $tempQueue
                            UTF8Encoding                           = $UTF8Encoding
                            VerbosePreference                      = $VerbosePreference
                        }
                    )

                    $Handle = $Powershell.BeginInvoke()

                    $temp = '' | Select-Object PowerShell, Handle, Object
                    $temp.PowerShell = $PowerShell
                    $temp.Handle = $Handle
                    [void]$runspaces.Add($Temp)
                }

                Write-Host ('    {0:0000000} Management Role Groups to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                            if ($x -eq 0) {
                                Write-Host
                            }
                            $lastCount = $x
                        }
                    }
                }

                Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

                if ($tempQueue.count -ne 0) {
                    Write-Host '      Not all Management Role Groups have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
                }

                foreach ($runspace in $runspaces) {
                    # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                    # $runspace.PowerShell.Stop()
                    $runspace.PowerShell.Dispose()
                }

                $RunspacePool.Close()
                $RunspacePool.Dispose()
                'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

                if ($DebugFile) {
                    $null = Stop-Transcript
                    Start-Sleep -Seconds 1
                    foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }

                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep -Seconds 1
            }
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }
} catch {
    Write-Host 'Unexpected error. Exiting.'
    $_
    (
        '"' + (
            @(
                (
                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                    '',
                    '',
                    $($_ | Out-String)
                ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
        ) + '"'
    ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
} finally {
    Write-Host
    Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    Write-Host "  Exchange connection(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    . ([scriptblock]::create($ConnectExchange)) -Disconnect -indent 4

    Write-Host "  Runspaces and RunspacePool @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($runspaces) {
        foreach ($runspace in $runspaces) {
            # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
            # $runspace.PowerShell.Stop()
            $runspace.PowerShell.Dispose()
        }
    }
    if ($RunspacePool) {
        $RunspacePool.Close()
        $RunspacePool.Dispose()
    }

    if ($ExportFile) {
        Write-Host "  Combine temporary export files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $JobResultFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))

        if ($JobResultFiles.count -gt 0) {
            Write-Host ('    {0:0000000} files to combine' -f $JobResultFiles.count)

            $ChunkSize = [math]::max(2, [math]::ceiling($JobResultFiles.count / $ParallelJobsLocal))

            $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new([Math]::Ceiling($JobResultFiles.count / $ChunkSize)))

            for ($x = 0; $x -lt $JobResultFiles.count; $x += $ChunkSize) {
                $tempQueue.enqueue(@($JobResultFiles[$x..$($x + $ChunkSize - 1)].fullname))
            }

            $tempQueueCount = $tempQueue.count

            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)


            if ($ParallelJobsNeeded -ge 1) {
                Write-Host ('      Pre-combine files' -f $JobResultFiles.count)
                Write-Host "        Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"


                $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
                $RunspacePool.Open()

                $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

                1..$ParallelJobsNeeded | ForEach-Object {
                    $Powershell = [powershell]::Create()
                    $Powershell.RunspacePool = $RunspacePool

                    [void]$Powershell.AddScript(
                        {
                            param(
                                $DebugFile,
                                $DebugPreference,
                                $ErrorFile,
                                $ScriptPath,
                                $tempQueue,
                                $UTF8Encoding,
                                $VerbosePreference
                            )
                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Pre-combine temporary export files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                while ($tempQueue.count -gt 0) {
                                    try {
                                        $ExportFileArray = $tempQueue.dequeue()
                                    } catch {
                                        continue
                                    }

                                    Write-Host "Target file $($ExportFileArray[0]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                    if ($ExportFileArray.count -gt 1) {
                                        foreach ($ExportFileTemp in ($ExportFileArray[1..($ExportFileArray.count - 1)])) {
                                            try {
                                                if ((Get-Item -LiteralPath $ExportFileTemp).length -gt 0) {
                                                    Get-Content -LiteralPath $ExportFileTemp -Encoding $UTF8Encoding -Force | Select-Object -Skip 1 | Out-File -LiteralPath $ExportFileArray[0] -Append -Encoding $UTF8Encoding -Force
                                                }
                                                Remove-Item -LiteralPath $ExportFileTemp -Force
                                            } catch {
                                                (
                                                    '"' + (
                                                        @(
                                                            (
                                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                                'Pre-combine temporary export files',
                                                                "$($ExportFileTemp)",
                                                                $($_ | Out-String)
                                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                                    ) + '"'
                                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                                            }
                                        }
                                    }

                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                'Pre-combine temporary export files',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            DebugFile         = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugPreference   = $DebugPreference
                            ErrorFile         = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ScriptPath        = $PSScriptRoot
                            tempQueue         = $tempQueue
                            UTF8Encoding      = $UTF8Encoding
                            VerbosePreference = $VerbosePreference
                        }
                    )

                    $Handle = $Powershell.BeginInvoke()

                    $temp = '' | Select-Object PowerShell, Handle, Object
                    $temp.PowerShell = $PowerShell
                    $temp.Handle = $Handle
                    [void]$runspaces.Add($Temp)
                }

                Write-Host ('        {0:0000000} file consolidation jobs. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle | Where-Object { $_.IsCompleted -eq $False }).count -ne 0) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle | Where-Object { $_.IsCompleted -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`r") + ('          {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                            if ($x -eq 0) {
                                Write-Host
                            }
                            $lastCount = $x
                        }
                    }
                }

                Write-Host (("`r") + ('          {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')))

                if ($tempQueue.count -ne 0) {
                    Write-Host '          Not all files have been checked. Enable ErrorFile and DebugFile options and check the log files.' -ForegroundColor red
                }

                foreach ($runspace in $runspaces) {
                    # $null = $runspace.PowerShell.EndInvoke($runspace.handle)
                    # $runspace.PowerShell.Stop()
                    $runspace.PowerShell.Dispose()
                }

                $RunspacePool.Close()
                $RunspacePool.Dispose()
                'temp', 'powershell', 'handle', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

                if ($DebugFile) {
                    $null = Stop-Transcript
                    Start-Sleep -Seconds 1
                    foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }

                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Out-File -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Append -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep -Seconds 1
            }

            $JobResultFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))

            Write-Host ('    {0:0000000} pre-consolidated files to combine. Done (in steps of {1:0000000}):' -f $JobResultFiles.count, $UpdateInterval)
            Write-Host ('      {0:0000000} @{1}@' -f 0, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))

            $lastCount = 1

            foreach ($JobResultFile in $JobResultFiles) {
                if ($JobResultFile.length -gt 0) {
                    Get-Content -LiteralPath $JobResultFile -Encoding $UTF8Encoding -Force | Select-Object -Skip 1 | Out-File -LiteralPath $ExportFile -Encoding $UTF8Encoding -Append -Force
                }

                Remove-Item -LiteralPath $JobResultFile -Force

                if (($lastCount % $UpdateInterval -eq 0) -or ($lastcount -eq $JobResultFiles.count)) {
                    Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $lastcount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                    if ($lastcount -eq $JobResultFiles.count) {
                        Write-Host
                    }
                }

                $lastCount++
            }

        } else {
            Write-Host ('    {0:0000000} files to check.' -f $JobResultFiles.count)
        }

        Write-Host "    '$($ExportFile)'"
    }

    if ($ErrorFile) {
        Write-Host "  Sort and combine temporary error files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $JobErrorFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))

        $x = Import-Csv $errorfile -Delimiter ';' -Encoding $UTF8Encoding

        if ($JobErrorFiles.count -gt 0) {
            Write-Host ('    {0:0000000} files to combine. Done (in steps of {1:0000000}):' -f $JobErrorFiles.count, $UpdateInterval)
            Write-Host ('      {0:0000000} @{1}@' -f 0, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))


            $lastCount = 1

            foreach ($JobErrorFile in $JobErrorFiles) {
                if ($JobErrorFile.length -gt 0) {
                    $x = $x + (Import-Csv -LiteralPath $JobErrorFile -Delimiter ';' -Encoding $UTF8Encoding -Header $errorfileheader)
                }

                Remove-Item -LiteralPath $JobErrorFile -Force

                if (($lastCount % $UpdateInterval -eq 0) -or ($lastcount -eq $JobErrorFiles.count)) {
                    Write-Host (("`r") + ('      {0:0000000} @{1}@' -f $lastcount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
                    if ($lastcount -eq $JobErrorFiles.count) {
                        Write-Host
                    }
                }

                $lastCount++
            }

        } else {
            Write-Host ('    {0:0000000} files to check.' -f $JobResultFiles.count)
        }

        if ($x.count -gt 0) {
            $x | Sort-Object -Property $ErrorFileHeader | Export-Csv -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force -Delimiter ';' -NoTypeInformation
        }

        Write-Host "    '$($ErrorFile)'"
    }

    if ($DebugFile) {
        Write-Host "  Combine temporary debug files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $JobDebugFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))

        Write-Host ('    {0:0000000} files to combine.' -f $JobDebugFiles.count)
        Write-Host '    Sort and combine will be performed after the step ''End script'' to ensure a complete debug log.'

        Write-Host "    '$($DebugFile)'"
    }

    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($DebugFile) {
        $null = Stop-Transcript
        Start-Sleep -Seconds 1
        foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
            if ($JobDebugFile.length -gt 0) {
                Get-Content -LiteralPath $JobDebugFile -Encoding $UTF8Encoding -Raw | Out-File -LiteralPath $DebugFile -Encoding $UTF8Encoding -Append -Force
            }

            Remove-Item -LiteralPath $JobDebugFile -Force
        }
    }

    Remove-Variable * -ErrorAction SilentlyContinue
    [GC]::Collect(); Start-Sleep -Seconds 1
}
