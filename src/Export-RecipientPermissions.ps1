<#
.SYNOPSIS
Export-RecipientPermissions XXXVersionStringXXX
Document, filter and compare Exchange permissions: Mailbox access rights, mailbox folder permissions, public folder permissions, send as, send on behalf, managed by, linked master accounts, forwarders, management role groups, distribution group members
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
- group members
- management role group members

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
  .RecipientType.Value, .RecipientTypeDetails.Value
  .DisplayName
  .PrimarySmtpAddress: .Local, .Domain, .Address
  .EmailAddresses: .PrefixString, .IsPrimaryAddress, .SmtpAddress, .ProxyAddressString
    This attribute is an array. Code example:
      $GrantorFilter = "if ((`$Grantor.EmailAddresses.SmtpAddress -ilike 'AddressA@example.com') -or (`$Grantor.EmailAddresses.SmtpAddress -ilike 'Test*@example.com')) { `$true } else { `$false }"
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


.PARAMETER ExportFileFilter
Only report results where the filter criteria matches $true.
This filter works against every single row of the results found. ExportFile will only contain lines where this filter returns $true.
The $ExportFileLine contains an object with the header names from $ExportFile as string properties
    'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Folder', 'Permission', 'Allow/Deny', 'Inherited', 'InheritanceType', 'Trustee Original Identity', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment'

Example: "`$ExportFileFilter.'Trustee Environment' -ieq 'On-Prem'"
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
    [boolean]$ExportFromOnPrem = $false,
    [uri[]]$ExchangeConnectionUriList = ('https://outlook.office365.com/powershell-liveid'),
    [boolean]$UseDefaultCredential = $false,
    [string]$ExchangeCredentialUsernameFile = '.\Export-RecipientPermissions_CredentialUsername.txt',
    [string]$ExchangeCredentialPasswordFile = '.\Export-RecipientPermissions_CredentialPassword.txt',
    [hashtable]$ExchangeOnlineConnectionParameters = @{ Credential = $null },
    [int]$ParallelJobsExchange = $ExchangeConnectionUriList.count * 3,
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
    [boolean]$ExportManagementRoleGroupMembers = $false,
    [boolean]$ExportForwarders = $true,
    [ValidateSet('None', 'All', 'OnlyTrustees')]$ExportDistributionGroupMembers = 'None',
    [boolean]$ExportGroupMembersRecurse = $false,
    [boolean]$ExportGuids = $false,
    [boolean]$ExpandGroups = $false,
    [boolean]$ExportGrantorsWithNoPermissions = $false,
    [ValidateSet('All', 'OnlyValid', 'OnlyInvalid')]$ExportTrustees = 'All',
    [string]$ExportFile = '.\export\Export-RecipientPermissions_Result.csv',
    [string]$ErrorFile = '.\export\Export-RecipientPermissions_Error.csv',
    [string]$DebugFile = '',
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
                Write-Verbose "Exchange session either not yet established or not working on try $($RetryCount)."

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
                    $ExchangeOnlineConnectionParameters['CommandName'] = ('Get-DynamicDistributionGroup', 'Get-Group', 'Get-Mailbox', 'Get-MailboxFolderPermission', 'Get-MailboxFolderStatistics', 'Get-MailboxPermission', 'Get-MailPublicFolder', 'Get-Publicfolder', 'Get-PublicFolderClientPermission', 'Get-Recipient', 'Get-RecipientPermission', 'Get-SecurityPrincipal', 'Get-UnifiedGroup', 'Get-UnifiedGroupLinks')

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
        $AllRecipientsIndex = $null
        $AllGroupsIndex = $null

        if ($GroupToCheck -match '^([a-fA-F\d]{8})-([a-fA-F\d]{4})-([a-fA-F\d]{4})-([a-fA-F\d]{4})-([a-fA-F\d]{12})$') {
            $AllRecipientsIndex = $AllRecipientsIdentityGuidToIndex[$GroupToCheck]
            $AllGroupsIndex = $AllGroupsIdentityGuidToIndex[$GroupToCheck]

            if (($AllRecipientsIndex -ge 0) -and ($AllGroupsIndex -ge 0)) {
                If (($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails.Value -ilike '*Group') -or ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails.Value -ilike 'Group*')) {
                    $GroupToCheckType = 'Group'
                } else {
                    $GroupToCheckType = 'Unknown'
                }
            } elseif (($AllRecipientsIndex -ge 0) -and ($AllGroupsIndex -lt 0)) {
                if ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails.Value -ilike 'DynamicDistributionGroup') {
                    $GroupToCheckType = 'DynamicDistributionGroup'
                } elseif (($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails.Value -inotlike '*Group') -and ($AllRecipients[$AllRecipientsIndex].RecipientTypeDetails.Value -inotlike 'Group*')) {
                    $GroupToCheckType = 'User'
                } else {
                    $GroupToCheckType = 'Unknown'
                }
            } elseif (($AllRecipientsIndex -lt 0) -and ($AllGroupsIndex -ge 0)) {
                $GroupToCheckType = 'ManagementRoleGroup'
            } else {
                $GroupToCheckType = 'Unknown'
            }
        }


        if ($GroupToCheckType -ieq 'User') {
            $AllRecipientsIndex
        } elseif (($GroupToCheckType -ieq 'Group') -or ($GroupToCheckType -ieq 'ManagementRoleGroup')) {
            foreach ($member in $AllGroups[$AllGroupsIndex].members) {
                if ($DirectMembersOnly.IsPresent) {
                    if ($member.ObjectGuid.Guid -and $AllRecipientsIdentityGuidToIndex.ContainsKey($member.ObjectGuid.Guid)) {
                        $AllRecipientsIdentityGuidToIndex[$member.ObjectGuid.Guid]
                    } else {
                        # $member.ObjectGuid.Guid is not known in $AllRecipients
                        "NotARecipient:$($member.ToString())"
                    }
                } else {
                    if (($member.ObjectGuid.Guid) -and ($AllGroupsIdentityGuidToIndex.ContainsKey($member.ObjectGuid.Guid) -or $AllRecipientsIdentityGuidToIndex.ContainsKey($member.ObjectGuid.Guid))) {
                        if ($member.ObjectGuid.Guid -notin $script:GetMemberRecurseTempLoopProtection) {
                            $script:GetMemberRecurseTempLoopProtection += $member.ObjectGuid.Guid
                            $member.ObjectGuid.Guid | GetMemberRecurse -DoNotResetGetMemberRecurseTempLoopProtection
                        }
                    } else {
                        # $member.ObjectGuid.Guid is neither known in $AllRecipients, nor in $AllGroups
                        "NotARecipient:$($member.ToString())"
                    }
                }
            }
        } elseif ($GroupToCheckType -ieq 'DynamicDistributionGroup') {
            if ($ExportFromOnPrem) {
                try {
                    $DynamicGroup = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DynamicDistributionGroup -identity $args[0] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object RecipientFilter, RecipientContainer } -ArgumentList $GroupToCheck -ErrorAction Stop
                    $members = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientPreviewFilter $args[0] -OrganizationalUnit $args[1] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object Guid } -ArgumentList $DynamicGroup.RecipientFilter, $DynamicGroup.RecipientContainer.DistinguishedName -ErrorAction Stop)
                } catch {
                    . ([scriptblock]::Create($ConnectExchange))
                    $DynamicGroup = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DynamicDistributionGroup -identity $args[0] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object RecipientFilter, RecipientContainer } -ArgumentList $GroupToCheck -ErrorAction Stop
                    $members = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientPreviewFilter $args[0] -OrganizationalUnit $args[1] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object Guid } -ArgumentList $AllRecipients[$AllRecipientsIndex].RecipientFilter, $AllRecipients[$AllRecipientsIndex].RecipientContainer.DistinguishedName -ErrorAction Stop)
                }
            } else {
                try {
                    $members = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DynamicDistributionGroupMember -identity $args[0] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object Guid } -ArgumentList $GroupToCheck -ErrorAction Stop)
                } catch {
                    . ([scriptblock]::Create($ConnectExchange))
                    $members = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DynamicDistributionGroupMember -identity $args[0] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object Guid, RecipientTypeDetails } -ArgumentList $GroupToCheck -ErrorAction Stop)
                }
            }

            foreach ($member in $members) {
                if ($DirectMembersOnly.IsPresent) {
                    if ($member.Guid.Guid -and $AllRecipientsIdentityGuidToIndex.ContainsKey($member.Guid.Guid)) {
                        $AllRecipientsIdentityGuidToIndex[$member.Guid.Guid]
                    } else {
                        # $member.ObjectGuid.Guid is not known in $AllRecipients
                        "NotARecipient:$($member.ToString())"
                    }
                } else {
                    if (($member.Guid.Guid) -and ($AllGroupsIdentityGuidToIndex.ContainsKey($member.Guid.Guid) -or $AllRecipientsIdentityGuidToIndex.ContainsKey($member.Guid.Guid))) {
                        if ($member.Guid.Guid -notin $script:GetMemberRecurseTempLoopProtection) {
                            $script:GetMemberRecurseTempLoopProtection += $member.Guid.Guid
                            $member.Guid.Guid | GetMemberRecurse -DoNotResetGetMemberRecurseTempLoopProtection
                        }
                    } else {
                        # $member.ObjectGuid.Guid is neither known in $AllRecipients, nor in $AllGroups
                        "NotARecipient:$($member.ToString())"
                    }
                }
            }
        } else {
            if (($AllRecipientsIndex -ge 0) -and ($AllRecipients[$AllRecipientsIndex].UserFriendlyName)) {
                "NotARecipient:$($AllRecipients[$AllRecipientsIndex].UserFriendlyName)"
            } elseif (($AllGroupsIndex -ge 0) -and (($AllGroups[$AllGroupsIndex].DisplayName) -or ($AllGroups[$AllGroupsIndex].Name) -or ($AllGroups[$AllGroupsIndex].DistinguishedName))) {
                "NotARecipient:$(@(($AllGroups[$AllGroupsIndex].DistinguishedName), ($AllGroups[$AllGroupsIndex].Name), ($AllGroups[$AllGroupsIndex].DisplayName)) | Where-Object { $_ } | Select-Object -First 1)"
            } else {
                "NotARecipient:$($GroupToCheck)"
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
        $null = Start-Transcript -LiteralPath $DebugFile -Force
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
            Remove-Item -LiteralPath $ErrorFile -Force
        }
        try {
            foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
                Remove-Item -LiteralPath $JobErrorFile -Force
            }
        } catch {}
        $null = New-Item -Path $ErrorFile -Force
        '"Timestamp";"Task";"TaskDetail";"Error"' | Out-File $ErrorFile -Encoding $UTF8Encoding -Force
    }


    if ($DebugFile) {
        foreach ($JobDebugFile in @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
            Remove-Item -LiteralPath $JobDebugFile -Force
        }
    }


    if ($ExportFile) {
        if (Test-Path $ExportFile) {
            Remove-Item -LiteralPath $ExportFile -Force
        }
        try {
            foreach ($RecipientFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))) -ErrorAction stop)) {
                Remove-Item -LiteralPath $Recipientfile -Force
            }
        } catch {}

        $null = New-Item -Path $ExportFile -Force

        if ($ExportGuids) {
            $ExportFileHeader = @(
                'Grantor Primary SMTP',
                'Grantor Display Name',
                'Grantor Exchange GUID',
                'Grantor AD ObjectGUID',
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
                'Trustee Exchange GUID',
                'Trustee AD ObjectGUID',
                'Trustee Recipient Type',
                'Trustee Environment'
            )

        } else {
            $ExportFileHeader = @(
                'Grantor Primary SMTP',
                'Grantor Display Name',
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
                'Trustee Recipient Type',
                'Trustee Environment'
            )
        }

        ('"' + ($ExportFileHeader -join '";"') + '"') | Out-File $ExportFile -Encoding $UTF8Encoding -Force
    }


    $tempConnectionUriQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new(10000))
    while ($tempConnectionUriQueue.count -le 100000) {
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
            @('Identity', 'DistinguishedName', 'ExchangeGuid', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'WhenSoftDeleted') | ForEach-Object {
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

        if ($ExpandGroups -or $ExportManagementRoleGroupMembers -or ($ExportDistributionGroupMembers -ine 'None')) {
            @('RecipientFilter', 'RecipientContainer') | ForEach-Object {
                if ($RecipientPropertiesExtended -inotcontains $_) {
                    $RecipientPropertiesExtended += $_
                }
            }
        }

        $RecipientPropertiesExtended = @($RecipientPropertiesExtended | Sort-Object -Unique)
    }


    # Credentials
    Write-Host
    Write-Host "Exchange credentials @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
    Write-Host "Connect to Exchange for single-thread operations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host '  Single-thread Exchange operation'

    . ([scriptblock]::Create($ConnectExchange))


    # Import recipients
    Write-Host
    Write-Host "Import recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host '  Enumerate possible RecipientTypeDetails values'
    try {
        $null = @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientTypeDetails '!!!Fail!!!' -resultsize 1 -ErrorAction Stop -WarningAction silentlycontinue } -ErrorAction Stop))
    } catch {
        $null = $error[0].exception -match '(?!.*: )(.*)(")$'
        $RecipientTypeDetailsListUnchecked = $matches[1].trim() -split ', ' | Where-Object { $_ } | Sort-Object -Unique
    }

    $RecipientTypeDetailsList = @()

    foreach ($RecipientTypeDetail in $RecipientTypeDetailsListUnchecked) {
        try {
            $null = @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientTypeDetails $args[0] -resultsize 1 -ErrorAction Stop -WarningAction silentlycontinue } -ArgumentList $RecipientTypeDetail -ErrorAction Stop))
            $RecipientTypeDetailsList += $RecipientTypeDetail
        } catch {
        }
    }

    $tempChars = ([char[]](0..255) -clike '[A-Z0-9]')
    $Filters = @()

    foreach ($tempChar in $tempChars) {
        $Filters += "(name -like '$($tempChar)*')"
    }

    $tempChars = $null

    $filters += ($filters -join ' -and ').replace('(name -like ''', '(name -notlike ''')

    $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

    foreach ($RecipientTypeDetail in $RecipientTypeDetailsList) {
        foreach ($Filter in $Filters) {
            $tempQueue.enqueue((, $RecipientTypeDetail, $Filter))
        }
    }

    $RecipientTypeDetailsList = $null
    $Filters = $null

    Write-Host "  Default recipients, filtered by RecipientTypeDetails and Name @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                        $UTF8Encoding,
                        $RecipientProperties,
                        $RecipientPropertiesExtended
                    )

                    try {
                        $DebugPreference = 'Continue'

                        Set-Location $ScriptPath

                        if ($DebugFile) {
                            $null = Start-Transcript -LiteralPath $DebugFile -Force
                        }

                        Write-Host "Import Recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                        . ([scriptblock]::Create($ConnectExchange))

                        while ($tempQueue.count -gt 0) {
                            try {
                                $QueueArray = $tempQueue.dequeue()
                            } catch {
                                continue
                            }

                            Write-Host "RecipientTypeDetails '$($QueueArray[0])', Filter '$($QueueArray[1])' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            try {
                                try {
                                    $x = @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientTypeDetails $args[0] -Filter $args[1] -Properties $args[2] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[3] } -ArgumentList $QueueArray[0], $QueueArray[1], $RecipientProperties, $RecipientPropertiesExtended -ErrorAction Stop))
                                } catch {
                                    . ([scriptblock]::Create($ConnectExchange))
                                    $x = @((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -RecipientTypeDetails $args[0] -Filter $args[1] -Properties $args[2] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[3] } -ArgumentList $QueueArray[0], $QueueArray[1], $RecipientProperties, $RecipientPropertiesExtended -ErrorAction Stop))
                                }

                                if ($x) {
                                    $AllRecipients.AddRange(@($x))
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                'Import Recipients',
                                                "RecipientTypeDetails '$($QueueArray[0])', Filter '$($QueueArray[1])'",
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                            }
                        }
                    } catch {
                        (
                            '"' + (
                                @(
                                    (
                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                        'Import Recipients',
                                        '',
                                        $($_ | Out-String)
                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                            ) + '"'
                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                    RecipientProperties                = $RecipientProperties
                    RecipientPropertiesExtended        = $RecipientPropertiesExtended
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

        Write-Host ('    {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

        $lastCount = -1
        while (($runspaces.Handle.IsCompleted -contains $False)) {
            Start-Sleep -Seconds 1
            $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
            for ($x = $lastCount; $x -le $done; $x++) {
                if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    if ($x -eq 0) { Write-Host }
                    $lastCount = $x
                }
            }
        }

        if ($tempQueue.count -eq 0) {
            Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
            Write-Host
        } else {
            Write-Host
            Write-Host '    Not all queries have been performed. Enable DebugFile option and check log file.' -ForegroundColor red
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
                Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                Remove-Item -LiteralPath $JobDebugFile -Force
            }
            $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
        }

        if ($ErrorFile) {
            foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                Remove-Item -LiteralPath $JobErrorFile -Force
            }
        }

        [GC]::Collect(); Start-Sleep 1
    }

    Write-Host ('    {0:0000000} recipients found' -f $($AllRecipients.count))

    Write-Host "  Additional recipients of specific types @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "    Single-thread Exchange operations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host '      Migration mailboxes'
    try {
        $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -Migration -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
    } catch {
        . ([scriptblock]::Create($ConnectExchange))
        $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -Migration -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
    }

    if ($ExportFromOnPrem) {
        Write-Host '      Arbitration mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -Arbitration -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -Arbitration -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }

        Write-Host '      AuditLog mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -AuditLog -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -AuditLog -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }

        Write-Host '      AuxAuditLog mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -AuxAuditLog -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -AuxAuditLog -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }

        Write-Host '      Monitoring mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -Monitoring -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -Monitoring -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }

        Write-Host '      RemoteArchive mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -RemoteArchive -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -RemoteArchive -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }
    } else {
        Write-Host '      Inactive mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -InactiveMailboxOnly -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -InactiveMailboxOnly -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }
        Write-Host '      Softdeleted mailboxes'
        try {
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -SoftDeletedMailbox -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -SoftDeletedMailbox -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property $args[0] } -ArgumentList (, $RecipientPropertiesExtended) -ErrorAction Stop)))
        }
    }

    Write-Host '  Sort list by PrimarySmtpAddress'
    $AllRecipients.TrimToSize()
    $x = @($AllRecipients | Where-Object { $_.PrimarySmtpAddress.Address } | Sort-Object -Property @{Expression = { $_.PrimarySmtpAddress.Address } })
    $AllRecipients.clear()
    $AllRecipients.AddRange(@($x))
    $AllRecipients.TrimToSize()
    $x = $null

    Write-Host ('  {0:0000000} total recipients found' -f $($AllRecipients.count))


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
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Distribution groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DistributionGroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DistributionGroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Dynamic Distribution Groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DynamicDistributionGroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-DynamicDistributionGroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        }

        Write-Host "  Unified Groups (Microsoft 365 Groups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        try {
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-UnifiedGroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-UnifiedGroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto } -ErrorAction Stop))
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


    # Import additional forwarding addresses
    Write-Host
    Write-Host "Import additional forwarding addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportForwarders) {
        Write-Host '  Single-thread Exchange operation'

        $AdditionalForwardingAddresses = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

        try {
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -filter '(ForwardingAddress -ne $null) -or (ForwardingSmtpAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Mailbox -filter '(ForwardingAddress -ne $null) -or (ForwardingSmtpAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        }

        try {
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicFolder -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicFolder -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        }

        try {
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailUser -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailUser -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        }

        try {
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicFolder -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        } catch {
            . ([scriptblock]::Create($ConnectExchange))
            $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailPublicFolder -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
        }

        if ($ExportFromOnPrem) {
            try {
                $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-RemoteMailbox -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
            } catch {
                . ([scriptblock]::Create($ConnectExchange))
                $AdditionalForwardingAddresses.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-RemoteMailbox -filter '(ForwardingAddress -ne $null)' -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -Property Identity, ForwardingAddress, DeliverToMailboxAndForward } -ErrorAction Stop))
            }
        }

        $AdditionalForwardingAddresses.TrimToSize()

        Write-Host ('  {0:0000000} additional forwarding addresses found' -f $($AdditionalForwardingAddresses.count))

        Write-Host '  Convert imported data'
        foreach ($Recipient in $AllRecipients) {
            if ($Recipient.ExternalEmailAddress) {
                if ($Recipient.ExternalEmailAddress.SmtpAddress) {
                    $Recipient.ExternalEmailAddress = $Recipient.ExternalEmailAddress.SmtpAddress
                } else {
                    if ($Recipient.RecipientTypeDetails -ieq 'PublicFolder') {
                        $Recipient.ExternalEmailAddress = $null
                    } else {
                        $Recipient.ExternalEmailAddress = $Recipient.ExternalEmailAddress.ToString()
                    }
                }
            }
        }

        $AdditionalForwardingAddresses | ForEach-Object {
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
                        } else {
                            $AllRecipients[$GrantorIndex].ForwardingAddress = $_.ForwardingAddress.ToString()
                        }
                    }

                    $Grantor.ForwardingSmtpAddress = $_.ForwardingSmtpAddress.SmtpAddress
                    $Grantor.DeliverToMailboxAndForward = $_.DeliverToMailboxAndForward
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
    Write-Host "Single-thread operations completed, remove connection to Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if (($ExportFromOnPrem -eq $false) -and ((Get-Module -Name 'ExchangeOnlineManagement').count -ge 1)) {
        Disconnect-ExchangeOnline -Confirm:$false
        Remove-Module ExchangeOnlineManagement
    }

    if ($ExchangeSession) {
        Remove-PSSession -Session $ExchangeSession
    }

    [GC]::Collect(); Start-Sleep 1


    # Create lookup hashtables for GUID, DistinguishedName and PrimarySmtpAddress
    Write-Host
    Write-Host "Create lookup hashtables for GUID, DistinguishedName and PrimarySmtpAddress @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  DistinguishedName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsDnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].distinguishedname) {
            if ($AllRecipientsDnToIndex.ContainsKey($(($AllRecipients[$x]).distinguishedname))) {
                # Same DN defined multiple times - set index to $null
                Write-Verbose "    '$(($AllRecipients[$x]).distinguishedname)' is not unique."
                $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $null
            } else {
                $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $x
            }
        }
    }

    Write-Host "  IdentityGuid to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsIdentityGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].identity.objectguid.guid) {
            if ($AllRecipientsIdentityGuidToIndex.ContainsKey($(($AllRecipients[$x]).identity.objectguid.guid))) {
                # Same GUID defined multiple times - set index to $null
                Write-Verbose "    '$(($AllRecipients[$x]).identity.objectguid.guid)' is not unique."
                $AllRecipientsIdentityGuidToIndex[$(($AllRecipients[$x]).identity.objectguid.guid)] = $null
            } else {
                $AllRecipientsIdentityGuidToIndex[$(($AllRecipients[$x]).identity.objectguid.guid)] = $x
            }
        }
    }

    Write-Host "  ExchangeGuid to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsExchangeGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if (($AllRecipients[$x].ExchangeGuid.Guid) -and ($AllRecipients[$x].ExchangeGuid.Guid -ine '00000000-0000-0000-0000-000000000000')) {
            if ($AllRecipientsExchangeGuidToIndex.ContainsKey($(($AllRecipients[$x]).ExchangeGuid.Guid))) {
                # Same GUID defined multiple times - set index to $null
                Write-Verbose "    '$(($AllRecipients[$x]).ExchangeGuid.Guid)' is not unique."
                $AllRecipientsExchangeGuidToIndex[$(($AllRecipients[$x]).ExchangeGuid.Guid)] = $null
            } else {
                $AllRecipientsExchangeGuidToIndex[$(($AllRecipients[$x]).ExchangeGuid.Guid)] = $x
            }
        }
    }

    Write-Host "  EmailAddresses to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsSmtpToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.EmailAddresses.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if ($AllRecipients[$x].EmailAddresses) {
            foreach ($EmailAddress in (@($AllRecipients[$x].EmailAddresses.SmtpAddress | Where-Object { $_ }))) {
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


    # Import LinkedMasterAccounts
    Write-Host
    Write-Host "Import LinkedMasterAccounts of each mailbox by database @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import LinkedMasterAccounts @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                                (
                                                    '"' + (
                                                        @(
                                                            (
                                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                                'Import LinkedMasterAccounts',
                                                                "Mailbox Identity GUID $($mailbox.identity.objectguid.guid)",
                                                                $($_ | Out-String)
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                                    ) + '"'
                                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Import LinkedMasterAccounts',
                                                    "Mailbox database GUID $(MailboxDatabaseGuid)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Import LinkedMasterAccounts',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import UserFriendlyNames
    Write-Host
    Write-Host "Import UserFriendlyNames @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import UserFriendlyNames @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                        $securityprincipals = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-SecurityPrincipal -filter "$($args[0])" -resultsize unlimited -WarningAction silentlycontinue | Select-Object userfriendlyname, guid } -ArgumentList $filterstring -ErrorAction Stop)
                                    } catch {
                                        . ([scriptblock]::Create($ConnectExchange))
                                        $securityprincipals = @(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-SecurityPrincipal -filter "$($args[0])" -resultsize unlimited -WarningAction silentlycontinue | Select-Object userfriendlyname, guid } -ArgumentList $filterstring -ErrorAction Stop)
                                    }

                                    foreach ($securityprincipal in $securityprincipals) {
                                        try {
                                            Write-Host "  '$($securityprincipal.guid.guid)' = '$($securityprincipal.UserFriendlyName)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                            ($AllRecipients[$($AllRecipientsIdentityGuidToIndex[$($securityprincipal.guid.guid)])]).UserFriendlyName = $securityprincipal.UserFriendlyName
                                        } catch {
                                            (
                                                '"' + (
                                                    @(
                                                        (
                                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                            'Import UserFriendlyNames',
                                                            "Security principal GUID $($securityprincipal.guid.guid)",
                                                            $($_ | Out-String)
                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                                ) + '"'
                                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                        }
                                    }
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Import UserFriendlyNames',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
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
                # Same UserFriendlyName defined multiple time - set index to $null
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


    # Define Grantors
    Write-Host
    Write-Host "Define grantors by filtering recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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


    # Get and export Mailbox Access Rights
    Write-Host
    Write-Host "Get and export Mailbox Access Rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxAccessRights) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.RecipientTypeDetails.Value -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $AllRecipientsSmtpToIndex,
                            $ExportGuids
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Mailbox Access Rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    foreach ($MailboxPermission in
                                        @($(
                                                if ($ExportFromOnPrem) {
                                                    try {
                                                        Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxPermission -identity $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                    } catch {
                                                        . ([scriptblock]::Create($ConnectExchange))
                                                        Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxPermission -identity $args[0] -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                    }
                                                } else {
                                                    if ($GrantorRecipientTypeDetails -ine 'GroupMailbox') {
                                                        if ($Grantor.WhenSoftDeleted) {
                                                            try {
                                                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxPermission -identity $args[0] -SoftDeletedMailbox -IncludeSoftDeletedUserPermissions -IncludeUnresolvedPermissions -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                            } catch {
                                                                . ([scriptblock]::Create($ConnectExchange))
                                                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxPermission -identity $args[0] -SoftDeletedMailbox -IncludeSoftDeletedUserPermissions -IncludeUnresolvedPermissions -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                            }
                                                        } else {
                                                            try {
                                                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxPermission -identity $args[0] -IncludeSoftDeletedUserPermissions -IncludeUnresolvedPermissions -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                            } catch {
                                                                . ([scriptblock]::Create($ConnectExchange))
                                                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxPermission -identity $args[0] -IncludeSoftDeletedUserPermissions -IncludeUnresolvedPermissions -resultsize unlimited -ErrorAction Stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                            }
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
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                foreach ($Accessright in ($TrusteeRight.Accessrights -split ', ')) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                        if ($ExportGuids) {
                                                            $ExportFileLines.add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $Grantor.ExchangeGuid.Guid,
                                                                            $Grantor.Identity.ObjectGuid.Guid,
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
                                                                            $Trustee.ExchangeGuid.Guid,
                                                                            $(
                                                                                if ($trustee.identity.objectguid.guid) {
                                                                                    $trustee.identity.objectguid.guid
                                                                                } else {
                                                                                    try {
                                                                                        (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-SecurityPrincipal -Filter "Sid -eq $($args[0])" -ResultSize 1 -WarningAction SilentlyContinue } -ArgumentList $TrusteeRight.User.SecurityIdentifier -ErrorAction Stop).Guid.Guid
                                                                                    } catch {
                                                                                        . ([scriptblock]::Create($ConnectExchange))
                                                                                        (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-SecurityPrincipal -Filter "Sid -eq '$($args[0])'" -ResultSize 1 -WarningAction SilentlyContinue } -ArgumentList $TrusteeRight.User.SecurityIdentifier -ErrorAction Stop).Guid.Guid
                                                                                    }
                                                                                }
                                                                            ),
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment

                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        } else {
                                                            $ExportFileLines.add(
                                                                ('"' + (@((
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
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        }
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
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Mailbox Access Rights',
                                                    "$($GrantorPrimarySMTP)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Mailbox Access Rights',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
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
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileFilter                        = $ExportFileFilter
                        ExportGuids                             = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Mailbox Folder permissions
    Write-Host
    Write-Host "Get and export Mailbox Folder Permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxFolderPermissions) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.RecipientTypeDetails.Value -ilike '*Mailbox') -and ($x -in $GrantorsToConsider) -and ($Recipient.RecipientTypeDetails.Value -ine 'PublicFolderMailbox') -and (-not $Recipient.WhenSoftDeleted)) {
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Mailbox Folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    $Folders = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderStatistics -identity $args[0] -ErrorAction Stop -WarningAction silentlycontinue | Select-Object folderid, folderpath, foldertype } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop
                                } catch {
                                    . ([scriptblock]::Create($ConnectExchange))
                                    $Folders = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderStatistics -identity $args[0] -ErrorAction Stop -WarningAction silentlycontinue | Select-Object folderid, folderpath, foldertype } -ArgumentList $GrantorPrimarySMTP -ErrorAction Stop
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
                                                            (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderPermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                        } catch {
                                                            . ([scriptblock]::Create($ConnectExchange))
                                                            (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderPermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                        }
                                                    } else {
                                                        if ($GrantorRecipientTypeDetails -ieq 'groupmailbox') {
                                                            try {
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderPermission -identity $args[0] -groupmailbox -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            } catch {
                                                                . ([scriptblock]::Create($ConnectExchange))
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderPermission -identity $args[0] -groupmailbox -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            }
                                                        } else {
                                                            try {
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderPermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
                                                            } catch {
                                                                . ([scriptblock]::Create($ConnectExchange))
                                                                (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-MailboxFolderPermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)" -ErrorAction Stop)
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

                                                            if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }

                                                            if ($ExportGuids) {
                                                                $ExportFileLines.Add(
                                                                    ('"' + (@((
                                                                                $GrantorPrimarySMTP,
                                                                                $GrantorDisplayName,
                                                                                $Grantor.ExchangeGuid.Guid,
                                                                                $Grantor.Identity.ObjectGuid.Guid,
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
                                                                                $Trustee.ExchangeGuid.Guid,
                                                                                $(($Trustee.Identity.ObjectGuid.Guid, $FolderPermission.User.AdRecipient.Guid.Guid, '') | Select-Object -First 1),
                                                                                $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                                $TrusteeEnvironment
                                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                                )
                                                            } else {
                                                                $ExportFileLines.Add(
                                                                    ('"' + (@((
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
                                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                                )
                                                            }
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

                                                            if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }

                                                            if ($ExportGuids) {
                                                                $ExportFileLines.Add(
                                                                    ('"' + (@((
                                                                                $GrantorPrimarySMTP,
                                                                                $GrantorDisplayName,
                                                                                $Grantor.ExchangeGuid.Guid,
                                                                                $Grantor.Identity.ObjectGuid.Guid,
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
                                                                                $Trustee.ExchangeGuid.Guid,
                                                                                $(($Trustee.Identity.ObjectGuid.Guid, $FolderPermission.User.RecipientPrincipcal.Guid.Guid, '') | Select-Object -First 1),
                                                                                $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                                $TrusteeEnvironment
                                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                                )
                                                            } else {
                                                                $ExportFileLines.Add(
                                                                    ('"' + (@((
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
                                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                                )
                                                            }
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
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                        'Get and export Mailbox Folder permissions',
                                                        "$($GrantorPrimarySMTP):$($Folder.folderid) ($($Folder.folderpath))",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Mailbox Folder permissions',
                                            "($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        ExportFileHeader                                = $ExportFileHeader
                        ExportFileFilter                                = $ExportFileFilter
                        ExportGuids                                     = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Send As permissions
    Write-Host
    Write-Host "Get and export Send As permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Send As permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
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
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                    if ($ExportGuids) {
                                                        $ExportFileLines.add(
                                                            ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $Grantor.AllRecipientsExchangeGuidToIndex,
                                                                        $Grantor.Identity.ObjectGuid.Guid,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendAs',
                                                                        $entry.AccessControlType,
                                                                        $entry.IsInherited,
                                                                        $entry.InheritanceType,
                                                                        $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $Trustee.ExchangeGuid.Guid,
                                                                        $(($Trustee.Identity.ObjectGuid.Guid, $FolderPermission.User.RecipientPrincipcal.Guid.Guid, '') | Select-Object -First 1),
                                                                        $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                        )

                                                    } else {
                                                        $ExportFileLines.add(
                                                            ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendAs',
                                                                        $entry.AccessControlType,
                                                                        $entry.IsInherited,
                                                                        $entry.InheritanceType,
                                                                        $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                        )
                                                    }
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
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                foreach ($AccessRight in $entry.AccessRights) {
                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                        if ($ExportGuids) {
                                                            $ExportFileLines.add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $Grantor.ExchangeGuid.Guid,
                                                                            $Grantor.Identity.ObjectGuid.Guid,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            $AccessRight,
                                                                            $entry.AccessControlType,
                                                                            $entry.IsInherited,
                                                                            $entry.InheritanceType,
                                                                            $(($Trustee.displayname, $entry.trustee, '') | Select-Object -First 1),
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $Trustee.ExchangeGuid.Guid,
                                                                            $(($Trustee.Identity.ObjectGuid.Guid, $Trustee.ObjectGuid.Guid, '') | Select-Object -First 1),
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        } else {
                                                            $ExportFileLines.add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            $AccessRight,
                                                                            $entry.AccessControlType,
                                                                            $entry.IsInherited,
                                                                            $entry.InheritanceType,
                                                                            $(($Trustee.displayname, $entry.trustee, '') | Select-Object -First 1),
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        }
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
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Send As permissions',
                                                    "$($GrantorPrimarySMTP)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Send As permissions',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileFilter                        = $ExportFileFilter
                        ExportGuids                             = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Send On Behalf permissions
    Write-Host
    Write-Host "Get and export Send On Behalf permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                            $AllRecipientsSmtpToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $AllRecipientsSendonbehalf,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Send On Behalf permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
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
                                                    $trustee = $delegateBindDn
                                                }

                                                if ($TrusteeFilter) {
                                                    if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                        continue
                                                    }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                    if ($ExportGuids) {
                                                        $ExportFileLines.add(
                                                            ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $Grantor.ExchangeGuid.Guid,
                                                                        $Grantor.Identity.ObjectGuid.Guid,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendOnBehalf',
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $Trustee.ExchangeGuid.Guid,
                                                                        $(
                                                                            if ($Trustee.Identity.ObjectGuid.Guid) {
                                                                                $Trustee.Identity.ObjectGuid.Guid
                                                                            } else {
                                                                                try {
                                                                                    $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                    $objNT = $objTrans.GetType()
                                                                                    $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                    $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$Trustee"))
                                                                                    $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                                } catch {
                                                                                    ''
                                                                                }
                                                                            }
                                                                        ),
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                        )
                                                    } else {
                                                        $ExportFileLines.add(
                                                            ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        'SendOnBehalf',
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                        )
                                                    }
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
                                                        if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                    } else {
                                                        if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                    }

                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                        if ($ExportGuids) {
                                                            $ExportFileLines.add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $Grantor.ExchangeGuid.Guid,
                                                                            $Grantor.Identity.ObjectGuid.Guid,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            'SendOnBehalf',
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $(($Trustee.displayname, $Truste, '') | Select-Object -First 1),
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $Trustee.ExchangeGuid.Guid,
                                                                            $(($Trustee.Identity.ObjectGuid.Guid, $AccessRight.ObjectGuid.Guid, '') | Select-Object -First 1),
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        } else {
                                                            $ExportFileLines.add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                            $GrantorEnvironment,
                                                                            '',
                                                                            'SendOnBehalf',
                                                                            'Allow',
                                                                            'False',
                                                                            'None',
                                                                            $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                            $Trustee.PrimarySmtpAddress.address,
                                                                            $Trustee.DisplayName,
                                                                            $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        }
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
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Send On Behalf permissions',
                                                    "$($GrantorPrimarySMTP)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Send On Behalf permissions',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllRecipientsSmtpToIndex         = $AllRecipientsSmtpToIndex
                        ErrorFile                        = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                        = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                 = $ExportFromOnPrem
                        ScriptPath                       = $PSScriptRoot
                        AllRecipientsSendonbehalf        = $AllRecipientsSendonbehalf
                        VerbosePreference                = $VerbosePreference
                        DebugPreference                  = $DebugPreference
                        TrusteeFilter                    = $TrusteeFilter
                        UTF8Encoding                     = $UTF8Encoding
                        ExportFileHeader                 = $ExportFileHeader
                        ExportFileFilter                 = $ExportFileFilter
                        ExportGuids                      = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Managed By permissions
    Write-Host
    Write-Host "Get and export Managed By permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                            $AllRecipientsSmtpToIndex,
                            $DebugFile,
                            $ScriptPath,
                            $ExportFromOnPrem,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Managed By permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
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
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                if ($ExportGuids) {
                                                    $ExportFileLines.add(
                                                       ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $Grantor.ExchangeGuid.Guid,
                                                                    $Grantor.IdentityGuid.Guid,
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'ManagedBy',
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    $Trustee.ExchangeGuid.Guid,
                                                                    $(($Trustee.IdentityGuid.Guid, $TrusteeRight.ObjectGuid.Guid, '') | Select-Object -First 1),
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                } else {
                                                    $ExportFileLines.add(
                                                       ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'ManagedBy',
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.displayname, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Managed By permissions',
                                                    "$($GrantorPrimarySMTP)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Managed By permissions',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllRecipientsSmtpToIndex         = $AllRecipientsSmtpToIndex
                        ErrorFile                        = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                        = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                       = $PSScriptRoot
                        ExportFromOnPrem                 = $ExportFromOnPrem
                        VerbosePreference                = $VerbosePreference
                        DebugPreference                  = $DebugPreference
                        TrusteeFilter                    = $TrusteeFilter
                        UTF8Encoding                     = $UTF8Encoding
                        ExportFileHeader                 = $ExportFileHeader
                        ExportFileFilter                 = $ExportFileFilter
                        ExportGuids                      = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Linked Master Accounts
    Write-Host
    Write-Host "Get and export Linked Master Accounts @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportLinkedMasterAccount -and $ExportFromOnPrem) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.recipienttypedetails.Value -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
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
                            $AllRecipientsSmtpToIndex,
                            $DebugFile,
                            $ScriptPath,
                            $ExportFromOnPrem,
                            $VerbosePreference,
                            $DebugPreference,
                            $TrusteeFilter,
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Linked Master Accounts @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
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
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                if ($ExportGuids) {
                                                    $ExportFileLines.add(
                                                        ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $Grantor.ExchangeGuid.Guid,
                                                                    $Grantor.Identity.ObjectGuid.Guid,
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
                                                                    $Trustee.ExchangeGuid.Guid,
                                                                    $(
                                                                        if ($Trustee.Identity.ObjectGuid.Guid) {
                                                                            $Trustee.Identity.ObjectGuid.Guid
                                                                        } else {
                                                                            try {
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$Trustee"))
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        }
                                                                    ),
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                } else {
                                                    $ExportFileLines.add(
                                                        ('"' + (@((
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
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Linked Master Accounts',
                                                    "$($GrantorPrimarySMTP)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Linked Master Accounts',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllRecipientsSmtpToIndex                = $AllRecipientsSmtpToIndex
                        ErrorFile                               = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                               = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                              = $PSScriptRoot
                        ExportFromOnPrem                        = $ExportFromOnPrem
                        VerbosePreference                       = $VerbosePreference
                        DebugPreference                         = $DebugPreference
                        TrusteeFilter                           = $TrusteeFilter
                        UTF8Encoding                            = $UTF8Encoding
                        ExportFileHeader                        = $ExportFileHeader
                        ExportFileFilter                        = $ExportFileFilter
                        ExportGuids                             = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Public Folder permissions
    Write-Host
    Write-Host "Get and export Public Folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )
                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Public folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

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
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value

                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    $folder.folderpath = '/' + $($folder.folderpath -join '/')

                                    Write-Host "  Folder '$($folder.EntryId)' ('$($Folder.Folderpath)') @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if ($ExportGuids) {
                                                $ExportFileLines.Add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $Grantor.ExchangeGuid.Guid,
                                                                $Grantor.Identity.ObjectGuid.Guid,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                $($Folder.Folderpath),
                                                                'MailEnabled',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $(($Trustee.primarysmtpaddress.address, $Trustee, '') | Select-Object -First 1),
                                                                $($Trustee.primarysmtpaddress.address),
                                                                $($Trustee.displayname),
                                                                $Trustee.ExchangeGuid.Guid,
                                                                $(($Trustee.Identity.ObjectGuid.Guid, $Trustee.ObjectGuid.Guid, '') | Select-Object -First 1),
                                                                $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
                                            } else {
                                                $ExportFileLines.Add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                $($Folder.Folderpath),
                                                                'MailEnabled',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $(($Trustee.primarysmtpaddress.address, $Trustee, '') | Select-Object -First 1),
                                                                $($Trustee.primarysmtpaddress.address),
                                                                $($Trustee.displayname),
                                                                $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
                                            }
                                        }
                                    }

                                    foreach ($FolderPermissions in
                                        @($(
                                                try {
                                                    Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-PublicFolderClientPermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList $($Folder.EntryId) -ErrorAction Stop
                                                } catch {
                                                    . ([scriptblock]::Create($ConnectExchange))
                                                    Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-PublicFolderClientPermission -identity $args[0] -ErrorAction stop -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList $($Folder.EntryId) -ErrorAction Stop
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

                                                        if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }

                                                        if ($ExportGuids) {
                                                            $ExportFileLines.Add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $Grantor.ExchangeGuid.Guid,
                                                                            $Grantor.Identity.ObjectGuid.Guid,
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
                                                                            $Trustee.ExchangeGuid.Guid,
                                                                            $(($Trustee.Identity.ObjectGuid.Guid, $FolderPermission.User.AdRecipient.Guid.Guid, '') | Select-Object -First 1),
                                                                            $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        } else {
                                                            $ExportFileLines.Add(
                                                                ('"' + (@((
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
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        }
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

                                                        if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }

                                                        if ($ExportGuids) {
                                                            $ExportFileLines.Add(
                                                                ('"' + (@((
                                                                            $GrantorPrimarySMTP,
                                                                            $GrantorDisplayName,
                                                                            $Grantor.ExchangeGuid.Guid,
                                                                            $Grantor.Identity.ObjectGuid.Guid,
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
                                                                            $Trustee.ExchangeGuid.Guid,
                                                                            $(($Trustee.Identity.ObjectGuid.Guid, $FolderPermission.User.RecipientPrincipal.Guid.Guid, '') | Select-Object -First 1),
                                                                            $("$($Trustee.recipienttype.value)/$($Trustee.recipienttypedetails.value)" -replace '^/$', ''),
                                                                            $TrusteeEnvironment
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        } else {
                                                            $ExportFileLines.Add(
                                                                ('"' + (@((
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
                                                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                            )
                                                        }
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
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Public Folder permissions',
                                                    "$($GrantorPrimarySMTP):$($Folder.Entryd) ($($Folder.folderpath))",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Public Folder permissions',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        ExportFileHeader                               = $ExportFileHeader
                        ExportFileFilter                               = $ExportFileFilter
                        ExportGuids                                    = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    #Import-Csv $JobErrorFile -Encoding $UTF8Encoding -Delimiter ';' | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv $ErrorFile -Encoding $UTF8Encoding -Force -Append -NoTypeInformation -Delimiter ';'
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding -Force | Select-Object -Skip 1 | Sort-Object -Unique | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force

                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            if ($ResultFile) {
                foreach ($JobResultFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ResultFile), ('TEMP.*.PF*.txt'))))) {
                    Get-Content -LiteralPath $JobResultFile -Encoding $UTF8Encoding | Select-Object * -Skip 1 | Add-Content ($JobResultFile.fullname -replace '\.PF\d{7}.txt$', '.txt') -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobResultFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Forwarders
    Write-Host
    Write-Host "Get and export Forwarders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportForwarders) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if ($x -in $GrantorsToConsider) {
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Forwarders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
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
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                    if ($ExportGuids) {
                                                        $ExportFileLines.add(
                                                            ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $Grantor.ExchangeGuid.Guid,
                                                                        $Grantor.Identity.ObjectGuid.Guid,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        $('Forward_' + $ForwarderType + $(if ((-not $Grantor.DeliverToMailboxAndForward) -or ($ForwarderType -ieq 'ExternalEmailAddress')) { '_ForwardOnly' } else { '_DeliverAndForward' } )),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $($Grantor.$ForwarderType),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $Trustee.ExchangeGuid.Guid,
                                                                        $(($Trustee.Identity.ObjectGuid.Guid, $Trustee.ObjectGuid.Guid, '') | select-first 1),
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                        )
                                                    } else {
                                                        $ExportFileLines.add(
                                                            ('"' + (@((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        '',
                                                                        $('Forward_' + $ForwarderType + $(if ((-not $Grantor.DeliverToMailboxAndForward) -or ($ForwarderType -ieq 'ExternalEmailAddress')) { '_ForwardOnly' } else { '_DeliverAndForward' } )),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $($Grantor.$ForwarderType),
                                                                        $Trustee.PrimarySmtpAddress.address,
                                                                        $Trustee.DisplayName,
                                                                        $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                        $TrusteeEnvironment
                                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                        )
                                                    }
                                                }
                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                        'Get and export Forwarders',
                                                        "$($GrantorPrimarySMTP)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Forwarders',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        ExportFileHeader         = $ExportFileHeader
                        ExportFileFilter         = $ExportFileFilter
                        ExportGuids              = $ExportGuids
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import direct group membership
    Write-Host
    Write-Host "Import direct group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($ExportManagementRoleGroupMembers -or $ExpandGroups -or ($ExportDistributionGroupMembers -ieq 'All') -or ($ExportDistributionGroupMembers -ieq 'OnlyTrustees')) {
        $tempChars = ([char[]](0..255) -clike '[A-Z0-9]')
        $Filters = @()

        foreach ($tempChar in $tempChars) {
            $Filters += "(name -like '$($tempChar)*')"
        }

        $tempChars = $null

        $filters += ($filters -join ' -and ').replace('(name -like ''', '(name -notlike ''')

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
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Import direct group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $filter = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Filter '$($filter)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    try {
                                        $x = @(Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Group -Filter $args[0] -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object Name, DisplayName, Guid, Members, RecipientType, RecipientTypeDetails } -ArgumentList $filter -ErrorAction Stop -WarningAction SilentlyContinue | Sort-Object -Property @{expression = { ($_.DisplayName, $_.Name) | Where-Object { $_ } | Select-Object -First 1 } })
                                    } catch {
                                        . ([scriptblock]::Create($ConnectExchange))
                                        $x = @(Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Group -Filter $args[0] -ResultSize Unlimited -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object Name, DisplayName, Guid, Members, RecipientType, RecipientTypeDetails } -ArgumentList $filter -ErrorAction Stop -WarningAction SilentlyContinue | Sort-Object -Property @{expression = { ($_.DisplayName, $_.Name) | Where-Object { $_ } | Select-Object -First 1 } })
                                    }

                                    if ($x) {
                                        $AllGroups.AddRange(@($x))
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Import direct group membership',
                                                    "Filter '$($filter)'",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Import Recipients',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllGroups                          = $AllGroups
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

            Write-Host ('    {0:0000000} queries to perform. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '    Not all queries have been performed. Enable DebugFile option and check log file.' -ForegroundColor red
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }

        $AllGroups.TrimToSize()
        Write-Host ('    {0:0000000} groups with direct members found' -f $($AllGroups.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Calculate group membership
    Write-Host
    Write-Host "Calculate group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($ExportManagementRoleGroupMembers -or $ExpandGroups -or ($ExportDistributionGroupMembers -ine 'None')) {
        Write-Host '  Create lookup hashtable: GroupIdentityGuid to group index'
        $AllGroupsIdentityGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllGroups.count, [StringComparer]::OrdinalIgnoreCase))

        for ($x = 0; $x -lt $AllGroups.Count; $x++) {
            $AllGroupsIdentityGuidToIndex.Add($AllGroups[$x].Guid.Guid, $x)
        }

        Write-Host '  Create lookup hashtable: GroupIdentityGuid to recursive members'
        $AllGroupMembers = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllGroups.count, [StringComparer]::OrdinalIgnoreCase))

        # Normal distribution groups and management role groups
        for ($x = 0; $x -lt $AllGroups.count; $x++) {
            try {
                $index = $AllRecipientsIdentityGuidToIndex[$AllGroups[$x].Guid.Guid]
            } catch {
                $index = $null
            }

            if (
                ($ExportManagementRoleGroupMembers -and ($AllGroups[$x].RecipientTypeDetails.value -ieq 'RoleGroup')) -or
                (($ExportDistributionGroupMembers -ieq 'All') -and ($index -ge 0) -and ($index -iin $GrantorsToConsider)) -or
                ((($ExpandGroups) -or ($ExportDistributionGroupMembers -ieq 'TrusteesOnly')) -and ($index -ge 0) -and ($AllRecipients[$index].IsTrustee -eq $true))
            ) {
                $AllGroupMembers.Add($AllGroups[$x].Guid.Guid, @())
            }
        }

        # Dynamic distribution groups
        for ($index = 0; $index -lt $AllRecipients.count; $index++) {
            if ($AllRecipients[$index].RecipientTypeDetails.Value -ine 'DynamicDistributionGroup') {
                continue
            }

            if (
                (($ExportDistributionGroupMembers -ieq 'All') -and ($index -iin $GrantorsToConsider)) -or
                ((($ExpandGroups) -or ($ExportDistributionGroupMembers -ieq 'TrusteesOnly')) -and ($AllRecipients[$index].IsTrustee -eq $true))
            ) {
                $AllGroupMembers.Add($AllRecipients[$index].Identity.ObjectGuid.Guid, @())
            }
        }

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

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
                            $AllRecipients,
                            $AllGroups,
                            $AllRecipientsIdentityGuidToIndex,
                            $AllGroupsIdentityGuidToIndex,
                            $AllGroupMembers,
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
                            $UTF8Encoding,
                            $FilterGetMember,
                            $ExportGroupMembersRecurse
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Calculate group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            . ([scriptblock]::Create($ConnectExchange))

                            . ([scriptblock]::Create($FilterGetMember))

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $GroupIdentityGuid = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "Group $($GroupIdentityGuid) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    if ($ExportGroupMembersRecurse) {
                                        $AllGroupMembers[$GroupIdentityGuid] = @($GroupIdentityGuid | GetMemberRecurse | Sort-Object -Unique)
                                    } else {
                                        $AllGroupMembers[$GroupIdentityGuid] = @($GroupIdentityGuid | GetMemberRecurse -DirectMembersOnly | Sort-Object -Unique)
                                    }

                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Calculate recursive group membership',
                                                    "Group Identity GUID $($GroupIdentityGuid)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Calculate group membership',
                                            '',
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllGroups                          = $AllGroups
                        AllRecipientsIdentityGuidToIndex   = $AllRecipientsIdentityGuidToIndex
                        AllGroupsIdentityGuidToIndex       = $AllGroupsIdentityGuidToIndex
                        AllGroupMembers                    = $AllGroupMembers
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
                        FilterGetMember                    = $FilterGetMember
                        ExportGroupMembersRecurse          = $ExportGroupMembersRecurse
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

            Write-Host ('    {0:0000000} groups to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }

            if ($tempQueue.count -eq 0) {
                Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                Write-Host
            } else {
                Write-Host
                Write-Host '    Not all groups have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Management Role Group Members
    Write-Host
    Write-Host "Get and export Management Role Group members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportManagementRoleGroupMembers) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllGroups.count))

        for ($x = 0; $x -lt $AllGroups.count; $x++) {
            if ($AllGroups[$x].RecipientTypeDetails.Value -ieq 'RoleGroup') {
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
                            $AllGroups,
                            $AllGroupMembers,
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids,
                            $ExportGroupMembersRecurse
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Management Role Group members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

                                try {
                                    $AllGroupsId = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                $RoleGroup = $AllGroups[$AllGroupsId]

                                $RoleGroupMembers = @($AllGroupMembers[$RoleGroup.Guid.Guid])

                                $GrantorPrimarySMTP = 'Management Role Group'
                                $GrantorDisplayName = $(($RoleGroup.DisplayName, $RoleGroup.Name) | Where-Object { $_ } | Select-Object -First 1)
                                $GrantorRecipientType = 'RoleGroup'

                                if ($ExportFromOnPrem) {
                                    $GrantorEnvironment = 'On-Prem'
                                } else {
                                    $GrantorEnvironment = 'Cloud'
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorDisplayName) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    foreach ($RoleGroupMember in $RoleGroupMembers) {
                                        if ($RoleGroupMember.tostring().startswith('NotARecipient:', 'CurrentCultureIgnoreCase')) {
                                            $Trustee = $RoleGroupMember -replace '^NotARecipient:', ''
                                        } else {
                                            try {
                                                $Trustee = $AllRecipients[$RoleGroupMember]
                                            } catch {
                                                $Trustee = $RoleGroupMember -replace '^NotARecipient:', ''
                                            }
                                        }

                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                if ($ExportGuids) {
                                                    $ExportFileLines.add(
                                                        ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    '',
                                                                    $RoleGroup.Guid.Guid,
                                                                    $GrantorRecipientType,
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $(if ($ExportGroupMembersRecurse) { 'MemberRecurse' } else { 'MemberDirect' }),
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.PrimarySmtpAddress.Address, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    $Trustee.ExchangeGuid.Guid,
                                                                    $(
                                                                        if ($Trustee.Identity.ObjectGuid.Guid) {
                                                                            $Trustee.Identity.ObjectGuid.Guid
                                                                        } else {
                                                                            try {
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$Trustee"))
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        }
                                                                    ),
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                } else {
                                                    $ExportFileLines.add(
                                                        ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $GrantorRecipientType,
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $(if ($ExportGroupMembersRecurse) { 'MemberRecurse' } else { 'MemberDirect' }),
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.PrimarySmtpAddress.Address, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                }
                                            }
                                        }
                                    }
                                } catch {
                                    (
                                        '"' + (
                                            @(
                                                (
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Get and export Management Role Group members',
                                                    "$($($GrantorPrimarySMTP), $($RoleGroupMember.RoleGroup), $($RoleGroupMember.TrusteeOriginalIdentity))",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Management Role Group members',
                                            "$($($GrantorPrimarySMTP), $($RoleGroupMember.RoleGroup), $($RoleGroupMember.TrusteeOriginalIdentity))",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllGroups                        = $AllGroups
                        AllGroupmembers                  = $AllGroupMembers
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
                        ExportFileHeader                 = $ExportFileHeader
                        ExportFileFilter                 = $ExportFileFilter
                        ExportGuids                      = $ExportGuids
                        ExportGroupMembersRecurse        = $ExportGroupMembersRecurse
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Get and export Distribution Group Members
    # Must be the last export step because of '(($ExportDistributionGroupMembers -ieq 'OnlyTrustees') -and ($AllRecipients[$x].IsTrustee -eq $true))'
    Write-Host
    Write-Host "Get and export Distribution Group Members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (($ExportDistributionGroupMembers -ieq 'All') -or ($ExportDistributionGroupMembers -ieq 'OnlyTrustees')) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($AllRecipients[$x].RecipientTypeDetails.Value -ilike 'Group*') -or ($AllRecipients[$x].RecipientTypeDetails.Value -ilike '*Group')) {
                if ((($ExportDistributionGroupMembers -ieq 'All') -and ($x -in $GrantorsToConsider)) -or (($ExportDistributionGroupMembers -ieq 'OnlyTrustees') -and ($AllRecipients[$x].IsTrustee -eq $true))) {
                    if ($AllGroupMembers.ContainsKey($AllRecipients[$x].Identity.ObjectGuid.Guid)) {
                        $tempQueue.enqueue($x)
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
                            $AllRecipients,
                            $AllGroups,
                            $AllGroupsIdentityGuidToIndex,
                            $AllGroupMembers,
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids,
                            $ExportGroupMembersRecurse
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Get and export Distribution Group Members @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                $ExportFileLines = [system.collections.arraylist]::new(1000)

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
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    $GrantorMembers = @($AllGroupMembers[$Grantor.Identity.ObjectGuid.Guid])
                                } catch {
                                    continue
                                }

                                foreach ($index in $GrantorMembers) {
                                    if ($index.tostring().startswith('NotARecipient:', 'CurrentCultureIgnoreCase')) {
                                        $Trustee = $index -replace '^NotARecipient:', ''
                                    } else {
                                        $Trustee = $AllRecipients[$index]
                                    }

                                    try {
                                        if ($TrusteeFilter) {
                                            if ((. ([scriptblock]::Create($TrusteeFilter))) -ne $true) {
                                                continue
                                            }
                                        }

                                        if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                if ($ExportGuids) {
                                                    $ExportFileLines.add(
                                                        ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $Grantor.ExchangeGuid.Guid,
                                                                    $Grantor.Identity.ObjectGuid.Guid,
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $(if ($ExportGroupMembersRecurse) { 'MemberRecurse' } else { 'MemberDirect' }),
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.PrimarySmtpAddress.Address, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.Address,
                                                                    $Trustee.DisplayName,
                                                                    $Trustee.ExchangeGuid.Guid,
                                                                    $(
                                                                        if ($Trustee.Identity.ObjectGuid.Guid) {
                                                                            $Trustee.Identity.ObjectGuid.Guid
                                                                        } else {
                                                                            try {
                                                                                $objTrans = New-Object -ComObject 'NameTranslate'
                                                                                $objNT = $objTrans.GetType()
                                                                                $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                                $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$Trustee"))
                                                                                $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                            } catch {
                                                                                ''
                                                                            }
                                                                        }
                                                                    ),
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                } else {
                                                    $ExportFileLines.add(
                                                    ('"' + (@((
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $(if ($ExportGroupMembersRecurse) { 'MemberRecurse' } else { 'MemberDirect' }),
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.PrimarySmtpAddress.Address, $Trustee, '') | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.Address,
                                                                    $Trustee.DisplayName,
                                                                    $("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                    )
                                                }
                                            }
                                        }
                                    } catch {
                                        (
                                            '"' + (
                                                @(
                                                    (
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                        'Get and export Distribution Group Members',
                                                        "$($GrantorPrimarySMTP)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Get and export Distribution Group Members',
                                            "$($GrantorPrimarySMTP)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                        AllRecipientsIdentityGuidToIndex = $AllRecipientsIdentityGuidToIndex
                        tempQueue                        = $tempQueue
                        ExportFile                       = $ExportFile
                        ExportTrustees                   = $ExportTrustees
                        ErrorFile                        = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                        = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                       = $PSScriptRoot
                        ExportFromOnPrem                 = $ExportFromOnPrem
                        VerbosePreference                = $VerbosePreference
                        DebugPreference                  = $DebugPreference
                        TrusteeFilter                    = $TrusteeFilter
                        UTF8Encoding                     = $UTF8Encoding
                        ExportFileHeader                 = $ExportFileHeader
                        ExportFileFilter                 = $ExportFileFilter
                        AllGroups                        = $AllGroups
                        AllGroupsIdentityGuidToIndex     = $AllGroupsIdentityGuidToIndex
                        AllGroupMembers                  = $AllGroupMembers
                        ExportGuids                      = $ExportGuids
                        ExportGroupMembersRecurse        = $ExportGroupMembersRecurse
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

            Write-Host ('  {0:0000000} distribution groups to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

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
                Write-Host '  Not all distribution groups have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Expand groups in temporary result files
    Write-Host
    Write-Host "Expand groups in temporary result files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                            $AllRecipients,
                            $AllRecipientsSmtpToIndex,
                            $AllGroups,
                            $AllGroupsIdentityGuidToIndex,
                            $AllGroupMembers,
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
                            $UTF8Encoding,
                            $ExportFileHeader,
                            $ExportFileFilter,
                            $ExportGuids,
                            $ExportGroupMembersRecurse
                        )

                        try {
                            $DebugPreference = 'Continue'

                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -LiteralPath $DebugFile -Force
                            }

                            Write-Host "Expand groups in temporary result files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            while ($tempQueue.count -gt 0) {
                                try {
                                    $JobResultFile = $tempQueue.dequeue()
                                } catch {
                                    continue
                                }

                                Write-Host "  $($JobResultFile) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                try {
                                    $ExportFileLines = [system.collections.arraylist]::new(1000)
                                    $ExportFileLinesOriginal = Import-Csv $JobResultFile -Encoding $UTF8Encoding -Delimiter ';'
                                    $ExportFileLinesExpanded = [system.collections.arraylist]::new(1000)

                                    foreach ($ExportFileLineOriginal in $ExportFileLinesOriginal) {
                                        if (($ExportFileLineOriginal.'Trustee Recipient Type' -ilike '*/Group*') -or ($ExportFileLineOriginal.'Trustee Recipient Type' -ilike '*Group')) {
                                            try {
                                                $Members = $null
                                                $Members = @($AllGroupMembers[$($AllRecipients[$($AllRecipientsSmtpToIndex[$($ExportFileLineOriginal.'Trustee Primary SMTP')])].Identity.ObjectGuid.Guid)])
                                            } catch {
                                                $Members = $null
                                            }

                                            if ($Members) {
                                                foreach ($Member in $Members) {
                                                    $ExportFileLineExpanded = $ExportFileLineOriginal.PSObject.Copy()

                                                    if ($Member.tostring().startswith('NotARecipient:', 'CurrentCultureIgnoreCase')) {
                                                        $Trustee = $Member -replace '^NotARecipient:', ''
                                                    } else {
                                                        $Trustee = $AllRecipients[$Member]
                                                    }

                                                    if ($ExportFromOnPrem) {
                                                        if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                    } else {
                                                        if ($Trustee.RecipientTypeDetails.Value -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                    }

                                                    if (($ExportTrustees -ieq 'All') -or (($ExportTrustees -ieq 'OnlyInvalid') -and (-not $Trustee.PrimarySmtpAddress.address)) -or (($ExportTrustees -ieq 'OnlyValid') -and ($Trustee.PrimarySmtpAddress.address))) {
                                                        if ($ExportGroupMembersRecurse) {
                                                            $ExportFileLineExpanded.'Trustee Original Identity' = "$($ExportFileLineExpanded.'Trustee Original Identity')     [MemberRecurse] $(($Trustee.PrimarySmtpAddress.Address, $Trustee.ToString()) | Select-Object -First 1)"
                                                        } else {
                                                            $ExportFileLineExpanded.'Trustee Original Identity' = "$($ExportFileLineExpanded.'Trustee Original Identity')     [MemberDirect] $(($Trustee.PrimarySmtpAddress.Address, $Trustee.ToString()) | Select-Object -First 1)"
                                                        }
                                                        $ExportFileLineExpanded.'Trustee Primary SMTP' = $Trustee.PrimarySmtpAddress.Address
                                                        $ExportFileLineExpanded.'Trustee Display Name' = $Trustee.DisplayName
                                                        if ($ExportGuids) {
                                                            $ExportFileLineExpanded.'Trustee Exchange GUID' = $Trustee.ExchangeGuid.Guid
                                                            $ExportFileLineExpanded.'Trustee AD ObjectGUID' = $(
                                                                if ($Trustee.Identity.ObjectGuid.Guid) {
                                                                    $Trustee.Identity.ObjectGuid.Guid
                                                                } else {
                                                                    try {
                                                                        $objTrans = New-Object -ComObject 'NameTranslate'
                                                                        $objNT = $objTrans.GetType()
                                                                        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                                                                        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$Trustee"))
                                                                        $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
                                                                    } catch {
                                                                        ''
                                                                    }
                                                                }
                                                            )
                                                        }
                                                        $ExportFileLineExpanded.'Trustee Recipient Type' = "$($Trustee.RecipientType.Value)/$($Trustee.RecipientTypeDetails.Value)" -replace '^/$', ''
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
                                                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                    'Expand groups in temporary result files',
                                                    "$($JobResultFile)",
                                                    $($_ | Out-String)
                                                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                        ) + '"'
                                    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                }
                            }
                        } catch {
                            (
                                '"' + (
                                    @(
                                        (
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                            'Expand groups in temporary result files',
                                            "$($JobResultFile)",
                                            $($_ | Out-String)
                                        ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                ) + '"'
                            ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                = $AllRecipients
                        AllRecipientsSmtpToIndex     = $AllRecipientsSmtpToIndex
                        tempQueue                    = $tempQueue
                        ExportFile                   = $ExportFile
                        ExportTrustees               = $ExportTrustees
                        ErrorFile                    = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        DebugFile                    = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath                   = $PSScriptRoot
                        ExportFromOnPrem             = $ExportFromOnPrem
                        VerbosePreference            = $VerbosePreference
                        DebugPreference              = $DebugPreference
                        TrusteeFilter                = $TrusteeFilter
                        UTF8Encoding                 = $UTF8Encoding
                        ExportFileHeader             = $ExportFileHeader
                        ExportFileFilter             = $ExportFileFilter
                        AllGroups                    = $AllGroups
                        AllGroupsIdentityGuidToIndex = $AllGroupsIdentityGuidToIndex
                        AllGroupMembers              = $AllGroupMembers
                        ExportGuids                  = $ExportGuids
                        ExportGroupMembersRecurse    = $ExportGroupMembersRecurse
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

            Write-Host ('  {0:0000000} files to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

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
                Write-Host '  Not all files have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobDebugFile -Force
                }
                $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
            }

            if ($ErrorFile) {
                foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                    Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                    Remove-Item -LiteralPath $JobErrorFile -Force
                }
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Export grantors with no permissions
    Write-Host
    Write-Host "Export grantors with no permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($ExportGrantorsWithNoPermissions) {
        # Recipients
        if ($GrantorsToConsider) {
            Write-Host "  Recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

            $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

            foreach ($x in $GrantorsToConsider) {
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
                                $AllRecipients,
                                $tempQueue,
                                $ExportFile,
                                $ErrorFile,
                                $DebugFile,
                                $ScriptPath,
                                $ExportFromOnPrem,
                                $VerbosePreference,
                                $DebugPreference,
                                $UTF8Encoding,
                                $ExportFileHeader,
                                $ExportFileFilter,
                                $ExportGuids
                            )

                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Export grantors with no permissions (recipients) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                            $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                            $GrantorRecipientType = $Grantor.RecipientType.value
                                            $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value

                                            if ($ExportFromOnPrem) {
                                                if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                            }

                                            Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                            if ($ExportGuids) {
                                                $ExportFileLines.add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $Grantor.ExchangeGuid.Guid,
                                                                $Grantor.Identity.ObjectGuid.Guid,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
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
                                                                ''
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
                                            } else {
                                                $ExportFileLines.add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
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
                                                                ''
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
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
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                        'Export grantors with no permissions (recipients)',
                                                        "$($GrantorPrimarySMTP)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                    }
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                'Export grantors with no permissions (recipients)',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            AllRecipients     = $AllRecipients
                            tempQueue         = $tempQueue
                            ExportFile        = $ExportFile
                            ErrorFile         = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugFile         = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ScriptPath        = $PSScriptRoot
                            ExportFromOnPrem  = $ExportFromOnPrem
                            VerbosePreference = $VerbosePreference
                            DebugPreference   = $DebugPreference
                            UTF8Encoding      = $UTF8Encoding
                            ExportFileHeader  = $ExportFileHeader
                            ExportFileFilter  = $ExportFileFilter
                            ExportGuids       = $ExportGuids
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

                Write-Host ('    {0:0000000} recipients to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle.IsCompleted -contains $False)) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                            if ($x -eq 0) { Write-Host }
                            $lastCount = $x
                        }
                    }
                }

                if ($tempQueue.count -eq 0) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    Write-Host
                } else {
                    Write-Host
                    Write-Host '    Not all recipients have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
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
                        Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }
                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep 1
            }
        }


        # Public Folders
        if ($ExportPublicFolderPermissions) {
            Write-Host "  Public Folders @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                $AllRecipients,
                                $AllPublicFolders,
                                $AllRecipientsExchangeGuidToIndex,
                                $tempQueue,
                                $ExportFile,
                                $ErrorFile,
                                $DebugFile,
                                $ScriptPath,
                                $ExportFromOnPrem,
                                $VerbosePreference,
                                $DebugPreference,
                                $UTF8Encoding,
                                $ExportFileHeader,
                                $ExportFileFilter,
                                $ExportGuids
                            )

                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Export grantors with no permissions (Public Folders) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                            $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                            $GrantorRecipientType = $Grantor.RecipientType.value
                                            $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value

                                            if ($ExportFromOnPrem) {
                                                if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Grantor.RecipientTypeDetails.Value -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                            }

                                            Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                            if ($ExportGuids) {
                                                $ExportFileLines.add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $Grantor.ExchangeGuid.Guid,
                                                                $Grantor.Identity.ObjectGuid.Guid,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
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
                                                                ''
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
                                            } else {
                                                $ExportFileLines.add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
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
                                                                ''
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
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
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                        'Export grantors with no permissions (Public Folders)',
                                                        "$($GrantorPrimarySMTP)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                    }
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                'Export grantors with no permissions (Public Folders)',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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
                            AllPublicFolders                 = $AllPublicFolders
                            AllRecipientsExchangeGuidToIndex = $AllRecipientsExchangeGuidToIndex
                            tempQueue                        = $tempQueue
                            ExportFile                       = $ExportFile
                            ErrorFile                        = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugFile                        = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ScriptPath                       = $PSScriptRoot
                            ExportFromOnPrem                 = $ExportFromOnPrem
                            VerbosePreference                = $VerbosePreference
                            DebugPreference                  = $DebugPreference
                            UTF8Encoding                     = $UTF8Encoding
                            ExportFileHeader                 = $ExportFileHeader
                            ExportFileFilter                 = $ExportFileFilter
                            ExportGuids                      = $ExportGuids
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

                Write-Host ('    {0:0000000} Public Folders to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle.IsCompleted -contains $False)) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                            if ($x -eq 0) { Write-Host }
                            $lastCount = $x
                        }
                    }
                }

                if ($tempQueue.count -eq 0) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    Write-Host
                } else {
                    Write-Host
                    Write-Host '    Not all Public Folders have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
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
                        Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }
                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                if ($ResultFile) {
                    foreach ($JobResultFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ResultFile), ('TEMP.*.PF*.txt'))))) {
                        Get-Content -LiteralPath $JobResultFile -Encoding $UTF8Encoding | Select-Object * -Skip 1 | Add-Content ($JobResultFile.fullname -replace '\.PF\d{7}.txt$', '.txt') -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobResultFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep 1
            }
        }


        # Management Role Groups
        if ($ExportManagementRoleGroupMembers) {
            Write-Host "  Management Role Groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

            $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllGroups.count))

            for ($AllGroupsId = 0; $AllGroupsId -lt $AllGroups.count; $AllGroupsId++) {
                if ($AllGroups[$AllGroupsId].RecipientTypeDetails.Value -ieq 'RoleGroup') {
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
                                $tempQueue,
                                $ExportFile,
                                $ErrorFile,
                                $DebugFile,
                                $ScriptPath,
                                $ExportFromOnPrem,
                                $VerbosePreference,
                                $DebugPreference,
                                $UTF8Encoding,
                                $ExportFileHeader,
                                $ExportFileFilter,
                                $ExportGuids
                            )

                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Export grantors with no permissions (Management Role Groups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

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
                                            $GrantorDisplayName = $(($RoleGroup.DisplayName, $RoleGroup.Name) | Select-Object -First 1)
                                            $GrantorRecipientType = 'RoleGroup'

                                            if ($ExportFromOnPrem) {
                                                $GrantorEnvironment = 'On-Prem'
                                            } else {
                                                $GrantorEnvironment = 'Cloud'
                                            }

                                            Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                            if ($ExportGuids) {
                                                $ExportFileLines.add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                '',
                                                                $RoleGroup.Guid.Guid,
                                                                $GrantorRecipientType,
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
                                                                ''
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
                                            } else {
                                                $ExportFileLines.add(
                                                    ('"' + (@((
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                $("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
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
                                                                ''
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"') + '"')
                                                )
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
                                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                        'Export grantors with no permissions (Management Role Groups)',
                                                        "$($GrantorDisplayName)",
                                                        $($_ | Out-String)
                                                    ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                            ) + '"'
                                        ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                    }
                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                'Export grantors with no permissions (Management Role Groups)',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            AllGroups         = $AllGroups
                            tempQueue         = $tempQueue
                            ExportFile        = $ExportFile
                            ErrorFile         = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugFile         = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ScriptPath        = $PSScriptRoot
                            ExportFromOnPrem  = $ExportFromOnPrem
                            VerbosePreference = $VerbosePreference
                            DebugPreference   = $DebugPreference
                            UTF8Encoding      = $UTF8Encoding
                            ExportFileHeader  = $ExportFileHeader
                            ExportFileFilter  = $ExportFileFilter
                            ExportGuids       = $ExportGuids
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

                Write-Host ('    {0:0000000} Management Role Groups to check. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle.IsCompleted -contains $False)) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                            if ($x -eq 0) { Write-Host }
                            $lastCount = $x
                        }
                    }
                }

                if ($tempQueue.count -eq 0) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    Write-Host
                } else {
                    Write-Host
                    Write-Host '    Not all Management Role Groups have been checked. Enable DebugFile option and check log file.' -ForegroundColor red
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
                        Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }
                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep 1
            }
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }
} catch {
    Write-Host 'Unexpected error. Exiting.'
    (
        '"' + (
            @(
                (
                    $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                    '',
                    '',
                    $($_ | Out-String)
                ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
        ) + '"'
    ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
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

    if ($ExportFile) {
        Write-Host "  Combine temporary export files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
                                $tempQueue,
                                $ErrorFile,
                                $DebugFile,
                                $ScriptPath,
                                $VerbosePreference,
                                $DebugPreference,
                                $UTF8Encoding
                            )
                            try {
                                $DebugPreference = 'Continue'

                                Set-Location $ScriptPath

                                if ($DebugFile) {
                                    $null = Start-Transcript -LiteralPath $DebugFile -Force
                                }

                                Write-Host "Combine temporary export files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                while ($tempQueue.count -gt 0) {
                                    try {
                                        $ExportFileArray = $tempQueue.dequeue()
                                    } catch {
                                        continue
                                    }

                                    Write-Host "Target file $($ExportFileArray[0]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                    if ($ExportFileArray.count -gt 1) {
                                        foreach ($ExportFileTemp in ($ExportFileArray[1..($ExportFileArray.count)])) {
                                            try {
                                                if ((Get-Item -LiteralPath $ExportFileTemp).length -gt 0) {
                                                    Get-Content -LiteralPath $ExportFileTemp -Encoding $UTF8Encoding -Force | Select-Object -Skip 1 | Add-Content -LiteralPath $ExportFileArray[0] -Encoding $UTF8Encoding -Force
                                                }
                                                Remove-Item -LiteralPath $ExportFileTemp -Force
                                            } catch {
                                                (
                                                    '"' + (
                                                        @(
                                                            (
                                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                                'Combine temporary export files',
                                                                "$($ExportFileTemp)",
                                                                $($_ | Out-String)
                                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                                    ) + '"'
                                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                                            }
                                        }
                                    }

                                }
                            } catch {
                                (
                                    '"' + (
                                        @(
                                            (
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'),
                                                'Combine temporary export files',
                                                '',
                                                $($_ | Out-String)
                                            ) | ForEach-Object { $_ -replace '"', '""' }) -join '";"'
                                    ) + '"'
                                ) | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                            } finally {
                                if ($DebugFile) {
                                    $null = Stop-Transcript
                                    Start-Sleep -Seconds 1
                                }
                            }
                        }
                    ).AddParameters(
                        @{
                            tempQueue         = $tempQueue
                            ErrorFile         = ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            DebugFile         = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                            ScriptPath        = $PSScriptRoot
                            VerbosePreference = $VerbosePreference
                            DebugPreference   = $DebugPreference
                            UTF8Encoding      = $UTF8Encoding
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

                Write-Host ('        {0:0000000} file consolidation jobs. Done (in steps of {1:0000000}):' -f $tempQueueCount, $UpdateInterval)

                $lastCount = -1
                while (($runspaces.Handle.IsCompleted -contains $False)) {
                    Start-Sleep -Seconds 1
                    $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                    for ($x = $lastCount; $x -le $done; $x++) {
                        if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                            Write-Host (("`b" * 100) + ('          {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                            if ($x -eq 0) { Write-Host }
                            $lastCount = $x
                        }
                    }
                }

                if ($tempQueue.count -eq 0) {
                    Write-Host (("`b" * 100) + ('          {0:0000000} @{1}@' -f $tempQueueCount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    Write-Host
                } else {
                    Write-Host
                    Write-Host '        Not all files have been combined. Enable DebugFile option and check log file.' -ForegroundColor red
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
                        Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobDebugFile -Force
                    }
                    $null = Start-Transcript -LiteralPath $DebugFile -Append -Force
                }

                if ($ErrorFile) {
                    foreach ($JobErrorFile in @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))) {
                        Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
                        Remove-Item -LiteralPath $JobErrorFile -Force
                    }
                }

                [GC]::Collect(); Start-Sleep 1
            }

            $JobResultFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))

            Write-Host ('    {0:0000000} pre-consolidated files to combine. Done (in steps of {1:0000000}):' -f $JobResultFiles.count, $UpdateInterval)
            Write-Host ('      {0:0000000} @{1}@' -f 0, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))

            $lastCount = 1

            foreach ($JobResultFile in $JobResultFiles) {
                if ($JobResultFile.length -gt 0) {
                    #Import-Csv $JobResultFile -Encoding $UTF8Encoding -Delimiter ';' | Sort-Object -Property $ExportFileHeader -Unique | Export-Csv $ExportFile -Delimiter ';' -Encoding $UTF8Encoding -Force -Append -NoTypeInformation
                    Get-Content -LiteralPath $JobResultFile -Encoding $UTF8Encoding -Force | Select-Object -Skip 1 | Add-Content -LiteralPath $ExportFile -Encoding $UTF8Encoding -Force
                }

                Remove-Item -LiteralPath $JobResultFile -Force

                if (($lastCount % $UpdateInterval -eq 0) -or ($lastcount -eq $JobResultFiles.count)) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $lastcount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    if ($lastcount -eq $JobResultFiles.count) { Write-Host }
                }

                $lastCount++
            }

        } else {
            Write-Host ('    {0:0000000} files to check.' -f $JobResultFiles.count)
        }

        Write-Host "    '$($ExportFile)'"
    }

    if ($ErrorFile) {
        Write-Host "  Sort and combine temporary error files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $JobErrorFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($ErrorFile), ('TEMP.*.txt'))))

        if ($JobErrorFiles.count -gt 0) {
            Write-Host ('    {0:0000000} files to combine. Done (in steps of {1:0000000}):' -f $JobErrorFiles.count, $UpdateInterval)
            Write-Host ('      {0:0000000} @{1}@' -f 0, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))

            $x = @()

            $lastCount = 1

            foreach ($JobErrorFile in $JobErrorFiles) {
                if ($JobErrorFile.length -gt 0) {
                    $x += @(Get-Content -LiteralPath $JobErrorFile -Encoding $UTF8Encoding)
                }

                Remove-Item -LiteralPath $JobErrorFile -Force

                if (($lastCount % $UpdateInterval -eq 0) -or ($lastcount -eq $JobErrorFiles.count)) {
                    Write-Host (("`b" * 100) + ('      {0:0000000} @{1}@' -f $lastcount, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                    if ($lastcount -eq $JobErrorFiles.count) { Write-Host }
                }

                $lastCount++
            }

            $x | Sort-Object -Unique | Add-Content -LiteralPath $ErrorFile -Encoding $UTF8Encoding -Force
        } else {
            Write-Host ('    {0:0000000} files to check.' -f $JobResultFiles.count)
        }

        Write-Host "    '$($ErrorFile)'"
    }

    if ($DebugFile) {
        Write-Host "  Combine temporary debug files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $JobDebugFiles = @(Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))

        Write-Host ('    {0:0000000} files to combine.' -f $JobResultFiles.count)
        Write-Host '    Sort and combine will be performed after the step ''End script'' to ensure a complete debug log.'

        Write-Host "    '$($DebugFile)'"
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
                    Get-Content -LiteralPath $JobDebugFile | Add-Content -LiteralPath $DebugFile -Encoding $UTF8Encoding -Force
                }

                Remove-Item -LiteralPath $JobDebugFile -Force
            }
        }
    }

    Remove-Variable * -ErrorAction SilentlyContinue
    [GC]::Collect(); Start-Sleep 1
}
