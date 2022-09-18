<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="../src/logo/Export-RecipientPermissions%20Logo.png" width="450" title="Export-RecipientPermissions" alt="Export-RecipientPermissions"></a>**<br>Document, filter and compare Exchange permissions: Mailbox Access Rights, Mailbox Folder permissions, Public Folder permissions, Send As, Send On Behalf, Managed By, Linked Master Accounts, Forwarders, Group members, Management Role Group members
<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
Document, filter and compare Exchange permissions:
- mailbox access rights,
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

# Table of Contents <!-- omit in toc -->
- [1. Export-RecipientPermissions.ps1](#1-export-recipientpermissionsps1)
  - [1.1. Output](#11-output)
  - [1.2. Parameters](#12-parameters)
    - [1.2.1. ExportFromOnPrem](#121-exportfromonprem)
    - [1.2.2. ExchangeConnectionUriList](#122-exchangeconnectionurilist)
    - [1.2.3. ExchangeOnlineConnectionParameters](#123-exchangeonlineconnectionparameters)
    - [1.2.4. ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential](#124-exchangecredentialusernamefile-exchangecredentialpasswordfile-usedefaultcredential)
    - [1.2.5. ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal](#125-paralleljobsexchange-paralleljobsad-paralleljobslocal)
    - [1.2.6. RecipientProperties](#126-recipientproperties)
    - [1.2.7. GrantorFilter](#127-grantorfilter)
    - [1.2.8. TrusteeFilter](#128-trusteefilter)
    - [1.2.9. ExportFileFilter](#129-exportfilefilter)
    - [1.2.10. ExportMailboxAccessRights](#1210-exportmailboxaccessrights)
    - [1.2.11. ExportMailboxAccessRightsSelf](#1211-exportmailboxaccessrightsself)
    - [1.2.12. ExportMailboxAccessRightsInherited](#1212-exportmailboxaccessrightsinherited)
    - [1.2.13. ExportMailboxFolderPermissions](#1213-exportmailboxfolderpermissions)
    - [1.2.14. ExportMailboxFolderPermissionsAnonymous](#1214-exportmailboxfolderpermissionsanonymous)
    - [1.2.15. ExportMailboxFolderPermissionsDefault](#1215-exportmailboxfolderpermissionsdefault)
    - [1.2.16. ExportMailboxFolderPermissionsOwnerAtLocal](#1216-exportmailboxfolderpermissionsowneratlocal)
    - [1.2.17. ExportMailboxFolderPermissionsMemberAtLocal](#1217-exportmailboxfolderpermissionsmemberatlocal)
    - [1.2.18. ExportMailboxFolderPermissionsExcludeFoldertype](#1218-exportmailboxfolderpermissionsexcludefoldertype)
    - [1.2.19. ExportSendAs](#1219-exportsendas)
    - [1.2.20. ExportSendAsSelf](#1220-exportsendasself)
    - [1.2.21. ExportSendOnBehalf](#1221-exportsendonbehalf)
    - [1.2.22. ExportManagedBy](#1222-exportmanagedby)
    - [1.2.23. ExportLinkedMasterAccount](#1223-exportlinkedmasteraccount)
    - [1.2.24. ExportPublicFolderPermissions](#1224-exportpublicfolderpermissions)
    - [1.2.25. ExportPublicFolderPermissionsAnonymous](#1225-exportpublicfolderpermissionsanonymous)
    - [1.2.26. ExportPublicFolderPermissionsDefault](#1226-exportpublicfolderpermissionsdefault)
    - [1.2.27. ExportPublicFolderPermissionsExcludeFoldertype](#1227-exportpublicfolderpermissionsexcludefoldertype)
    - [1.2.28. ExportSendAs](#1228-exportsendas)
    - [1.2.29. ExportManagementRoleGroupMembers](#1229-exportmanagementrolegroupmembers)
    - [1.2.30. ExportForwarders](#1230-exportforwarders)
    - [1.2.31. ExportDistributionGroupMembers](#1231-exportdistributiongroupmembers)
    - [1.2.32. ExportGroupMembersRecurse](#1232-exportgroupmembersrecurse)
    - [1.2.33. ExportGuids](#1233-exportguids)
    - [1.2.34. ExpandGroups](#1234-expandgroups)
    - [1.2.35. ExportGrantorsWithNoPermissions](#1235-exportgrantorswithnopermissions)
    - [1.2.36. ExportTrustees](#1236-exporttrustees)
    - [1.2.37. ExportFile](#1237-exportfile)
    - [1.2.38. ErrorFile](#1238-errorfile)
    - [1.2.39. DebugFile](#1239-debugfile)
    - [1.2.40. UpdateInverval](#1240-updateinverval)
  - [1.3. Runtime](#13-runtime)
  - [1.4. Requirements](#14-requirements)
- [2. FAQ](#2-faq)
  - [2.1. Which permissions are required?](#21-which-permissions-are-required)
  - [2.2. Can the script resolve permissions granted to a group to it's individual members?](#22-can-the-script-resolve-permissions-granted-to-a-group-to-its-individual-members)
  - [2.3. Where can I find the changelog?](#23-where-can-i-find-the-changelog)
  - [2.4. How can I contribute, propose a new feature or file a bug?](#24-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [2.5. How can I get more script output for troubleshooting?](#25-how-can-i-get-more-script-output-for-troubleshooting)
  - [2.6. A permission is reported, but the trustee details (primary SMTP address etc.) are empty](#26-a-permission-is-reported-but-the-trustee-details-primary-smtp-address-etc-are-empty)
  - [2.7. Isn't a plural noun in the script name against PowerShell best practices?](#27-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
  - [2.8. Is there a roadmap for future versions?](#28-is-there-a-roadmap-for-future-versions)
  - [2.9. Is there a GUI available?](#29-is-there-a-gui-available)
- [3. Sample code](#3-sample-code)
  - [3.1. Get-DependentRecipients.ps1](#31-get-dependentrecipientsps1)
  - [3.2. Compare-RecipientPermissions.ps1](#32-compare-recipientpermissionsps1)
  - [3.3. FiltersAndSidhistory.ps1](#33-filtersandsidhistoryps1)
- [4. Recommendations](#4-recommendations)

# 1. Export-RecipientPermissions.ps1
Finds all recipients with a primary SMTP address in an on on-prem or online Exchange environment and documents their
- mailbox access rights,
- mailbox folder permissions,
- "send as" permissions,
- "send on behalf" permissions, and
- "managed by" permissions
## 1.1. Output
The report is saved to the file 'Export-RecipientPermissions_Result.csv', which consists of the following columns:
- Grantor Primary SMTP: The primary SMTP address of the object granting a permission
  - When management role group members are exported, this column contains 'Management Role Group'
  - When public folder permissions are exported, this column represents the folder's content mailbox
- Grantor Display Name: The display name of the grantor.
  - When management role group members are exported, this column contains the name of the Management Role Group
  - When public folder permissions are exported, this column represents the folder's content mailbox
- Grantor Recipient Type: The recipient type and recipient type detail of the grantor.
  - When management role group members are exported, this column contains 'ManagementRoleGoup'
  - When public folder permissions are exported, this column represents the folder's content mailbox ('UserMailbox/PublicFolderMailbox')
- Grantor Environment: Shows if the grantor is held on-prem or in the cloud.
  - When public folder permissions are exported, this column represents the folder's content mailbox
- Folder: Folder the permission is granted on
  - Empty for non-folder permissions
  - All folder names start with '/', '/' representing the root folder
- Permission: The permission granted/received (e.g., FullAccess, SendAs, SendOnBehalf etc.)
  - When public folder permissions are exported and a folder is mail-enabled, a "virtual" right 'MailEnabled' is exported
  - When management role group members are exported, a "virtual" right 'MemberRecurse' or 'MemberDirect' is exported
  - When forwarders are exported, one or more of the following "virtual" rights are exported:
    - Forward_ExternalEmailAddress_ForwardOnly
    - Forward_ForwardingAddress_DeliverAndForward
    - Forward_ForwardingAddress_ForwardOnly
    - Forward_ForwardingSmtpAddress_DeliverAndForward
    - Forward_ForwardingSmtpAddress_ForwardOnly
- Allow/Deny: Shows if the permission is an allow or a deny permission.
- Inherited: Shows if the permission is inherited or not.
- InheritanceType: Shows if the permission is also valid for child objects, and if yes, which child objects.
- Trustee Original Identity: The original identity string of the trustee.
  - When 'ExpandGroups' is enabled, this column contains the original identity string of the original trustee groups, extended with the string '     [MemberRecurse] ' or '     [MemberDirect] ' and the original identity of the resolved group member
- Trustee Primary SMTP: The primary SMTP address of the object receiving a permission.
  - When 'ExpandGroups' is enabled, the primary SMTP address comes from the resolved group member
- Trustee Display Name: The display name of the trustee.
  - When 'ExpandGroups' is enabled, the display name comes from the resolved group member
- Trustee Recipient Type: The recipient type of the trustee.
-   - When 'ExpandGroups' is enabled, the recipient type comes from the resolved group member
- Trustee Environment: Shows if the trustee is held on-prem or in the cloud.
  - When 'ExpandGroups' is enabled, the trustee environment comes from the resolved group member
## 1.2. Parameters
### 1.2.1. ExportFromOnPrem
Export from On-Prem or from Exchange Online

$true for export from on-prem, $false for export from Exchange Online

Default: $false
### 1.2.2. ExchangeConnectionUriList
Server URIs to connect to

For on-prem installations, list all Exchange Server Remote PowerShell URIs the script can use

For Exchange Online use 'https://outlook.office365.com/powershell-liveid/', or the URI specific to your cloud environment
### 1.2.3. ExchangeOnlineConnectionParameters
This hashtable will be passed as parameter to Connect-ExchangeOnline

Allowed values: AppId, AzureADAuthorizationEndpointUri, BypassMailboxAnchoring, Certificate, CertificateFilePath, CertificatePassword, CertificateThumbprint, Credential, DelegatedOrganization, EnableErrorReporting, ExchangeEnvironmentName, LogDirectoryPath, LogLevel, Organization, PageSize, TrackPerformance, UseMultithreading, UserPrincipalName

Values not in the allow list are removed or replaced with values determined by the script
### 1.2.4. ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential
Credentials for Exchange connection

Username and password are stored as encrypted secure strings, if UseDefaultCredential is not enabled
### 1.2.5. ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal
Maximum Exchange, AD and local sessions/jobs running in parallel.

Watch CPU and RAM usage, and your Exchange throttling policy.

### 1.2.6. RecipientProperties
Recipient properties to import.

Be aware that these properties are not queried with a simple '`Get-Recipient`', but with '`Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Recipient -ResultSize Unlimited | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties)`'.  
This way, some properties have sub-values. For example, the property .PrimarySmtpAddress has .Local, .Domain and .Address as sub-values.

These properties are available for GrantorFilter and TrusteeFilter. 

Properties that are always included: 'Identity', 'DistinguishedName', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'UserFriendlyName', 'LinkedMasterAccount'

### 1.2.7. GrantorFilter
Only check grantors where the filter criteria matches $true.

The variable $Grantor has all attributes defined by '`RecipientProperties`'. For example:
- .DistinguishedName
- .RecipientType.Value, .RecipientTypeDetails.Value
- .DisplayName
- .PrimarySmtpAddress: .Local, .Domain, .Address
- .EmailAddresses: .PrefixString, .IsPrimaryAddress, .SmtpAddress, .ProxyAddressString  
  This attribute is an array. Code example:
    ```
    $GrantorFilter = "if ((`$Grantor.EmailAddresses.SmtpAddress -ilike 'AddressA@example.com') -or (`$Grantor.EmailAddresses.SmtpAddress -ilike 'Test*@example.com')) { `$true } else { `$false }"
    ```
- .UserFriendlyName: User account holding the mailbox in the '`<NetBIOS domain name>\<sAMAccountName>`' format
- .ManagedBy: .Rdn, .Parent, .DistinguishedName, .DomainId, .Name
    This attribute is an array. Code example:
    ```
    $GrantorFilter = "foreach (`$XXXSingleManagedByXXX in `$Grantor.ManagedBy) { if (`$XXXSingleManagedByXXX -iin @(
                        'example.com/OU1/OU2/ObjectA',
                        'example.com/OU3/OU4/ObjectB',
    )) { `$true; break } }"
    ```
  On-prem only:
    .Identity: .tostring() (CN), .DomainId, .Parent (parent CN)
    .LinkedMasterAccount: Linked Master Account in the '`<NetBIOS domain name>\<sAMAccountName>`' format

Set to \$null or '' to define all recipients as grantors to consider

Example:
```
"`$Grantor.primarysmtpaddress.domain -ieq 'example.com'"
```

Default: $null
### 1.2.8. TrusteeFilter
Only report trustees where the filter criteria matches $true.

If the trustee matches a recipient, the available attributes are the same as for GrantorFilter, only the reference variable is $Trustee instead of $Grantor.

If the trustee does not match a recipient (because it no longer exists, for exampe), $Trustee is just a string. In this case, the export shows the following:
- Column "Trustee Original Identity" contains the trustee description string as reported by Exchange
- Columns "Trustee Primary SMTP" and "Trustee Display Name" are empty

Example:
```
"`$Trustee.primarysmtpaddress.domain -ieq 'example.com'"
```

Default: $null
### 1.2.9. ExportFileFilter
Only report results where the filter criteria matches $true.

This filter works against every single row of the results found. ExportFile will only contain lines where this filter returns $true.

The $ExportFileLine contains an object with the header names from $ExportFile as string properties:
- 'Grantor Primary SMTP'
- 'Grantor Display Name'
- 'Grantor Exchange GUID' (only when '`ExportGuids`' is enabled)
- 'Grantor AD ObjectGUID' (only when '`ExportGuids`' is enabled)
- 'Grantor Recipient Type'
- 'Grantor Environment'
- 'Folder'
- 'Permission'
- 'Allow/Deny'
- 'Inherited'
- 'InheritanceType'
- 'Trustee Original Identity'
- 'Trustee Primary SMTP'
- 'Trustee Display Name'
- 'Trustee Exchange GUID' (only when '`ExportGuids`' is enabled)
- 'Trustee AD ObjectGUID' (only when '`ExportGuids`' is enabled)
- 'Trustee Recipient Type'
- 'Trustee Environment'

Example: "`$ExportFileFilter.'Trustee Environment' -ieq 'On-Prem'"

Default: $null
### 1.2.10. ExportMailboxAccessRights
Rights set on the mailbox itself, such as "FullAccess" and "ReadAccess"

Default: $true
### 1.2.11. ExportMailboxAccessRightsSelf
Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)

Default: $false
### 1.2.12. ExportMailboxAccessRightsInherited
Report inherited mailbox access rights (only works on-prem)

Default: $false
### 1.2.13. ExportMailboxFolderPermissions
This part of the report can take very long

Default: $false
### 1.2.14. ExportMailboxFolderPermissionsAnonymous
Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)

Default: $true
### 1.2.15. ExportMailboxFolderPermissionsDefault
Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)

Default: $true
### 1.2.16. ExportMailboxFolderPermissionsOwnerAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.

Default: $false
### 1.2.17. ExportMailboxFolderPermissionsMemberAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.
Default: $false
### 1.2.18. ExportMailboxFolderPermissionsExcludeFoldertype
List of Foldertypes to ignore.

Some known folder types are: Audits, Calendar, CalendarLogging, CommunicatorHistory, Conflicts, Contacts, ConversationActions, DeletedItems, Drafts, ExternalContacts, Files, GalContacts, ImContactList, Inbox, Journal, JunkEmail, LocalFailures, Notes, Outbox, QuickContacts, RecipientCache, RecoverableItemsDeletions, RecoverableItemsPurges, RecoverableItemsRoot, RecoverableItemsVersions, Root, RssSubscription, SentItems, ServerFailures, SyncIssues, Tasks, WorkingSet, YammerFeeds, YammerInbound, YammerOutbound, YammerRoot

Default: 'audits'
### 1.2.19. ExportSendAs
Export Send As permissions

Default: $true
### 1.2.20. ExportSendAsSelf
Export Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)

Default: $false
### 1.2.21. ExportSendOnBehalf
Export Send On Behalf permissions

Default: $true
### 1.2.22. ExportManagedBy
Only for distribution groups, and not to be confused with the "Manager" attribute

Default: $true
### 1.2.23. ExportLinkedMasterAccount
Export Linked Master Account

Only works on-prem

Default: $true
### 1.2.24. ExportPublicFolderPermissions
Export public folder permissions

This part of the report can take very long

GrantorFilter refers to the public folder content mailbox

Default: $true
### 1.2.25. ExportPublicFolderPermissionsAnonymous
Report public folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)

Default: $true
### 1.2.26. ExportPublicFolderPermissionsDefault
Report public folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)

Default: $true
### 1.2.27. ExportPublicFolderPermissionsExcludeFoldertype
List of Foldertypes to ignore.

Some known folder types are: IPF.Appointment, IPF.Contact, IPF.Note, IPF.Task

Default: ''
### 1.2.28. ExportSendAs
Export Send As permissions

Default: $true
### 1.2.29. ExportManagementRoleGroupMembers
Export members of management role groups

The virtual right 'MemberRecurse' or 'MemberDirect' is used in the export file

GrantorFilter does not apply to the export of management role groups, but TrusteeFilter and ExportFileFilter do

Default: $true
### 1.2.30. ExportForwarders
Export forwarders:
- '`ExternalEmailAddress`' ('`targetAddress`' in Active Directory)
  - Highest priority
  - Can be configured on basically every mail-enabled object
  - Can point to any SMTP address, existing in your directory or somewhere outside
  - Is typically used for contacts, migration and co-existence scenarios
  - '`DeliverToMailboxAndForward`' ('`deliverAndRedirect`' in Active Directory) is ignored, all e-mails sent to the recipient will unconditionally be forwarded without storing a copy or sending it to group members
- '`ForwardingAddress`' ('`altRecipient`' in Active Directory)
  - Medium priority
  - Can be configured for mailboxes and mail-enabled public folders
  - Needs a mail-enabled Object existing in your directory as target (a contact is required to forward to external e-mail addresses)
  - This property is used when forwarding is configured in the Exchange Control Panel oder the Exchange Admin Center
  - '`DeliverToMailboxAndForward`' ('`deliverAndRedirect`' in Active Directory) defines if the e-mail is forwarded only, or forwarded and stored
- '`ForwardingSmtpAddress`' ('`msExchGenericForwardingAddress`' in Active Directory)
  - Lowest priority
  - Can be configured for mailboxes
  - Can point to any SMTP address, existing in your directory or somewhere outside
  - This property is used when a user configures forwarding for his mailbox in Outlook Web
  - '`DeliverToMailboxAndForward`' ('`deliverAndRedirect`' in Active Directory) defines if the e-mail is forwarded only, or forwarded and stored

When forwarders are exported, one or more of the following "virtual" rights are exported:
- Forward_ExternalEmailAddress_ForwardOnly
- Forward_ForwardingAddress_DeliverAndForward
- Forward_ForwardingAddress_ForwardOnly
- Forward_ForwardingSmtpAddress_DeliverAndForward
- Forward_ForwardingSmtpAddress_ForwardOnly

Default: $true
### 1.2.31. ExportDistributionGroupMembers
Export distribution group members, including nested groups and dynamic groups

The parameter ExpandGroups can be used independently:
  ExpandGroups acts when a group is used as trustee: It adds every recurse member of the group as a separate trustee entry
  ExportDistributionGroupMembers exports the distribution group as grantor, which the recurse members as trustees

Valid values: 'None', 'All', 'OnlyTrustees'
  'None': Distribution group members are not exported Parameter ExpandGroups can still be used.
  'All': Members of all distribution groups are exported, parameter GrantorFilter is considerd
  'OnlyTrustees': Only export members of those distribution groups that are used as trustees, even when they are excluded via GrantorFilter

Default: 'None'
### 1.2.32. ExportGroupMembersRecurse
When disabled, only direct members of groups are exported, and the virtual right 'MemberDirect' is used in the export file.

When enabled, recursive members of groups are exported, and the virtual right 'MemberRecurse' is used in the export file.

Default: $false
### 1.2.33. ExportGuids
When enabled, the export contains the Exchange GUID and the AD ObjectGUID for each grantor and trustee

Default: $false
### 1.2.34. ExpandGroups
Expand groups to their recursive members, including nested groups and dynamic groups

This is useful in cases where users are sent permission reports, as not only permission changes but also changes in the underlying trustee groups are documented and directly associated with a grantor-permission-trustee triplet.  
For example: User A has granted Group B permission X a long time ago. The permission itself does not change, but the recursive members of Group B change over time. With ExpandGroups enabled, the members and therefore the changes of Group B are documented with every run of Export-RecipientPermissions.

This may drastically increase script run time and file size

The original permission is still documented, with one additional line for each member of the group
- For each member of the group, 'Trustee Original Identity' is preserved but suffixed with
  ```
       [MemberRecurse] 
  ```
  or
  ```
       [MemberDirect] 
  ```
  The whitespace consists of five space characters in front of 'MemberRecurse'/'MemberDirect' for sorting reasons, and one space at the end. Then the original identity string of the resolved group member is added.
  The other trustee properties are the ones of the member

TrusteeFilter is applied to trustee groups as well as to their finally expanded individual members
- Nested groups are expanded to individual members, but TrusteeFilter is not applied to the nested group

Default value: $false
### 1.2.35. ExportGrantorsWithNoPermissions
Per default, Export-RecipientPermissions only exports grantors which have set at least one permission for at least one trustee.
If all grantors should be exported, set this parameter to $true.

If enabled, a grantor that that not grant any permission is included in the list with the following columns: "Grantor Primary SMTP", "Grantor Display Name", "Grantor Recipient Type", "Grantor Environment". The other columns for this recipient are empty.

Default value: $false
### 1.2.36. ExportTrustees
Include all trustees in permission report file, only valid or only invalid ones

Valid trustees are trustees which can be resolved to an Exchange recipient

Valid values: 'All', 'OnlyValid', 'OnlyInvalid'

Default: 'All'
### 1.2.37. ExportFile
Name (and path) of the permission report file

Default: '.\export\Export-RecipientPermissions_Result.csv'
### 1.2.38. ErrorFile
Name (and path) of the error log file

Set to $null or '' to disable debugging

Default: '.\export\Export-RecipientPermissions_Error.csv',
### 1.2.39. DebugFile
Name (and path) of the debug log file

Set to $null or '' to disable debugging

Default: ''
### 1.2.40. UpdateInverval
Interval to update the job progress

Updates are based von recipients done, not on duration

Number must be 1 or higher, lower numbers mean bigger debug files

Default: 100
## 1.3. Runtime
The script can run many hours, depending on the number of recipients and the speed of the environments to check.

Exporting mailbox folder permissions takes even more time because of how Exchange is designed to query these permissions.
## 1.4. Requirements
The script needs to be run with an account that has read permissions to all recipients in the cloud as well as Active Directory and Exchange on premises. The script asks for credentials.

As the credentials are stored in the encrypted secure string file format and can be re-used, the script can be fully automated and run as a scheduled job.

Per default, the script uses multiple parallel threads, each one consuming one Exchange PowerShell session. Please watch CPU and RAM usage, as well as your Exchange throttling policy:
```
(Get-ThrottlingPolicyAssociation -Identity ([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name) | foreach {
	"THROTTLING POLICY ASSOCIATION"
	$_
	"THROTTLING POLICY DETAILS"
	$_.throttlingpolicyid | Get-ThrottlingPolicy
}
```
# 2. FAQ
## 2.1. Which permissions are required?
Export-RecipientPermissions uses the following Exchange PowerShell cmdlets:
- '`Get-DistributionGroup`'
- '`Get-DistributionGroupMember`'
- '`Get-DynamicDistributionGroup`'
- '`Get-DynamicDistributionGroupMember`' (this cmdlet is only available in Exchange Online)
- '`Get-Mailbox`'
- '`Get-MailboxDatabase`' (this cmdlet is only used on premises)
- '`Get-MailboxFolderPermission`'
- '`Get-MailboxFolderStatistics`'
- '`Get-MailboxPermission`'
- '`Get-MailPublicFolder`'
- '`Get-Publicfolder`'
- '`Get-PublicFolderClientPermission`'
- '`Get-Recipient`'
- '`Get-RecipientPermission`'
- '`Get-RoleGroup`'
- '`Get-RoleGroupMember`'
- '`Get-SecurityPrincipal`'
- '`Get-UnifiedGroup`' (this cmdlet is only available in Exchange Online)
- '`Get-UnifiedGroupLinks`' (this cmdlet is only available in Exchange Online)

In on-premises environments, membership in the Exchange management role group 'View-Only Organization Management' is sufficient.

In Exchange Online, the Exchange management role group 'View-Only Organization Management' (which contains the Azure AD role group 'Global Reader' per default) is not sufficient, as - for an unknown reason - the cmdlets '`Get-RecipientPermission`' and '`Get-SecurityPrincipal`' are not included in this management role group.
- '`Get-RecipientPermission`' is included in the role groups '`Organization Management`' and '`Recipient Management`'
- '`Get-SecurityPrincipal`' is included in the role group '`Organization Management`'.  

You can use the following script to find out which cmdlet is assisgned to which management role:
```
$ExportFile = '.\Required Cmdlets and their management role assignment.csv'

$Cmdlets = (
    'Get-DistributionGroup',
    'Get-DistributionGroupMember',
    'Get-DynamicDistributionGroup',
    'Get-DynamicDistributionGroupMember', # this cmdlet is only available in Exchange Online
    'Get-Mailbox',
    'Get-MailboxDatabase', # this cmdlet is only used on premises
    'Get-MailboxFolderPermission',
    'Get-MailboxFolderStatistics',
    'Get-MailboxPermission',
    'Get-MailPublicFolder',
    'Get-Publicfolder',
    'Get-PublicFolderClientPermission',
    'Get-Recipient',
    'Get-RecipientPermission',
    'Get-RoleGroup',
    'Get-RoleGroupMember',
    'Get-SecurityPrincipal',
    'Get-UnifiedGroup', # this cmdlet is only available in Exchange Online
    'Get-UnifiedGroupLinks' # this cmdlet is only available in Exchange Online
)


if ($PSVersionTable.PSEdition -ieq 'desktop') {
    $UTF8Encoding = 'UTF8'
} else {
    $UTF8Encoding = 'UTF8BOM'
}


Write-Host 'Get management role assignment per cmdlet'

$ResultTable = New-Object System.Data.DataTable 'ResultTable'
$null = $ResultTable.Columns.Add('Cmdlet')

foreach ($Cmdlet in @($Cmdlets | Sort-Object -Unique)) {
    Write-Host "  $($cmdlet)"
    $TempRoleAssigneeNames = @()

    foreach ($CmdletPerm in (Get-ManagementRole -Cmdlet $Cmdlet)) {
        foreach ($ManagementRoleAssignment in @(Get-ManagementRoleAssignment -Role $CmdletPerm.Name -Delegating $false | Select-Object RoleAssigneeType, RoleAssigneeName)) {
            $TempRoleAssigneeNames += "$($ManagementRoleAssignment.RoleAssigneeType): $($ManagementRoleAssignment.RoleAssigneeName)"
        }
    }

    foreach ($TempRoleAssigneeName in @($TempRoleAssigneeNames | Where-Object { $_ } | Sort-Object -Unique)) {
        if ($TempRoleAssigneeName -notin $ResultTable.Columns.ColumnName) {
            $null = $ResultTable.Columns.Add($TempRoleAssigneeName)
        }
    }

    $CmdletRow = $ResultTable.NewRow()
    $CmdletRow.'Cmdlet' = $Cmdlet

    foreach ($TempRoleAssigneeName in @($TempRoleAssigneeNames | Where-Object { $_ })) {
        $CmdletRow."$TempRoleAssigneeName" = $true
    }

    $ResultTable.Rows.Add($CmdletRow)
}


Write-Host
Write-Host 'Create export file'
Write-Host "  '$ExportFile'"

$ResultTable | Select-Object @(@('Cmdlet') + @($ResultTable.Columns.ColumnName | Where-Object { $_ -ne 'Cmdlet' } | Sort-Object -Unique)) | Export-Csv -Path $ExportFile -Force -Encoding $UTF8Encoding -Delimiter ';' -NoTypeInformation


Write-Host
Write-Host 'Done'
```

In both environments, a tailored custom management role group with the required permissions and recipient restrictions can be created.
## 2.2. Can the script resolve permissions granted to a group to it's individual members?
Yes, Export-RecipientPermissions can resolve trustee groups to their individual members. Use the parameter `'ExpandGroups'` to enable this feature.
## 2.3. Where can I find the changelog?
The changelog is located in the `'.\docs'` folder, along with other documents related to Export-RecipientPermissions.
## 2.4. How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, please have a look at `'.\docs\CONTRIBUTING'` for a rough overview of the proposed process.
## 2.5. How can I get more script output for troubleshooting?
Start the script with the '-verbose' parameter to get the maximum output for troubleshooting.
## 2.6. A permission is reported, but the trustee details (primary SMTP address etc.) are empty
When excluding a bug in the script, there are three possible reasons why a trustee does not have details like a primary SMTP address in the result file:
- The trustee is a valid Active Directory object, but not a valid Exchange recipient.  
Examples: 'NT AUTHORITY\SELF', 'Domain Admins', 'Enterprise Admins'
- The initial trustee no longer exists.  
Exchange does not check if trustees still exist and remove the according permissions in case of deletion - this would be a problem when restoring deleted Active Directory objects.  
As Exchange stores trustees in different formats, the trustee original identity can be a SID, an NT4-style logon name or just about any string.
- Multiple recipients share the same linked master account, user friendly name, distinguished name, GUID or primary SMTP address. As the search for this value returns multiple recipients, no details are shown.  
This should not happen when using the built-in Exchange tools, due to their built-in quality checks. It happens more often, when Exchange attributes are modified directly in Active Directory.  
Whe passing the '-verbose' PowerShell parameter, the script outputs recipients with non-unique attributes in the verbose stream.  
## 2.7. Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Export-RecipientPermissions contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.
## 2.8. Is there a roadmap for future versions?
There is no binding roadmap for future versions, although I maintain a list of ideas in the 'Contribution opportunities' chapter of '.\docs\CONTRIBUTING.html'.

Fixing issues has priority over new features, of course.
## 2.9. Is there a GUI available?
There is no dedicated graphical user interface.

A basic GUI for configuring the script is accessible via the following built-in PowerShell command:
```
Show-Command .\Export-RecipientPermissions.ps1
```
# 3. Sample code
## 3.1. Get-DependentRecipients.ps1
The script can be found in '`.\sample code\Get-DependentRecipients`'.

Currently only some recipient permissions work cross-premises according to Microsoft. All other permissions, including the one to manage the members of distribution lists, only work when both the grantor and the trustee are hosted on the same environment.
There are environments where permissions work cross-premises, but there is no offical support from Microsoft.

This script takes a list of recipients and the output of Export-RecipientPermissions.ps1 to create a list of all recipients groups that have a grantor-trustee dependency beyond "full access" to each other.

The script not only considers situations where recipient A grants rights to recipient B, but the whole permission chain ("X-Z-A-B-C-D" etc.).

The script does not consider group membership.

The following outputs are created:
- Export-RecipientPermissions_Output_Permissions.csv  
  The original permission input file, reduced to the rows that have a connection with the recipient input file.  
  Enhanced with information if a grantor or trustee is part of the initial recipient file or has to be migrated additionally to keep permission chains working.
  Enhanced with information which single permissions start permissions chains outside the initial recipients.
- Get-DependentRecipients_Output_InitialRecipients.csv  
  List of initial recipients.
- Get-DependentRecipients_Output_AdditionalRecipients.csv  
  List of additional recipients.
- Get-DependentRecipients_Output_AllRecipients.csv  
  List of all initial and additional recipients.
- Get-DependentRecipients_Output_GML.gml  
  All recipients and their permissions in a graphical representation. The GML (Graph Modeling Language) file format used is human readable. Free tools like yWorks yEd Graph Editor, Gephi and others can be used to easily create visual representations from this file.  
- Get-DependentRecipients_Output_Summary.txt  
  Number of initial recipients, number of additional recipients, number of total recipients, number of root cause mailbox permissions.
## 3.2. Compare-RecipientPermissions.ps1
The script can be found in '`.\sample code\Compare-RecipientPermissions`'.

Compare two result files from Export-RecipientPermissions.ps1 to see which permissions have changed over time

Changes are marked in the column 'Change' with
- 'Deleted' if a line exists in the old file but not in the new one
- 'New' if a line exists in the new file but not in the old one
- 'Unchanged' if a line exists as well in the new file as in the old one
## 3.3. FiltersAndSidhistory.ps1
The script can be found in '`.\sample code\other samples`'.

This sample code shows how to use TrusteeFilter to find permissions which may be affected by SIDHistory removal.

GrantorFilter behaves exactly like TrusteeFilter, only the reference variable is $Grantor instead of $Trustee.
# 4. Recommendations
Make sure you have the latest updates installed to avoid memory leaks and CPU spikes (PowerShell, .Net framework).

If possible, allow Export-RecipientPermissions.ps1 to use your on premises infrastructure. This will dramatically increase the initial enumeration of recipients.

Start the script from PowerShell, not from within the PowerShell ISE. This makes especially Get-DependentMailboxes.ps1 run faster due to a different default thread apartment mode.

When running the scripts as scheduled job, make sure to include the "-ExecutionPolicy Bypass" parameter.
Example: `powershell.exe -ExecutionPolicy Bypass -file "c:\path\Export-RecipientPermissions.ps1"`
