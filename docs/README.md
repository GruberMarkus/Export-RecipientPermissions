<!-- omit in toc -->
## **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="../src/logo/Export-RecipientPermissions%20Logo.png" width="400" title="Export-RecipientPermissions" alt="Export-RecipientPermissions"></a>**<br>Document, filter and compare Exchange permissions<br><br><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://explicitconsulting.at/open-source/export-recipientpermissions/" target="_blank"><img src="https://img.shields.io/badge/get%20fee--based%20support%20from-ExplicIT%20Consulting-lawngreen?labelColor=deepskyblue" alt="get fee-based support from ExplicIT Consulting"></a>

# Features <!-- omit in toc -->
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
    - [1.2.31. ExportModerators](#1231-exportmoderators)
    - [1.2.32. ExportRequireAllSendersAreAuthenticated](#1232-exportrequireallsendersareauthenticated)
    - [1.2.33. ExportAcceptMessagesOnlyFrom](#1233-exportacceptmessagesonlyfrom)
    - [1.2.34. ExportResourceDelegates](#1234-exportresourcedelegates)
    - [1.2.35. ExportDistributionGroupMembers](#1235-exportdistributiongroupmembers)
    - [1.2.36. ExportGroupMembersRecurse](#1236-exportgroupmembersrecurse)
    - [1.2.37. ExportGuids](#1237-exportguids)
    - [1.2.38. ExpandGroups](#1238-expandgroups)
    - [1.2.39. ExportGrantorsWithNoPermissions](#1239-exportgrantorswithnopermissions)
    - [1.2.40. ExportTrustees](#1240-exporttrustees)
    - [1.2.41. ExportFile](#1241-exportfile)
    - [1.2.42. ErrorFile](#1242-errorfile)
    - [1.2.43. DebugFile](#1243-debugfile)
    - [1.2.44. UpdateInverval](#1244-updateinverval)
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
  - [2.10. Which resources does a particular user or group have access to?](#210-which-resources-does-a-particular-user-or-group-have-access-to)
  - [2.11. How to find distribution lists without members?](#211-how-to-find-distribution-lists-without-members)
    - [2.11.1. How to export permissions for specific public folders?](#2111-how-to-export-permissions-for-specific-public-folders)
    - [2.11.2. I receive an error message when connecting to Exchange on premises](#2112-i-receive-an-error-message-when-connecting-to-exchange-on-premises)
- [3. Sample code](#3-sample-code)
  - [3.1. Get-DependentRecipients.ps1](#31-get-dependentrecipientsps1)
  - [3.2. Compare-RecipientPermissions.ps1](#32-compare-recipientpermissionsps1)
  - [3.3. FiltersAndSidhistory.ps1](#33-filtersandsidhistoryps1)
  - [3.4. MemberOfRecurse.ps1](#34-memberofrecurseps1)
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
Exchange remote PowerShell URIs to connect to

For on-prem installations, list all Exchange Server Remote PowerShell URIs the script can use
For Exchange Online, use 'https://outlook.office365.com/powershell-liveid/' or the URI specific to your cloud environment

Default:  
- If ExportFromOnPrem ist set to false: 'https://outlook.office365.com/powershell-liveid/'
- If ExportFromOnPrem ist set to true: 'http://\<server\>/powershell' for each Exchange server with the mailbox server role
### 1.2.3. ExchangeOnlineConnectionParameters
This hashtable will be passed as parameter to Connect-ExchangeOnline

All values are allowed, but CommandName and ConnectionUri are set by the script. By default, ShowBanner and ShowProgress are set to $false; SkipLoadingFormatData to $true.
### 1.2.4. ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential
Credentials for Exchange connection

Username and password are stored as encrypted secure strings, if UseDefaultCredential is not enabled
### 1.2.5. ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal
Maximum Exchange, AD and local sessions/jobs running in parallel.

Watch CPU and RAM usage, and your Exchange throttling policy. Frequent connection errors indicate that the values are set too high.

Default values:
- ParallelJobsExchange: $ExchangeConnectionUriList.count
- ParallelJobsAD: 50
- ParallelJobsLocal: 50


### 1.2.6. RecipientProperties
Recipient properties to import.

Be aware that these properties are not queried with '`Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Recipient -ResultSize Unlimited | Select-Object -Property $args[0] } -ArgumentList @(, $RecipientProperties)`', but with a simple '`Get-Recipient`'.

These properties are available for GrantorFilter and TrusteeFilter.

Properties that are always included: 'Identity', 'DistinguishedName', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'Name', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'UserFriendlyName', 'LinkedMasterAccount'

### 1.2.7. GrantorFilter
Only check grantors where the filter criteria matches $true.

The variable $Grantor has all attributes defined by '`RecipientProperties`'. For example:
- .DistinguishedName
- .RecipientType, .RecipientTypeDetails
- .DisplayName
- .Identity
- .PrimarySmtpAddress
- .EmailAddresses  
  This attribute is an array. Code example:
    ```
    $GrantorFilter = "
        if (
            (`$Grantor.EmailAddresses -ilike 'smtp:AddressA@example.com') -or
            (`$Grantor.EmailAddresses -ilike 'smtp:Test*@example.com')
        ) {
            `$true
        } else {
            `$false
        }
    "
    ```
- .UserFriendlyName: User account holding the mailbox in the '`<NetBIOS domain name>\<sAMAccountName>`' format
- .ManagedBy
    This attribute is an array. Code example:
    ```
    $GrantorFilter = "
        foreach (
            `$XXXSingleManagedByXXX in `$Grantor.ManagedBy
        ) {
            if (
                `$XXXSingleManagedByXXX -iin @(
                    'example.com/OU1/OU2/ObjectA',
                    'example.com/OU3/OU4/ObjectB',
                )
            ) {
                `$true; break
            }
        }
    "
    ```
  On-prem only:
    .LinkedMasterAccount: Linked Master Account in the '`<NetBIOS domain name>\<sAMAccountName>`' format

Set to \$null or '' to define all recipients as grantors to consider

Example:
```
"`$Grantor.primarysmtpaddress -ilike '*@example.com'"
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
"`$Trustee.primarysmtpaddress -ieq '*@example.com'"
```

Default: $null
### 1.2.9. ExportFileFilter
Only report results where the filter criteria matches $true.

This filter works against every single row of the results found. ExportFile will only contain lines where this filter returns $true.

The $ExportFileLine variable contains an object with the header names from $ExportFile as string properties
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

Example: "`$ExportFileLine.'Trustee Environment' -ieq 'On-Prem'"

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
  - This property is used when forwarding is configured in the Exchange Control Panel or the Exchange Admin Center
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
### 1.2.31. ExportModerators
Exports the virtual rights 'ModeratedBy' and 'ModeratedByBypass', listing all users and groups which are configured as moderators for a recipient or can bypass moderation.

Only works for recipients with moderation enabled.

Default: $true
### 1.2.32. ExportRequireAllSendersAreAuthenticated
Exports the virtual right 'RequireAllSendersAreAuthenticated' with the trustee 'NT AUTHORITY\Authenticated Users' for each recipient which is configured to only receive messages from authenticated (internal) senders.

Default: $true

### 1.2.33. ExportAcceptMessagesOnlyFrom
Exports the virtual right 'AcceptMessagesOnlyFrom' for each recipient which is configured to only receive messages from selected (internal) senders.

The attributes 'AcceptMessagesOnlyFrom' and 'AcceptMessagesOnlyFromDLMembers' are exported as the same virtual right 'AcceptMessagesOnlyFrom'.

Default: $true

### 1.2.34. ExportResourceDelegates
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
### 1.2.35. ExportDistributionGroupMembers
Export distribution group members, including nested groups and dynamic groups

The parameter ExpandGroups can be used independently:
  ExpandGroups acts when a group is used as trustee: It adds every recurse member of the group as a separate trustee entry
  ExportDistributionGroupMembers exports the distribution group as grantor, which the recurse members as trustees

Valid values: 'None', 'All', 'OnlyTrustees'
  'None': Distribution group members are not exported Parameter ExpandGroups can still be used.
  'All': Members of all distribution groups are exported, parameter GrantorFilter is considerd
  'OnlyTrustees': Only export members of those distribution groups that are used as trustees, even when they are excluded via GrantorFilter

Default: 'None'
### 1.2.36. ExportGroupMembersRecurse
When disabled, only direct members of groups are exported, and the virtual right 'MemberDirect' is used in the export file.

When enabled, recursive members of groups are exported, and the virtual right 'MemberRecurse' is used in the export file.

Default: $false
### 1.2.37. ExportGuids
When enabled, the export contains the Exchange GUID and the AD ObjectGUID for each grantor and trustee

Default: $false
### 1.2.38. ExpandGroups
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
### 1.2.39. ExportGrantorsWithNoPermissions
Per default, Export-RecipientPermissions only exports grantors which have set at least one permission for at least one trustee.
If all grantors should be exported, set this parameter to $true.

If enabled, a grantor that that not grant any permission is included in the list with the following columns: "Grantor Primary SMTP", "Grantor Display Name", "Grantor Recipient Type", "Grantor Environment". The other columns for this recipient are empty.

Default value: $false
### 1.2.40. ExportTrustees
Include all trustees in permission report file, only valid or only invalid ones

Valid trustees are trustees which can be resolved to an Exchange recipient

Valid values: 'All', 'OnlyValid', 'OnlyInvalid'

Default: 'All'
### 1.2.41. ExportFile
Name (and path) of the permission report file

Default: '.\export\Export-RecipientPermissions_Result.csv'
### 1.2.42. ErrorFile
Name (and path) of the error log file

Set to $null or '' to disable debugging

Default: '.\export\Export-RecipientPermissions_Error.csv',
### 1.2.43. DebugFile
Name (and path) of the debug log file

Set to $null or '' to disable debugging

Default: ''
### 1.2.44. UpdateInverval
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

Windows Powershell 5.1 is supported.  
For best results, especially when Export-RecipientPermissions is used with Exchange Online, PowerShell 7 (or newer) is recommended.
# 2. FAQ
## 2.1. Which permissions are required?
Export-RecipientPermissions uses the following Exchange PowerShell cmdlets:
- 'Get-CASMailbox',
- 'Get-CalendarProcessing',
- 'Get-DistributionGroup',
- 'Get-DynamicDistributionGroup',
- 'Get-DynamicDistributionGroupMember', # Exchange Online only
- 'Get-EXOMailbox', # Exchange Online only
- 'Get-EXOMailboxFolderPermission', # Exchange Online only
- 'Get-EXOMailboxFolderStatistics', # Exchange Online only
- 'Get-EXOMailboxPermission', # Exchange Online only
- 'Get-EXORecipient', # Exchange Online only
- 'Get-EXORecipientPermission', # Exchange Online only
- 'Get-Group',
- 'Get-LinkedUser',
- 'Get-Mailbox',
- 'Get-MailboxDatabase', # Exchange on-prem only
- 'Get-MailboxFolderPermission',
- 'Get-MailboxFolderStatistics',
- 'Get-MailboxPermission',
- 'Get-MailContact',
- 'Get-MailPublicFolder',
- 'Get-MailUser',
- 'Get-Publicfolder',
- 'Get-PublicFolderClientPermission',
- 'Get-Recipient',
- 'Get-RecipientPermission',
- 'Get-RemoteMailbox', # Exchange on-prem only
- 'Get-SecurityPrincipal',
- 'Get-UMMailbox',
- 'Get-UnifiedGroup', # Exchange Online only
- 'Get-UnifiedGroupLinks', # Exchange Online only
- 'Get-User',
- 'Set-AdServerSettings' # Exchange on-prem only


In on-premises environments, membership in the Exchange management role group 'View-Only Organization Management' is sufficient.

In Exchange Online, the Exchange management role group 'View-Only Organization Management' (which contains the Azure AD role group 'Global Reader' per default) is not sufficient, as - for an unknown reason - the cmdlets '`Get-RecipientPermission`' and '`Get-SecurityPrincipal`' are not included in this management role group.
- '`Get-RecipientPermission`' is included in the role groups '`Organization Management`' and '`Recipient Management`'
- '`Get-SecurityPrincipal`' is included in the role group '`Organization Management`'.  

You can use the following script to find out which cmdlet is assisgned to which management role:
```
$ExportFile = '.\Required Cmdlets and their management role assignment.csv'

$Cmdlets = (
        'Get-CASMailbox',
        'Get-CalendarProcessing',
        'Get-DistributionGroup',
        'Get-DynamicDistributionGroup',
        'Get-DynamicDistributionGroupMember', # Exchange Online only
        'Get-EXOMailbox', # Exchange Online only
        'Get-EXOMailboxFolderPermission', # Exchange Online only
        'Get-EXOMailboxFolderStatistics', # Exchange Online only
        'Get-EXOMailboxPermission', # Exchange Online only
        'Get-EXORecipient', # Exchange Online only
        'Get-EXORecipientPermission', # Exchange Online only
        'Get-Group',
        'Get-LinkedUser',
        'Get-Mailbox',
        'Get-MailboxDatabase', # Exchange on-prem only
        'Get-MailboxFolderPermission',
        'Get-MailboxFolderStatistics',
        'Get-MailboxPermission',
        'Get-MailContact',
        'Get-MailPublicFolder',
        'Get-MailUser',
        'Get-Publicfolder',
        'Get-PublicFolderClientPermission',
        'Get-Recipient',
        'Get-RecipientPermission',
        'Get-RemoteMailbox', # Exchange on-prem only
        'Get-SecurityPrincipal',
        'Get-UMMailbox',
        'Get-UnifiedGroup', # Exchange Online only
        'Get-UnifiedGroupLinks', # Exchange Online only
        'Get-User',
        'Set-AdServerSettings' # Exchange on-prem only
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
## 2.10. Which resources does a particular user or group have access to?
As many other IT systems, Exchange stores permissions as forward links and not as backward links. This means that the information about granted permissions is stored on the granting object (forwarding to the trustee), without storing any information about the granted permission at the trustee object (pointing backwards to the grantor).

This means that we have to query all granted permissions first and then filter those for trustees that involve the user or group we are looking for.

There are some exceptions where Exchange stores permissions not only as forward links but also as backward links, to allow getting a list of certain permission with just one fast query from the grantor perspective. All these cases rely on automatic calculation of the backlink attribute in Active Directory. Well-known examples are publicDelegates/publicDelegatesBL, member/memberOf, manager/directReports and owner/ownerBL.

But back to the initial question which resources a particular user or group has access to.  
We already know that we need to query all permissions of interest first, and then filter the results. But what should we filter for?

If we are only interested in permissions granted directly to a certain user or group, the search is straight forward:
- If the user or group is an Exchange recipient, use '`TrusteeFilter`' to filter for 'Trustee Primary SMTP'
- If The user or group is not an Exchange recipient, enable '`ExportGUIDs`' and use '`ExportFileFilter`' to filter for '`Trustee AD ObjectGuid`'

If we are interested in permissions granted directly or indirectly to a certain user or group, it get's more complicated.  
Export-RecipientPermissions can resolve permissions granted to groups in three ways: Do not resolve groups, resolve groups to their direct members, or resolve groups to their recurse members.
- Not resolving groups does not take into consideration nested groups (permissions granted indirectly)
- Resolving groups also does not consider nested groups (permissions granted indirectly) below the first membership level
- Resolving groups to their recurse members requires relatively high CPU and RAM ressources and results in large result files.
  - '`ExportDistributionGroupMembers`' only helps when the group in question might be a security group
  - '`ExpandGroups`' results in large result files
  - Neither '`ExportDistributionGroupMembers`' nor '`ExpandGroups`' can handle the following case: User A grants group X a certain permission. Group Y is a member of group X, group Z is a member of group Y. Group Z does not have any members, but we need to know that future members of group Z will have access to the permission granted by user A.

The most economic solution to all these problems is the following:
- Export all the permissions you are interested in
- Do not use '`ExportDistributionGroupMembers`' or '`ExpandGroups`'
- Use '`ExportFileFilter`' to filter for '`Trustee AD ObjectGUID`', looking for the following AD ObjectGUIDs:
  - The AD ObjectGUID of the object you are looking for
  - All AD ObjectGUIDs of groups the initial object is a direct or indirect member of.
  - In the example above, the following GUIDs are needed:
    - AD ObjectGUID of group Z (because we are looking for permissions granted to group Z directly or indirectly)
    - AD ObjectGUID of group Y (because group Z is a member of group Y)
    - AD ObjectGUID of group X (because group Z is a member of group Y, and group Y is a member of group X)

Getting all these GUIDs can be a lot of work. Use the sample code '`MemberOfRecurse.ps1`', which is described later in this document, to make this task as simple as possible.
## 2.11. How to find distribution lists without members?
When a distribution group has no members, E-Mails sent to the group's address are lost because there is no member Exchange could distribute the e-mail to. The sender is not informed about this, as an empty distribution group is a valid recipient, so no Non-Delivery Report is generated.

When looking for distribution groups, counting the direct members is not enough. A group can have another group as only member, and this other group can be empty.

Use the following configuration to reliably identify empty distribution groups:
```
$params = @{
    ExportMailboxAccessRights                   = $false
    ExportMailboxFolderPermissions              = $false
    ExportSendAs                                = $false
    ExportSendOnBehalf                          = $false
    ExportManagedBy                             = $false
    ExportLinkedMasterAccount                   = $false
    ExportPublicFolderPermissions               = $false
    ExportForwarders                            = $false
    ExportManagementRoleGroupMembers            = $false
    ExportDistributionGroupMembers              = 'All'
    ExportGroupMembersRecurse                   = $true
    ExpandGroups                                = $false
    ExportGuids                                 = $true
    ExportGrantorsWithNoPermissions             = $true
    ExportTrustees                              = 'All'

    GrantorFilter                               = "
        if (
            `$Grantor.RecipientTypeDetails -ilike ""*Group*""
        ) {
            `$true
        } else {
            `$false 
        }
    "
    TrusteeFilter                               = $null
    ExportFileFilter                            = "
        if ([string]::IsNullOrEmpty(`$ExportFileLine.Permission) -eq `$true) {
            `$true
        } else {
            `$false 
        }
    "

    ExportFile                                  = '..\export\Export-RecipientPermissions_DVSV-Verteiler_Result.csv'
    ErrorFile                                   = '..\export\Export-RecipientPermissions_DVSV-Verteiler_Error.csv'
    DebugFile                                   = $null

    verbose                                     = $false
}


& .\Export-RecipientPermissions\Export-RecipientPermissions.ps1 @params
```
### 2.11.1. How to export permissions for specific public folders?
You need three things for this:
- GrantorFilter should only include Public Folder Mailboxes
- ExportFileFilter needs to remove everything not of interest

The following example shows how to export permissions granted on the public folder '/X', '/Y' and their subfolders, plus all members of groups granted permissions:
```
$params = @{
    ExportFromOnPrem                            = $true
    UseDefaultCredential                        = $true

    ExportMailboxAccessRights                   = $false
    ExportMailboxAccessRightsSelf               = $false
    ExportMailboxAccessRightsInherited          = $false
    ExportMailboxFolderPermissions              = $false
    ExportMailboxFolderPermissionsAnonymous     = $true
    ExportMailboxFolderPermissionsDefault       = $true
    ExportMailboxFolderPermissionsOwnerAtLocal  = $true
    ExportMailboxFolderPermissionsMemberAtLocal = $true
    ExportSendAs                                = $false
    ExportSendAsSelf                            = $false
    ExportSendOnBehalf                          = $true
    ExportManagedBy                             = $false
    ExportLinkedMasterAccount                   = $false
    ExportPublicFolderPermissions               = $true
    ExportPublicFolderPermissionsAnonymous      = $true
    ExportPublicFolderPermissionsDefault        = $true
    ExportForwarders                            = $false
    ExportManagementRoleGroupMembers            = $false
    ExportDistributionGroupMembers              = 'OnlyTrustees'
    ExportGroupMembersRecurse                   = $true
    ExpandGroups                                = $false
    ExportGuids                                 = $true
    ExportGrantorsWithNoPermissions             = $true
    ExportTrustees                              = 'All'

    RecipientProperties                         = @()
    GrantorFilter                               = "
        if (
            (`$Grantor.RecipientTypeDetails -ieq 'PublicFolderMailbox')
        ) {
            `$true
        } else {
            `$false
        }
    "
    TrusteeFilter                               = $null
    ExportFileFilter                            = "
        if (
            (
                (`$ExportFileLine.'Grantor Recipient Type' -ieq 'UserMailbox/PublicFolderMailbox') -and
                (
                    (`$ExportFileLine.'Folder' -ieq '/X') -or
                    (`$ExportFileLine.'Folder' -ilike '/X/*') -or
                    (`$ExportFileLine.'Folder' -ieq '/Y') -or
                    (`$ExportFileLine.'Folder' -ilike '/Y/*')
                )
            ) -or
            (
                `$ExportFileLine.'Grantor Recipient Type' -ine 'UserMailbox/PublicFolderMailbox'
            )
        ) {
            `$true
        } else {
            `$false 
        }
    "

    ExportFile                                  = '..\export\Export-RecipientPermissions_Result.csv'
    ErrorFile                                   = '..\export\Export-RecipientPermissions_Error.csv'
    DebugFile                                   = ''

    verbose                                     = $true
}


& .\Export-RecipientPermissions\Export-RecipientPermissions.ps1 @params
```
### 2.11.2. I receive an error message when connecting to Exchange on premises
You receive the following error message when connecting to Exchange on-prem:
- English: `'Import-PSSession : Index was out of range. Must be non-negative and less than the size of the collection.'`
- German: `'Import-PSSession : Der Index lag außerhalb des Bereichs. Er darf nicht negativ und kleiner als die Sammlung sein.'`

The error does not come up when creating a Remote PowerShell session, only when creating a local PowerShell connection to Exchange.

The root cause seems to be an AppLocker configuration. I do not yet know which exact setting causes the problem, but the solution consists of the following steps:
- Configure an AppLocker exclusion that allows you to run programs from a defined local file folder, for example `'C:\AppLockerExclude'`
- In the PowerShell session you start Export-RecipientPermissions from, set the temp to this folder or a subfolder:  
  ```
  $env:tmp = 'c:\AppLockerExclude\PowerShell.temp'
  ```
# 3. Sample code
## 3.1. Get-DependentRecipients.ps1
The script can be found in '`.\sample code\Get-DependentRecipients`'.

Currently, only some recipient permissions work cross-premises according to Microsoft (see https://learn.microsoft.com/en-us/exchange/permissions for details).

All other permissions, including the one to manage the members of distribution lists, only work when both the grantor and the trustee are hosted on the same environment.
There are environments where additional permissions work cross-premises, but there is no offical support from Microsoft.

This sample script takes a list of recipients and the output of Export-RecipientPermissions.ps1 to create a list of all recipients and groups that have a grantor-trustee dependency beyond "full access" to each other.

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
## 3.4. MemberOfRecurse.ps1
The script can be found in '`.\sample code\other samples`'.

This sample code shows how to list the GUIDs of all groups a certain AD object is a member of. The script considers
- nested groups
- security groups, no matter of which group scope they are
- distribution groups (static only), no matter of which group scope they are
- group membership in trusted domains/forests (incl. nested groups)

These GUIDs can then be used to answer the question 'Which resources does a particular user or group have access to?', which is described in detail in the FAQ section of this document. 
# 4. Recommendations
Make sure you have the latest updates installed to avoid memory leaks and CPU spikes (PowerShell, .Net).

If possible, allow Export-RecipientPermissions.ps1 to use your on premises infrastructure for best performance.

Start the script from PowerShell, not from within the PowerShell ISE. This makes especially Get-DependentMailboxes.ps1 run faster due to a different default thread apartment mode.

When running the scripts as scheduled job, make sure to include the "-ExecutionPolicy Bypass" parameter.
Example: `powershell.exe -ExecutionPolicy Bypass -file "c:\path\Export-RecipientPermissions.ps1"`
