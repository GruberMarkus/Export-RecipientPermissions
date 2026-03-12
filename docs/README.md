<!-- omit in toc -->
## **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="../src/logo/Export-RecipientPermissions%20Logo.png" width="400" title="Export-RecipientPermissions" alt="Export-RecipientPermissions"></a>**<br>Document, filter and compare Exchange permissions<br><br><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions?labelColor=black&color=informational" alt=""></a> <!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational?labelColor=black&color=informational" alt=""></a> XXXRemoveWhenBuildingXXX--> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational&labelColor=black" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions?labelColor=black" alt="" data-external="1"></a> <a href="https://explicitconsulting.at/open-source/export-recipientpermissions/" target="_blank"><img src="https://img.shields.io/badge/professional%20support-ExplicIT%20Consulting-lawngreen?labelColor=black" alt="professional support from ExplicIT Consulting"></a>

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

Easens the move to the cloud, as permission dependencies beyond the [supported cross-premises permissions](https://docs.microsoft.com/en-us/Exchange/permissions) can easily be identified and even be represented graphically (sample code included).

Compare exports from different times to detect permission changes (sample code included). 

# Table of Contents <!-- omit in toc -->
- [Export-RecipientPermissions.ps1](#export-recipientpermissionsps1)
  - [Output](#output)
  - [Parameters](#parameters)
    - [ConnectionParametersCloud](#connectionparameterscloud)
    - [ConnectionParametersOnprem](#connectionparametersonprem)
    - [PreferredEnvrionment](#preferredenvrionment)
    - [ParallelJobsAD](#paralleljobsad)
    - [ParallelJobsLocal](#paralleljobslocal)
    - [RecipientProperties](#recipientproperties)
    - [GrantorFilter](#grantorfilter)
    - [TrusteeFilter](#trusteefilter)
    - [ExportFileFilter](#exportfilefilter)
    - [ExportMailboxAccessRights](#exportmailboxaccessrights)
    - [ExportMailboxAccessRightsSelf](#exportmailboxaccessrightsself)
    - [ExportMailboxAccessRightsInherited](#exportmailboxaccessrightsinherited)
    - [ExportMailboxFolderPermissions](#exportmailboxfolderpermissions)
    - [ExportMailboxFolderPermissionsAnonymous](#exportmailboxfolderpermissionsanonymous)
    - [ExportMailboxFolderPermissionsDefault](#exportmailboxfolderpermissionsdefault)
    - [ExportMailboxFolderPermissionsOwnerAtLocal](#exportmailboxfolderpermissionsowneratlocal)
    - [ExportMailboxFolderPermissionsMemberAtLocal](#exportmailboxfolderpermissionsmemberatlocal)
    - [ExportMailboxFolderPermissionsExcludeFoldertype](#exportmailboxfolderpermissionsexcludefoldertype)
    - [ExportSendAs](#exportsendas)
    - [ExportSendAsSelf](#exportsendasself)
    - [ExportSendOnBehalf](#exportsendonbehalf)
    - [ExportManagedBy](#exportmanagedby)
    - [ExportLinkedMasterAccount](#exportlinkedmasteraccount)
    - [ExportPublicFolderPermissions](#exportpublicfolderpermissions)
    - [ExportPublicFolderPermissionsAnonymous](#exportpublicfolderpermissionsanonymous)
    - [ExportPublicFolderPermissionsDefault](#exportpublicfolderpermissionsdefault)
    - [ExportPublicFolderPermissionsExcludeFoldertype](#exportpublicfolderpermissionsexcludefoldertype)
    - [ExportForwarders](#exportforwarders)
    - [ExportModerators](#exportmoderators)
    - [ExportRequireAllSendersAreAuthenticated](#exportrequireallsendersareauthenticated)
    - [ExportAcceptMessagesOnlyFrom](#exportacceptmessagesonlyfrom)
    - [ExportResourceDelegates](#exportresourcedelegates)
    - [ExportManagementRoleGroupMembers](#exportmanagementrolegroupmembers)
    - [ExportDistributionGroupMembers](#exportdistributiongroupmembers)
    - [ExportGroupMembersRecurse](#exportgroupmembersrecurse)
    - [ExpandGroups](#expandgroups)
    - [ExportGuids](#exportguids)
    - [ExportSids](#exportsids)
    - [ExportGrantorsWithNoPermissions](#exportgrantorswithnopermissions)
    - [ExportTrustees](#exporttrustees)
    - [ExportFile](#exportfile)
    - [ErrorFile](#errorfile)
    - [DebugFile](#debugfile)
    - [UpdateInterval](#updateinterval)
  - [Requirements](#requirements)
    - [Common](#common)
    - [Cloud](#cloud)
    - [On-prem](#on-prem)
    - [Hybrid environment considerations](#hybrid-environment-considerations)
- [FAQ](#faq)
  - [Which permissions are required?](#which-permissions-are-required)
  - [Can the script resolve permissions granted to a group to it's individual members?](#can-the-script-resolve-permissions-granted-to-a-group-to-its-individual-members)
  - [Where can I find the changelog?](#where-can-i-find-the-changelog)
  - [How can I contribute, propose a new feature or file a bug?](#how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [How can I get more script output for troubleshooting?](#how-can-i-get-more-script-output-for-troubleshooting)
  - [A permission is reported, but the trustee details (primary SMTP address etc.) are empty](#a-permission-is-reported-but-the-trustee-details-primary-smtp-address-etc-are-empty)
  - [Isn't a plural noun in the script name against PowerShell best practices?](#isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
  - [Is there a GUI available?](#is-there-a-gui-available)
  - [Which resources does a particular user or group have access to?](#which-resources-does-a-particular-user-or-group-have-access-to)
  - [How to find distribution lists without members?](#how-to-find-distribution-lists-without-members)
    - [How to export permissions for specific public folders only?](#how-to-export-permissions-for-specific-public-folders-only)
    - [I receive an error message when connecting to Exchange on premises](#i-receive-an-error-message-when-connecting-to-exchange-on-premises)
- [Sample code](#sample-code)
  - [Get-DependentRecipients.ps1](#get-dependentrecipientsps1)
  - [Compare-RecipientPermissions.ps1](#compare-recipientpermissionsps1)
  - [FiltersAndSidhistory.ps1](#filtersandsidhistoryps1)
  - [MemberOfRecurse.ps1](#memberofrecurseps1)


# Export-RecipientPermissions.ps1
Finds all recipients with a primary SMTP address in an on on-prem or online Exchange environment and documents their
- mailbox access rights,
- mailbox folder permissions,
- "send as" permissions,
- "send on behalf" permissions, and
- "managed by" permissions,
- and many more.

The idea is to export and document these different permissions in a common format, something that is not possible with the standard Exchange cmdlets.

This common format is not only useful for documentation, but also to compare permissions over time, between different recipients, or event between different tenants - all of this can be automated.

The ability to resolve groups to its transitive members not only allows you to detect that a permission has changed, but also to detect that a permission granted to a group has not changed, but the members of the group have.

You can use this information for interesting use cases:
- Regular auditing of sensitive mailboxes or sensitive permissions.
- Regularly inform your VIPs about the permissions they have set in their mailboxes and what has changed since the last report.
- Find out which recipient has access to which ressources.
- Use the output of Export-RecipientPermissions in [Set-OutlookSignatures](https://github.com/Set-OutlookSignatures), which allows for dynamic assignment of email signatures based on up-to-date permissions set in your Exchange environment.

## Output
The report is saved to the file 'Export-RecipientPermissions_Result.csv', which consists of the following columns:
- Grantor Primary SMTP: The primary SMTP address of the object granting a permission
  - When management role group members are exported, this column contains 'Management Role Group'
  - When public folder permissions are exported, this column represents the folder's content mailbox
- Grantor Display Name: The display name of the grantor.
  - When management role group members are exported, this column contains the name of the Management Role Group
  - When public folder permissions are exported, this column represents the folder's content mailbox
- Grantor Exchange GUID: The Exchange GUID of the grantor.
  - Only available when ExportGuids is enabled.
- Grantor AD ObjectGUID: The Active Directory ObjectGUID of the grantor.
  - Only available when ExportGuids is enabled.
- Grantor SID: The SID (wecurity identifier) of the grantor.
  - Only available when ExportSids is enabled.
- Grantor Recipient Type: The recipient type and recipient type detail of the grantor.
  - When management role group members are exported, this column contains 'ManagementRoleGoup'
  - When public folder permissions are exported, this column represents the folder's content mailbox ('UserMailbox/PublicFolderMailbox')
- Grantor Environment: Shows if the grantor is held on-prem or in the cloud.
  - When public folder permissions are exported, this column represents the folder's content mailbox
- Folder: Folder the permission is granted on
  - Empty for non-folder permissions
  - All folder names start with '/', '/' representing the root folder
- Permission: The permission granted/received (e.g., FullAccess, SendAs, SendOnBehalf etc.)
  - When public folder permissions are exported and a folder is mail-enabled, the virtual permission 'MailEnabled' is exported
  - When management role group members are exported, a "virtual" right 'MemberRecurse' or 'MemberDirect' is exported
  - When forwarders are exported, one or more of the following virtual permissions are exported:
    - Forward_ExternalEmailAddress_ForwardOnly
    - Forward_ForwardingAddress_DeliverAndForward
    - Forward_ForwardingAddress_ForwardOnly
    - Forward_ForwardingSmtpAddress_DeliverAndForward
    - Forward_ForwardingSmtpAddress_ForwardOnly
- Allow/Deny: Shows if the permission is an allow or a deny permission.
- Inherited: Shows if the permission is inherited or not.
- InheritanceType: Shows if the permission is also valid for child objects, and if yes, which child objects.
- Trustee Original Identity: The original identity string of the trustee.
  - When [ExpandGroups](#expandgroups) is enabled, this column contains the original identity string of the original trustee groups, appended with the string '`     [MemberRecurse] `' or '`     [MemberDirect] `' and the original identity of the resolved group member
- Trustee Primary SMTP: The primary SMTP address of the object receiving a permission.
  - When [ExpandGroups](#expandgroups) is enabled, the primary SMTP address comes from the resolved group member
- Trustee Display Name: The display name of the trustee.
  - When [ExpandGroups](#expandgroups) is enabled, the display name comes from the resolved group member
- Trustee Exchange GUID: The Exchange GUID of the trustee.
  - Only available when [ExportGuids](#exportguids) is enabled.
- Trustee AD ObjectGUID: The Active Directory ObjectGUID of the trustee.
  - Only available when [ExportGuids](#exportguids) is enabled.
- Trustee SID: The SID (wecurity identifier) of the trustee.
  - Only available when [ExportSids](#exportsids) is enabled.
- Trustee Recipient Type: The recipient type of the trustee.
-   - When [ExpandGroups](#expandgroups) is enabled, the recipient type comes from the resolved group member
- Trustee Environment: Shows if the trustee is held on-prem or in the cloud.
  - When [ExpandGroups](#expandgroups) is enabled, the trustee environment comes from the resolved group member


## Parameters
### ConnectionParametersCloud
A hashtable containing all connection parameters and options for the connection to Exchange Online.

The parameter definition is self-describing:

```
[hashtable]$ConnectionParametersCloud = @{
        UriList              = @('https://outlook.office365.com/') # Only change when not using the public M365 cloud
        ConnectionParameters = @{
            # Define parameters for Connect-ExchangeOnline here
        }
        ParallelJobs         = 3 # Number of concurrent connections
    }
```

Here is a real-world example: 
```
$params = @{
    # Exchange connection parameters
    ## Cloud
    ConnectionParametersCloud                       = @{
        UriList              = @('https://outlook.office365.com/')
        ConnectionParameters = @{
            AppId                 = 'The app ID of the Entra ID app to use'
            CertificateThumbprint = 'Thumbprint string of certificate'
            organization          = 'Your onmicrosoft.com domain'
        }
        ParallelJobs         = 3
    }

    # Define more script parameters here
}

.\Export-RecipientPermissions.ps1 @params
```

It is strongly recommended to use app-only authentication for attended as well as unattended scripts, see the '[Requirements](#requirements)' chapter for details.

### ConnectionParametersOnprem
A hashtable containing all connection parameters and options for the connection to Exchange Online.

The parameter definition is self-describing:

```
[hashtable]$ConnectionParametersOnprem = @{
    UriList              = @() # Empty array will later be replaced with all Exchange servers with the mailbox role (from an AD query)
    UseDefaultCredential = $true # Use default credential of logged-on user?
    UsernameFile         = '.\Export-RecipientPermissions_CredentialUsername.txt' # Ignored when UseDefaultCredention is not $true
    PasswordFile         = '.\Export-RecipientPermissions_CredentialPassword.txt' # Ignored when UseDefaultCredention is not $true
    ParallelJobs         = $null # Number of concurrent connections. $null will be replaced with the count of UriList entries (1 connection per server).
}
```

Here is a real-world example using the credential of the logged-on user: 
```
$params = @{
    # Exchange connection parameters
    ## On-prem
    ConnectionParametersOnprem                      = @{
        UriList              = @() # Empty array will later be replaced with all Exchange servers with the mailbox role (from an AD query)
        UseDefaultCredential = $true
        ParallelJobs         = $null # $null will later be replaced with the count of UriList entries
    }

    # Define more script parameters here
}

.\Export-RecipientPermissions.ps1 @params
```

Here is a real-world example using a separate credential. UsernameFile and PasswordFile do not contain clear-text information, but their representations as secure strings (decryptable only by the user who encrypted it, and only on the machine used to encrypt it). If the files do not yet exist, you are interactively asked to provide the credential.
```
$params = @{
    # Exchange connection parameters
    ## On-prem
    ConnectionParametersOnprem                      = @{
        UriList              = @() # Empty array will later be replaced with all Exchange servers with the mailbox role (from an AD query)
        UseDefaultCredential = $false
        UsernameFile         = '.\Export-RecipientPermissions_CredentialUsername.txt'
        PasswordFile         = '.\Export-RecipientPermissions_CredentialPassword.txt'
        ParallelJobs         = $null # $null will later be replaced with the count of UriList entries
    }

    # Define more script parameters here
}

.\Export-RecipientPermissions.ps1 @params
```

### PreferredEnvrionment
When connected to both, cloud and on-prem environments, this parameter defines which environment should be preferred to gather data.

The other environment is only queried for information that is available exclusively on the other side.

See the '[Requirements](#requirements)' chapter for good practices for hybrid environments.

Default value: 'On-prem'

Allowed values: 'Cloud', 'On-prem'

### ParallelJobsAD
Maximum Active Directory queries running in parallel.

Default value: 50

Allowed values: Integer greater than 0

### ParallelJobsLocal
Maximum local jobs running in parallel.

Default value: 50

Allowed values: Integer greater than 0

### RecipientProperties
Additional recipient properties to import for use with [GrantorFilter](#grantorfilter) and [TrusteeFilter](#trusteefilter).

Properties that are always included: 'Identity', 'DistinguishedName', 'RecipientType', 'RecipientTypeDetails', 'DisplayName', 'Name', 'PrimarySmtpAddress', 'EmailAddresses', 'ManagedBy', 'UserFriendlyName', 'LinkedMasterAccount'

### GrantorFilter
Only check grantors where the filter criteria matches $true.

The variable `$Grantor` has all attributes defined by '`RecipientProperties`'. For example:
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

Pay attention how the reference variable `$Grantor` is prefixed with a backtick ('`` ` ``'), so it is passed to Export-RecipientPermissions.ps1 as literal string and therefore resolved within the export script and not already before (which would lead to an empty string in this sample code).

Default value: $null

### TrusteeFilter
Only report trustees where the filter criteria matches $true.

If the trustee matches a recipient, the available attributes are the same as for [GrantorFilter](#grantorfilter), only the reference variable is $Trustee instead of $Grantor.

If the trustee does not match a recipient (because it no longer exists, for exampe), `$Trustee` is just a string. In this case, the export shows the following:
- Column "Trustee Original Identity" contains the trustee description string as reported by Exchange
- Columns "Trustee Primary SMTP" and "Trustee Display Name" are empty

Example:
```
"`$Trustee.primarysmtpaddress -ilike '*@example.com'"
```

Pay attention how the reference variable `$Trustee` is prefixed with a backtick ('`` ` ``'), so it is passed to Export-RecipientPermissions.ps1 as literal string and therefore resolved within the export script and not already before (which would lead to an empty string in this sample code).

Default value: $null

### ExportFileFilter
Only report results where the filter criteria matches `$true`.

This filter works against every single row (line) of the results found. ExportFile will only contain lines where this filter returns $true.

The `$ExportFileLine` variable is an object with the header names from $ExportFile as string properties
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

Example:
```
"`$ExportFileLine.'Trustee Environment' -ieq 'On-prem'"
```

Pay attention how the reference variable `$ExportFileLine` is prefixed with a backtick ('`` ` ``'), so it is passed to Export-RecipientPermissions.ps1 as literal string and therefore resolved within the export script and not already before (which would lead to an empty string in this sample code).

Default value: $null

### ExportMailboxAccessRights
Report rights set on the mailbox itself, such as "FullAccess" and "ReadAccess".

Default value: $true

Allowed values: $true, $false

### ExportMailboxAccessRightsSelf
Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.).

[ExportMailboxAccessRights](#exportmailboxaccessrights) must be enabled for this feature to work.

Default value: $false

Allowed values: $true, $false

### ExportMailboxAccessRightsInherited
Report inherited mailbox access rights (only works on-prem, as Exchange Online does not expose these rights).

[ExportMailboxAccessRights](#exportmailboxaccessrights) must be enabled for this feature to work.

Default value: $false

Allowed values: $true, $false

### ExportMailboxFolderPermissions
Report permissions for each maibox folder. This can take very long, as Exchange is notoriously slow reporting folder permissions, no matter if on-prem or in the cloud.

Default value: $false

Allowed values: $true, $false

### ExportMailboxFolderPermissionsAnonymous
Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.).

[ExportMailboxFolderPermissions](#exportmailboxfolderpermissions) must be enabled for this feature to work.

Default value: $true

Allowed values: $true, $false

### ExportMailboxFolderPermissionsDefault
Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.).

[ExportMailboxFolderPermissions](#exportmailboxfolderpermissions) must be enabled for this feature to work.

Default value: $true

Allowed values: $true, $false

### ExportMailboxFolderPermissionsOwnerAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.

[ExportMailboxFolderPermissions](#exportmailboxfolderpermissions) must be enabled for this feature to work.

Default value: $false

Allowed values: $true, $false

### ExportMailboxFolderPermissionsMemberAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.

[ExportMailboxFolderPermissions](#exportmailboxfolderpermissions) must be enabled for this feature to work.

Default value: $false

Allowed values: $true, $false

### ExportMailboxFolderPermissionsExcludeFoldertype
List of mailbox folder types to ignore.

[ExportMailboxFolderPermissions](#exportmailboxfolderpermissions) must be enabled for this feature to work.

Some known folder types are: Audits, Calendar, CalendarLogging, CommunicatorHistory, Conflicts, Contacts, ConversationActions, DeletedItems, Drafts, ExternalContacts, Files, GalContacts, ImContactList, Inbox, Journal, JunkEmail, LocalFailures, Notes, Outbox, QuickContacts, RecipientCache, RecoverableItemsDeletions, RecoverableItemsPurges, RecoverableItemsRoot, RecoverableItemsVersions, Root, RssSubscription, SentItems, ServerFailures, SyncIssues, Tasks, WorkingSet, YammerFeeds, YammerInbound, YammerOutbound, YammerRoot

Default value: 'audits'

### ExportSendAs
Export Send As permissions.

Default value: $true

Allowed values: $true, $false

### ExportSendAsSelf
Export Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.).

[ExportSendAs](#exportsendas) must be enabled for this feature to work.

Default value: $false

Allowed values: $true, $false

### ExportSendOnBehalf
Export Send On Behalf permissions.

Default value: $true

Allowed values: $true, $false

### ExportManagedBy
Export Managed By permissions. Only works for distribution groups, and is not to be confused with the "Manager" attribute.

Default value: $true

Allowed values: $true, $false

### ExportLinkedMasterAccount
Export Linked Master Account information. Only works on-prem, as there are no such objects in Exchange Online.

Default value: $true

Allowed values: $true, $false

### ExportPublicFolderPermissions
Export public folder permissions. This part of the report can take very long, as Exchange is notoriously slow reporting folder permissions, no matter if on-prem or in the cloud.

When enabled, the `$GrantorFilter` variable refers to the according public folder content mailbox.

Default value: $true

Allowed values: $true, $false

### ExportPublicFolderPermissionsAnonymous
Report public folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.).

[ExportPublicFolderPermissions](#exportpublicfolderpermissions) must be enabled for this feature to work.

Default value: $true

Allowed values: $true, $false

### ExportPublicFolderPermissionsDefault
Report public folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.).

[ExportPublicFolderPermissions](#exportpublicfolderpermissions) must be enabled for this feature to work.

Default value: $true

Allowed values: $true, $false

### ExportPublicFolderPermissionsExcludeFoldertype
List of public folder types to ignore.

Some known folder types are: IPF.Appointment, IPF.Contact, IPF.Note, IPF.Task

[ExportPublicFolderPermissions](#exportpublicfolderpermissions) must be enabled for this feature to work.

Default value: ''

### ExportForwarders
Export forwarders.

There are three types of forwarders, often leading to confusion:
- '`ExternalEmailAddress`' ('`targetAddress`' in Active Directory)
  - Highest priority.
  - Can be configured on basically every mail-enabled object.
  - Can point to any SMTP address, existing in your directory or somewhere outside.
  - Is typically used for contacts, migration and co-existence scenarios.
  - '`DeliverToMailboxAndForward`' ('`deliverAndRedirect`' in Active Directory) is ignored, all e-mails sent to the recipient will unconditionally be forwarded without storing a copy or sending it to group members.
- '`ForwardingAddress`' ('`altRecipient`' in Active Directory)
  - Medium priority.
  - Can be configured for mailboxes and mail-enabled public folders.
  - Needs a mail-enabled Object existing in your directory as target (a contact is required to forward to external e-mail addresses).
  - This property is used when forwarding is configured in the Exchange Control Panel or the Exchange Admin Center.
  - '`DeliverToMailboxAndForward`' ('`deliverAndRedirect`' in Active Directory) defines if the e-mail is forwarded only, or forwarded and stored.
- '`ForwardingSmtpAddress`' ('`msExchGenericForwardingAddress`' in Active Directory)
  - Lowest priority.
  - Can be configured for mailboxes.
  - Can point to any SMTP address, existing in your directory or somewhere outside.
  - This property is used when a user configures forwarding for his mailbox in Outlook Web.
  - '`DeliverToMailboxAndForward`' ('`deliverAndRedirect`' in Active Directory) defines if the e-mail is forwarded only, or forwarded and stored.

When forwarders are exported, one or more of the following virtual permissions are exported:
- Forward_ExternalEmailAddress_ForwardOnly
- Forward_ForwardingAddress_DeliverAndForward
- Forward_ForwardingAddress_ForwardOnly
- Forward_ForwardingSmtpAddress_DeliverAndForward
- Forward_ForwardingSmtpAddress_ForwardOnly

Default value: $true

Allowed values: $true, $false

### ExportModerators
Exports the virtual permissions 'ModeratedBy' and 'ModeratedByBypass', listing all users and groups which are configured as moderators for a recipient or can bypass moderation.

Only works for recipients with moderation enabled.

Default value: $true

Allowed values: $true, $false

### ExportRequireAllSendersAreAuthenticated
Exports the virtual permission 'RequireAllSendersAreAuthenticated' with the trustee 'NT AUTHORITY\Authenticated Users' for each recipient which is configured to only receive messages from authenticated (internal) senders.

Default value: $true

Allowed values: $true, $false

### ExportAcceptMessagesOnlyFrom
Exports the virtual permission 'AcceptMessagesOnlyFrom' for each recipient configured to only receive messages from selected (internal) senders.

The attributes 'AcceptMessagesOnlyFrom' and 'AcceptMessagesOnlyFromDLMembers' are exported as the same virtual permission 'AcceptMessagesOnlyFrom'.

Default value: $true

Allowed values: $true, $false

### ExportResourceDelegates
Exports information about who is allowed or denied to book resources (rooms or equipment) and to accept or reject booking requests.

The following virtual permissions are exported:
- ResourceDelegate
- ResourcePolicyDelegate_AllBookInPolicy
- ResourcePolicyDelegate_AllRequestInPolicy
- ResourcePolicyDelegate_AllRequestOutOfPolicy
- ResourcePolicyDelegate_BookInPolicy
- ResourcePolicyDelegate_RequestInPolicy
- ResourcePolicyDelegate_RequestOutOfPolicy

ResourcePolicyDelegate_AllBookInPolicy, ResourcePolicyDelegate_AllRequestinPolicy, ResourcePolicyDelegate_AllRequestOutOfPolicy: 'Everyone' is used as trustee.

ResourcePolicyDelegate_BookInPolicy, ResourcePolicyDelegate_RequestInPolicy, ResourcePolicyDelegate_RequestOutOfPolicy: Each of these virtual permissions is reported even when the corresponding 'All'-right is enabled.

Default value: $true

Allowed values: $true, $false

### ExportManagementRoleGroupMembers
Export members of management role groups.

The virtual permission 'MemberRecurse' or 'MemberDirect' is used in the export file.

[GrantorFilter](#grantorfilter) does not apply to the export of management role groups, but [TrusteeFilter](#trusteefilter) and [ExportFileFilter](#exportfilefilter) do.

Default value: $true

Allowed values: $true, $false

### ExportDistributionGroupMembers
Export distribution group members, including nested groups and dynamic groups.

The parameter [ExpandGroups](#expandgroups) can be used independently from ExportDistributionGroupMembers:
- ExpandGroups acts when a group is used as trustee: It adds every recurse member of the group as a separate trustee entry.
- ExportDistributionGroupMembers exports the distribution group as grantor, which the recurse members as trustees.

Default value: 'None'

Allowed values: 'None', 'All', 'OnlyTrustees'
- 'None': Distribution group members are not exported. [ExpandGroups](#expandgroups) can still be used.
- 'All': Members of all distribution groups are exported, parameter [GrantorFilter](#grantorfilter) is considerd.
- 'OnlyTrustees': Only export members of those distribution groups that are used as trustees, even when they are excluded via [GrantorFilter](#grantorfilter).

### ExportGroupMembersRecurse
When disabled, only direct members of groups are exported, and the virtual permission 'MemberDirect' is used in the export file.

When enabled, recursive members of groups are exported, and the virtual permission 'MemberRecurse' is used in the export file.

Default value: $false

Allowed values: $true, $false

### ExpandGroups
Expand groups to their recursive members, including nested groups and dynamic groups.

This is useful in cases where users are sent permission reports, as not only permission changes but also changes in the underlying trustee groups are documented and directly associated with a grantor-permission-trustee triplet.  
For example: User A has granted Group B permission X a long time ago. The permission itself did not change, but the recursive members of Group B changed over time. With ExpandGroups enabled, the members and therefore the changes of Group B are documented with every run of Export-RecipientPermissions.

This may drastically increase script run time and file size.

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

[TrusteeFilter](#trusteefilter) is applied to trustee groups as well as to their finally expanded individual members
- Nested groups are expanded to individual members, but [TrusteeFilter](#trusteefilter) is not applied to the nested group

Default value: $false

Allowed values: $true, $false

### ExportGuids
When enabled, the export contains the Exchange GUID and the AD/Entra ObjectGUID for each grantor and trustee.

Default value: $false

Allowed values: $true, $false

### ExportSids
When enabled, the export contains the SID (security identifier) for each grantor and trustee.

Default value: $false

Allowed values: $true, $false

### ExportGrantorsWithNoPermissions
Per default, Export-RecipientPermissions only exports grantors which have set at least one permission for at least one trustee.
If all grantors should be exported, set this parameter to $true.

If enabled, a grantor that that not grant any permission is included in the list with the following columns: "Grantor Primary SMTP", "Grantor Display Name", "Grantor Recipient Type", "Grantor Environment". The other columns for this recipient are empty.

Default value: $false

Allowed values: $true, $false

### ExportTrustees
Include either all trustees in the permission report file, only valid ones, or only invalid ones.

Valid trustees are trustees which can be resolved to an Exchange recipient.

Default value: 'All'

Allowed values: 'All', 'OnlyValid', 'OnlyInvalid'

### ExportFile
Name (and path) of the permission report file.

Default value: '.\export\Export-RecipientPermissions_Result.csv'

### ErrorFile
Name (and path) of the error log file.

Set to $null or '' to disable error logging.

Default value: '.\export\Export-RecipientPermissions_Error.csv',

### DebugFile
Name (and path) of the debug log file.

Set to $null or '' to disable debug logging.

Default value: ''

### UpdateInterval
Interval how often the job progress is updated in the debug file only. The console output is updated every second.

Updates are based von recipients (or other elements) done, not on duration.

Default value: 100

Allowed values: Integer greater 0


## Requirements
### Common
Windows Powershell 5.1 and PowerShell 7+ are supported. Use PowerShell 7+ for best performance, especially when Export-RecipientPermissions is used with Exchange Online.

### Cloud
For connecting to the cloud:
- Make sure the ExchangeOnlineManagement PowerShell module is installed.
- Do not use account credentials to connect to Exchange Online, as ExchangeOnlineManagement is retiring this feature. Use [certificate based authenticatino or app-only authentication](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps).

The script needs to be run with an account that has at least read permissions to all recipients. See '[Which permissions are required?](#which-permissions-are-required)' for a complete overview.

Connection parameters are defined via the '[ConnectionParametersCloud](#connectionparameterscloud)' parameter.

### On-prem
The script needs to be run with an account that has at least read permissions to all recipients. See '[Which permissions are required?](#which-permissions-are-required)' for a complete overview.

In automation scenarios, the credential of the account running the script can be used, or a separeate credential can be defined.

Per default, the script uses multiple parallel threads, each one consuming one Exchange PowerShell session. Please watch CPU and RAM usage, as well as your Exchange throttling policy:
```
$IdentityToCheck = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).Name

(Get-ThrottlingPolicyAssociation -Identity $IdentityToCheck) | foreach {
	"THROTTLING POLICY ASSOCIATION"
	$_
	"THROTTLING POLICY DETAILS"
	$_.throttlingpolicyid | Get-ThrottlingPolicy
}
```

### Hybrid environment considerations
Starting with v4.0.0, Export-RecipientPermissions allows to connect to the cloud and the on-prem end of your hybrid configuration simultaneously. Just one script run is now required to get the following permissions which are only available on one side or differ between on-prem and the cloud:
- Mailbox access rights
- Mailbox folder permissions
- Public folder permissions

The following must be considered for this configruation:
- The '`PreferredEnvironment`' parameter defines the preferred or default environment for all queries. All information is gathered from this environment, unless asking the other end of your hybrid connection is absolutely neccessary. Per default, the preferred environment is on-prem due to performance reasons.
- Follow Microsofts good practices for the best experience for users, admins and Export-RecipientPermissions:
  - Make sure that all objects are synced between your on-prem environment and your cloud environment. This is not only about Exchange recipients, but also about security groups.
  - [Enable ACLable object synchronization at the organization level](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/set-up-delegated-mailbox-permissions#enable-aclable-object-synchronization), and [manually enable ACLs on each mailbox moved to the cloud before ACLable object synchronization was enabled at the organization level](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/set-up-delegated-mailbox-permissions#enable-acls-on-remote-mailboxes).
  - [Send As permissions need to be set in both environments](https://learn.microsoft.com/en-us/exchange/permissions?source=recommendations#mailbox-permissions-and-capabilities-not-supported-in-hybrid-environments). Ideally, this is automated with two scripts:
    - Use one script to grant and revoke Send As permissions, making sure the change is performed in both environments.
    - Another script should run regularly, getting the Send As permissions for each maibox from the environment they are hosted in and mirroring these permissions to the other end of the hybrid connection.
  - For Send On Behalf, [Exchange hybrid writeback](https://learn.microsoft.com/en-us/entra/identity/hybrid/connect/reference-connect-sync-attributes-synchronized#exchange-hybrid-writeback) must be enabled in Entra ID Connect, so that the PublicDelegates attribute (GrantSendOnBehalfTo in Exchange on-prem) is synchronized.
  - Regularly run a script that [writes back the ExchangeGUID from mailboxes initially created in Exchange Online to your on-prem environment](https://learn.microsoft.com/en-us/troubleshoot/exchange/move-mailboxes/migrationpermanentexception-when-moving-mailboxes).


# FAQ
## Which permissions are required?
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

You can use the following script to find out which cmdlet is assigned to which management role:
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

## Can the script resolve permissions granted to a group to it's individual members?
Yes, Export-RecipientPermissions can resolve trustee groups to their individual members. Use the parameter [ExpandGroups](#expandgroups) to enable this feature.

## Where can I find the changelog?
The changelog is located in the `'.\docs'` folder, along with other documents related to Export-RecipientPermissions.

## How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, please have a look at `'.\docs\CONTRIBUTING'`.

## How can I get more script output for troubleshooting?
Start the script with the '-verbose' parameter to get the maximum output for troubleshooting.

## A permission is reported, but the trustee details (primary SMTP address etc.) are empty
When excluding a bug in the script, there are three possible reasons why a trustee does not have details like a primary SMTP address in the result file:
- The trustee is a valid directory object, but not a valid Exchange recipient.  
  Examples: 'NT AUTHORITY\SELF', 'Domain Admins', 'Enterprise Admins'
- The initial trustee no longer exists.  
  Exchange does not check if trustees still exist and remove the according permissions in case of deletion - this would be a problem when restoring deleted Active Directory objects.  
  As Exchange stores trustees in different formats, the trustee original identity can be a SID, an NT4-style logon name or just about any string.
- Multiple recipients share the same linked master account, user friendly name, distinguished name, GUID or primary SMTP address. As the search for this value returns multiple recipients, no details are shown.  
  This should not happen when using the built-in Exchange tools, due to their built-in quality checks. It usually happens when Exchange attributes are modified directly in Active Directory.  
  Whe passing the '-verbose' PowerShell parameter, the script outputs recipients with non-unique attributes in the verbose stream.  

## Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Export-RecipientPermissions contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.

## Is there a GUI available?
There is no dedicated graphical user interface.

A basic GUI for configuring the script is accessible via the following built-in PowerShell command:
```
Show-Command .\Export-RecipientPermissions.ps1
```

## Which resources does a particular user or group have access to?
As many other IT systems, Exchange stores permissions as forward links and not as backward links. This means that the information about granted permissions is stored on the granting object (forwarding to the trustee), without storing any information about the granted permission at the trustee object (pointing backwards to the grantor).

This means that we have to query all granted permissions first and then filter those for trustees that involve the user or group we are looking for.

There are some exceptions where Exchange stores permissions not only as forward links but also as backward links, to allow getting a list of certain permissions with just one fast query from the trustee perspective. All these cases rely on automatic calculation of the backlink attribute in Active Directory and Entra ID. Well-known examples are publicDelegates/publicDelegatesBL, member/memberOf, manager/directReports and owner/ownerBL.

But back to the initial question which resources a particular user or group has access to.  
We already know that we need to query all permissions of interest first, and then filter the results. But what should we filter for?

If you are only interested in permissions granted directly to a certain user or group, the search is straight forward:
- If the user or group is an Exchange recipient, use [TrusteeFilter](#trusteefilter) to filter for 'Trustee Primary SMTP'
- If the user or group is not an Exchange recipient, enable '`ExportGUIDs`' and use [ExportFileFilter](#exportfilefilter) to filter for '`Trustee AD ObjectGuid`'

If you are interested in permissions granted directly and indirectly to a certain user or group, it gets more complicated.

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
- Use [ExportFileFilter](#exportfilefilter) to filter for '`Trustee AD ObjectGUID`', looking for the following AD ObjectGUIDs:
  - The AD ObjectGUID of the object you are looking for
  - All AD ObjectGUIDs of groups the initial object is a direct or indirect member of.
  - In the example above, the following GUIDs are needed:
    - AD ObjectGUID of group Z (because we are looking for permissions granted to group Z directly or indirectly)
    - AD ObjectGUID of group Y (because group Z is a member of group Y)
    - AD ObjectGUID of group X (because group Z is a member of group Y, and group Y is a member of group X)

Getting all these GUIDs can be a lot of work. Use the sample code '`MemberOfRecurse.ps1`', which is described later in this document, to make this task as simple as possible.

## How to find distribution lists without members?
When a distribution group has no members, emails sent to the group are lost because there is no member Exchange could distribute the emails to. The sender is not informed about this because an empty distribution group itself is still a valid recipient, so no Non-Delivery Report is generated.

When looking for distribution groups, counting the direct members is not enough. A group can have another group as only member, and this other group can be empty.

Use the following configuration to reliably identify empty distribution groups:
```
$params = @{
    # Do not forget to define ConnectionParametersCloud and/or ConnectionParametersOnprem

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

    ExportFileFilter                            = "
        if ([string]::IsNullOrEmpty(`$ExportFileLine.Permission) -eq `$true) {
            `$true
        } else {
            `$false 
        }
    "
}


& .\Export-RecipientPermissions\Export-RecipientPermissions.ps1 @params
```

### How to export permissions for specific public folders only?
You need two things for this:
- [GrantorFilter](#grantorfilter) should only include Public Folder Mailboxes
- [ExportFileFilter](#exportfilefilter) needs to remove everything not of interest

The following example shows how to export permissions granted on the public folder '/X', '/Y' and their subfolders, plus all members of groups granted permissions:
```
$params = @{
    # Do not forget to define ConnectionParametersCloud and/or ConnectionParametersOnprem

    ExportMailboxAccessRights                   = $false
    ExportMailboxFolderPermissions              = $false
    ExportSendAs                                = $false
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

    GrantorFilter                               = "
        if (
            (`$Grantor.RecipientTypeDetails -ieq 'PublicFolderMailbox')
        ) {
            `$true
        } else {
            `$false
        }
    "

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
}


& .\Export-RecipientPermissions\Export-RecipientPermissions.ps1 @params
```

### I receive an error message when connecting to Exchange on premises
You receive the following error message when connecting to Exchange on-prem:
- English: `'Import-PSSession : Index was out of range. Must be non-negative and less than the size of the collection.'`
- German: `'Import-PSSession : Der Index lag außerhalb des Bereichs. Er darf nicht negativ und kleiner als die Sammlung sein.'`

The error does not come up when creating a Remote PowerShell session, only when creating a local PowerShell connection to Exchange.

The root cause seems to be an AppLocker configuration. I do not yet know which exact setting causes the problem, but the solution consists of the following steps:
- Configure an AppLocker exclusion that allows you to run programs from a defined local file folder, for example 'C:\AppLockerExclude'.
- In the PowerShell session you start Export-RecipientPermissions from, set the temp to this folder or a subfolder:  
  ```
  $env:tmp = 'c:\AppLockerExclude\PowerShell.temp'
  ```

# Sample code
## Get-DependentRecipients.ps1
The script can be found in '`.\sample code\Get-DependentRecipients`'.

Not all recipient permissions work cross-premises, as described in the Microsoft article '[Permissions in Exchange hybrid deployments](see https://learn.microsoft.com/en-us/exchange/permissions)'.  Some permissions only work when both grantor and trustee are in the same environment or when additional manual steps are performed. See '[Hybrid environment considerations](#hybrid-environment-considerations)' for more details on this.

This makes planning a staged migration of mailboxes from on-prem to the cloud a hard task.

Get-DependentRecipients.p1 makes this much easier: It takes a list of recipients and the output of Export-RecipientPermissions to create a list of all recipients and groups that have a grantor-trustee dependency beyond "full access" to each other.

The script not only considers situations where recipient A grants rights to recipient B, but the whole permission chain ("X-Z-A-B-C-D" etc.).

The script does not consider group membership.

The following outputs are created:
- Get-DependentRecipients_Output_InitialRecipients.csv  
  List of initial recipients you want to move to the cloud.
- Export-RecipientPermissions_Output_Permissions.csv  
  The output of Export-RecipientPermissions, reduced to entries that have a connection with the list of recipients you want to move to the cloud.  
  Enhanced with information if a grantor or trustee is part of the initial recipient file or has to be migrated additionally to keep permission chains working.
  Enhanced with information which single permissions start permissions chains outside the initial recipients.
- Get-DependentRecipients_Output_AdditionalRecipients.csv  
  List of additional recipients that need to be migrated to cloud if you do not want to loose permissions.
- Get-DependentRecipients_Output_AllRecipients.csv  
  Combined ist of all initial and additional recipients.
- Get-DependentRecipients_Output_GML.gml  
  All recipients and their permissions in a graphical representation. The GML (Graph Modeling Language, GraphML) file format used is human readable. Free tools like yWorks yEd Graph Editor, Gephi and others can be used to easily create visual representations from this file.  
- Get-DependentRecipients_Output_Summary.txt  
  Number of initial recipients, number of additional recipients, number of total recipients, number of root cause mailbox permissions.

## Compare-RecipientPermissions.ps1
The script can be found in '`.\sample code\Compare-RecipientPermissions`'.

Compare two result files from Export-RecipientPermissions.ps1 to see how permissions have changed. This is extremely helpful for audit purposes or when you want to inform a group of users about their current permissions and how they changed over time.

Changes are marked in the column 'Change' with
- 'Deleted' if a line exists in the old file but not in the new one
- 'New' if a line exists in the new file but not in the old one
- 'Unchanged' if a line exists as well in the new file as in the old one

## FiltersAndSidhistory.ps1
The script can be found in '`.\sample code\other samples`'.

This sample code shows how to use [TrusteeFilter](#trusteefilter) to find permissions which may be affected by SIDHistory removal.

[GrantorFilter](#grantorfilter) behaves exactly like [TrusteeFilter](#trusteefilter), only the reference variable is $Grantor instead of $Trustee.

## MemberOfRecurse.ps1
The script can be found in '`.\sample code\other samples`'.

This sample code shows how to list the GUIDs of all groups a certain AD object is a member of. The script considers
- nested groups
- security groups, no matter of which group scope they are
- distribution groups (static only), no matter of which group scope they are
- group membership in trusted domains/forests (incl. nested groups)

These GUIDs can then be used to answer the question '[Which resources does a particular user or group have access to?](#which-resources-does-a-particular-user-or-group-have-access-to)'.
