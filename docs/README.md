<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">Export-RecipientPermissions</a>**<br>Document Exchange mailbox access rights, folder permissions, "send as", "send on behalf", "managed by" and linked master accounts<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
Finds all recipients with a primary SMTP address in an on on-prem or online Exchange environment and documents their
- mailbox access rights,
- mailbox folder permissions,
- "send as" permissions,
- "send on behalf" permissions,
- "managed by" permissions,
- and linked master accounts

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
		- [1.2.6. GrantorFilter](#126-grantorfilter)
		- [1.2.7. TrusteeFilter](#127-trusteefilter)
		- [1.2.8. ExportMailboxAccessRights](#128-exportmailboxaccessrights)
		- [1.2.9. ExportMailboxAccessRightsSelf](#129-exportmailboxaccessrightsself)
		- [1.2.10. ExportMailboxAccessRightsInherited](#1210-exportmailboxaccessrightsinherited)
		- [1.2.11. ExportMailboxFolderPermissions](#1211-exportmailboxfolderpermissions)
		- [1.2.12. ExportMailboxFolderPermissionsAnonymous](#1212-exportmailboxfolderpermissionsanonymous)
		- [1.2.13. ExportMailboxFolderPermissionsDefault](#1213-exportmailboxfolderpermissionsdefault)
		- [1.2.14. ExportMailboxFolderPermissionsOwnerAtLocal](#1214-exportmailboxfolderpermissionsowneratlocal)
		- [1.2.15. ExportMailboxFolderPermissionsMemberAtLocal](#1215-exportmailboxfolderpermissionsmemberatlocal)
		- [1.2.16. ExportMailboxFolderPermissionsExcludeFoldertype](#1216-exportmailboxfolderpermissionsexcludefoldertype)
		- [1.2.17. ExportSendAs](#1217-exportsendas)
		- [1.2.18. ExportSendAsSelf](#1218-exportsendasself)
		- [1.2.19. ExportSendOnBehalf](#1219-exportsendonbehalf)
		- [1.2.20. ExportManagedBy](#1220-exportmanagedby)
		- [1.2.21. ExportLinkedMasterAccount](#1221-exportlinkedmasteraccount)
		- [1.2.22. ExportTrustees](#1222-exporttrustees)
		- [1.2.23. ExportFile](#1223-exportfile)
		- [1.2.24. ErrorFile](#1224-errorfile)
		- [1.2.25. DebugFile](#1225-debugfile)
		- [1.2.26. UpdateInverval](#1226-updateinverval)
	- [1.3. Runtime](#13-runtime)
	- [1.4. Requirements](#14-requirements)
- [2. FAQ](#2-faq)
	- [2.1. Where can I find the changelog?](#21-where-can-i-find-the-changelog)
	- [2.2. How can I contribute, propose a new feature or file a bug?](#22-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
	- [2.3. How can I get more script output for troubleshooting?](#23-how-can-i-get-more-script-output-for-troubleshooting)
	- [2.4. A permission is reported, but the trustee details (primary SMTP address etc.) are empty](#24-a-permission-is-reported-but-the-trustee-details-primary-smtp-address-etc-are-empty)
	- [2.5. Isn't a plural noun in the script name against PowerShell best practices?](#25-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
	- [2.6. Is there a roadmap for future versions?](#26-is-there-a-roadmap-for-future-versions)
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
- Grantor Primary SMTP: The primary SMTP address of the object granting a permission.
- Grantor Display Name: The display name of the grantor.
- Grantor Recipient Type: The recipient type and recipient type detail of the grantor.
- Grantor Environment: Shows if the grantor is held on-prem or in the cloud.
- Folder: Mailbox folder the permission is granted on. Empty for non-folder permissions.
- Permission: The permission granted/received (e.g., FullAccess, SendAs, SendOnBehalf etc.)
- Allow/Deny: Shows if the permission is an allow or a deny permission.
- Inherited: Shows if the permission is inherited or not.
- InheritanceType: Shows if the permission is also valid for child objects, and if yes, which child objects.
- Trustee Original Identity: The original identity string of the trustee.
- Trustee Primary SMTP: The primary SMTP address of the object receiving a permission.
- Trustee Display Name: The display name of the trustee.
- Trustee Recipient Type: The recipient type of the trustee.
- Trustee Environment: Shows if the trustee is held on-prem or in the cloud.
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
### 1.2.6. GrantorFilter
Only check grantors where the filter criteria matches $true.

Attributes of the variable $Grantor that can be filtered:
- .DistinguishedName
- .RecipientType, .RecipientTypeDetails
- .DisplayName
- .PrimarySmtpAddress: .Local, .Domain, .Address
- .EmailAddresses: .PrefixString, .IsPrimaryAddress, .SmtpAddress, .ProxyAddressString
  - This attribute is an array. Code example:
    ```
    $GrantorFilter = "foreach (`$XXXSingleSmtpAddressXXX in `$Grantor.EmailAddresses.SmtpAddress) { if (`$XXXSingleSmtpAddressXXX -iin @(
                      'addressA@example.com’,
                      'addressB@example.com’
      )) { `$true; break } }"
    ```
- .UserFriendlyName: User account holding the mailbox in the `"<NetBIOS domain name>\<sAMAccountName>"` format
- .ManagedBy: .Rdn, .Parent, .DistinguishedName, .DomainId, .Name
  - This attribute is an array. Code example:
    ```
    $GrantorFilter = "foreach (`$XXXSingleManagedByXXX in `$Grantor.ManagedBy) { if (`$XXXSingleManagedByXXX -iin @(
                          'example.com/OU1/OU2/ObjectA’,
                          'example.com/OU3/OU4/ObjectB’,
      )) { `$true; break } }"
    ```
- On-prem only:
  - .Identity: .tostring() (CN), .DomainId, .Parent (parent CN)
  - .LinkedMasterAccount: Linked Master Account in the "<NetBIOS domain name>\<sAMAccountName>" format

Set to \$null or '' to define all recipients as grantors to consider

Example:
    ```
    "`$Grantor.primarysmtpaddress.domain -ieq 'example.com'"
    ```

Default: $null
### 1.2.7. TrusteeFilter
Only report trustees where the filter criteria matches $true.

If the trustee matches a recipient, the available attributes are the same as für GrantorFilter, only the reference variable is $Trustee instead of $Grantor.

If the trustee does not match a recipient (because it no longer exists, for exampe), $Trustee is just a string. In this case, the export shows the following:
- Column "Trustee Original Identity" contains the trustee description string as reported by Exchange
- Columns "Trustee Primary SMTP" and "Trustee Display Name" are empty

Example:
    ```
    "`$Trustee.primarysmtpaddress.domain -ieq 'example.com'"
    ```

Default: $null
### 1.2.8. ExportMailboxAccessRights
Rights set on the mailbox itself, such as "FullAccess" and "ReadAccess"

Default: $true
### 1.2.9. ExportMailboxAccessRightsSelf
Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)

Default: $false
### 1.2.10. ExportMailboxAccessRightsInherited
Report inherited mailbox access rights (only works on-prem)

Default: $false
### 1.2.11. ExportMailboxFolderPermissions
This part of the report can take very long

Default: $true
### 1.2.12. ExportMailboxFolderPermissionsAnonymous
Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)

Default: $false
### 1.2.13. ExportMailboxFolderPermissionsDefault
Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)

Default: $false
### 1.2.14. ExportMailboxFolderPermissionsOwnerAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.

Default: $false
### 1.2.15. ExportMailboxFolderPermissionsMemberAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.
Default: $false
### 1.2.16. ExportMailboxFolderPermissionsExcludeFoldertype
List of Foldertypes to ignore.

Some known folder types are: Audits, Calendar, CalendarLogging, CommunicatorHistory, Conflicts, Contacts, ConversationActions, DeletedItems, Drafts, ExternalContacts, Files, GalContacts, ImContactList, Inbox, Journal, JunkEmail, LocalFailures, Notes, Outbox, QuickContacts, RecipientCache, RecoverableItemsDeletions, RecoverableItemsPurges, RecoverableItemsRoot, RecoverableItemsVersions, Root, RssSubscription, SentItems, ServerFailures, SyncIssues, Tasks, WorkingSet, YammerFeeds, YammerInbound, YammerOutbound, YammerRoot

Default: 'audits'
### 1.2.17. ExportSendAs
Export Send As permissions

Default: $true
### 1.2.18. ExportSendAsSelf
Export Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)

Default: $false
### 1.2.19. ExportSendOnBehalf
Export Send On Behalf permissions

Default: $true
### 1.2.20. ExportManagedBy
Only for distribution groups, and not to be confused with the "Manager" attribute

Default: $true
### 1.2.21. ExportLinkedMasterAccount
Export Linked Master Account

Only works on-prem

Default: $true
### 1.2.22. ExportTrustees
Include all trustees in permission report file, only valid or only invalid ones

Valid trustees are trustees which can be resolved to an Exchange recipient

Valid values: 'All', 'OnlyValid', 'OnlyInvalid'

Default: 'All'
### 1.2.23. ExportFile
Name (and path) of the permission report file

Default: '.\export\Export-RecipientPermissions_Result.csv'
### 1.2.24. ErrorFile
Name (and path) of the error log file

Set to $null or '' to disable debugging

Default: '.\export\Export-RecipientPermissions_Error.csv',
### 1.2.25. DebugFile
Name (and path) of the debug log file

Set to $null or '' to disable debugging

Default: ''
### 1.2.26. UpdateInverval
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
## 2.1. Where can I find the changelog?
The changelog is located in the `'.\docs'` folder, along with other documents related to Set-OutlookSignatures.
## 2.2. How can I contribute, propose a new feature or file a bug?
If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, please have a look at `'.\docs\CONTRIBUTING'` for a rough overview of the proposed process.
## 2.3. How can I get more script output for troubleshooting?
Start the script with the '-verbose' parameter to get the maximum output for troubleshooting.
## 2.4. A permission is reported, but the trustee details (primary SMTP address etc.) are empty
When excluding a bug in the script, there are three possible reasons why a trustee does not have details like a primary SMTP address in the result file:
- The trustee is a valid Active Directory object, but not a valid Exchange recipient.  
Examples: 'NT AUTHORITY\SELF', 'Domain Admins', 'Enterprise Admins'
- The initial trustee no longer exists.  
Exchange does not check if trustees still exist and remove the according permissions in case of deletion - this would be a problem when restoring deleted Active Directory objects.  
As Exchange stores trustees in different formats, the trustee original identity can be a SID, an NT4-style logon name or just about any string.
- Multiple recipients share the same linked master account, user friendly name, distinguished name, GUID or primary SMTP address. As the search for this value returns multiple recipients, no details are shown.  
This should not happen when using the built-in Exchange tools, due to their built-in quality checks. It happens more often, when Exchange attributes are modified directly in Active Directory.  
Whe passing the '-verbose' PowerShell parameter, the script outputs recipients with non-unique attributes in the verbose stream.  
## 2.5. Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Export-RecipientPermissions contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.
## 2.6. Is there a roadmap for future versions?
There is no binding roadmap for future versions, although I maintain a list of ideas in the 'Contribution opportunities' chapter of '.\docs\CONTRIBUTING.html'.

Fixing issues has priority over new features, of course.
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