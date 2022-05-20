<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">Export-RecipientPermissions</a>**<br>Document Exchange mailbox access rights, folder permissions, "send as", "send on behalf", "managed by" and linked master accounts<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

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
		- [1.2.3. ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential](#123-exchangecredentialusernamefile-exchangecredentialpasswordfile-usedefaultcredential)
		- [1.2.4. ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal](#124-paralleljobsexchange-paralleljobsad-paralleljobslocal)
		- [1.2.5. GrantorFilter](#125-grantorfilter)
		- [1.2.6. ExportMailboxAccessRights](#126-exportmailboxaccessrights)
		- [1.2.7. ExportMailboxAccessRightsSelf](#127-exportmailboxaccessrightsself)
		- [1.2.8. ExportMailboxAccessRightsInherited](#128-exportmailboxaccessrightsinherited)
		- [1.2.9. ExportMailboxFolderPermissions](#129-exportmailboxfolderpermissions)
		- [1.2.10. ExportMailboxFolderPermissionsAnonymous](#1210-exportmailboxfolderpermissionsanonymous)
		- [1.2.11. ExportMailboxFolderPermissionsDefault](#1211-exportmailboxfolderpermissionsdefault)
		- [1.2.12. ExportMailboxFolderPermissionsOwnerAtLocal](#1212-exportmailboxfolderpermissionsowneratlocal)
		- [1.2.13. ExportMailboxFolderPermissionsMemberAtLocal](#1213-exportmailboxfolderpermissionsmemberatlocal)
		- [1.2.14. ExportMailboxFolderPermissionsExcludeFoldertype](#1214-exportmailboxfolderpermissionsexcludefoldertype)
		- [1.2.15. ExportSendAs](#1215-exportsendas)
		- [1.2.16. ExportSendAsSelf](#1216-exportsendasself)
		- [1.2.17. ExportSendOnBehalf](#1217-exportsendonbehalf)
		- [1.2.18. ExportManagedBy](#1218-exportmanagedby)
		- [1.2.19. ExportLinkedMasterAccount](#1219-exportlinkedmasteraccount)
		- [1.2.20. ExportTrustees](#1220-exporttrustees)
		- [1.2.21. ExportFile](#1221-exportfile)
		- [1.2.22. ErrorFile](#1222-errorfile)
		- [1.2.23. DebugFile](#1223-debugfile)
		- [1.2.24. UpdateInverval](#1224-updateinverval)
	- [1.3. Runtime](#13-runtime)
	- [1.4. Requirements](#14-requirements)
- [2. FAQ](#2-faq)
	- [2.1. Where can I find the changelog?](#21-where-can-i-find-the-changelog)
	- [2.2. How can I contribute, propose a new feature or file a bug?](#22-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
	- [2.3. How can I get more script output for troubleshooting?](#23-how-can-i-get-more-script-output-for-troubleshooting)
	- [2.4. A permission is reported, but the trustee details (primary SMTP address etc.) are empty](#24-a-permission-is-reported-but-the-trustee-details-primary-smtp-address-etc-are-empty)
	- [2.5. Isn't a plural noun in the script name against PowerShell best practices?](#25-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
	- [2.6. Is there a roadmap for future versions?](#26-is-there-a-roadmap-for-future-versions)
- [3. Get-DependentRecipients.ps1](#3-get-dependentrecipientsps1)
- [4. Compare-RecipientPermissions.ps1](#4-compare-recipientpermissionsps1)
- [5. Recommendations](#5-recommendations)

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
### 1.2.3. ExchangeCredentialUsernameFile, ExchangeCredentialPasswordFile, UseDefaultCredential
Credentials for Exchange connection
Username and password are stored as encrypted secure strings, if UseDefaultCredential is not enabled
### 1.2.4. ParallelJobsExchange, ParallelJobsAD, ParallelJobsLocal
Maximum Exchange, AD and local sessions/jobs running in parallel
Watch CPU and RAM usage, and your Exchange throttling policy
### 1.2.5. GrantorFilter
Grantors to consider
Only checks recipients that match the filter criteria. Only reduces the number of grantors, not the number of trustees.
Attributes that can filtered:
- .DistinguishedName
- .RecipientType, .RecipientTypeDetails
- .DisplayName
- .PrimarySmtpAddress: .Local, .Domain, .Address
- .EmailAddresses: .PrefixString, .IsPrimaryAddress, .SmtpAddress, .ProxyAddressString
- On-prem only: .Identity: .tostring() (CN), .DomainId, .Parent (parent CN)
Set to $null or '' to define all recipients as grantors to consider
Example: " `$Recipient.primarysmtpaddress.domain -ieq 'example.com'" },
### 1.2.6. ExportMailboxAccessRights
Rights set on the mailbox itself, such as "FullAccess" and "ReadAccess"
Default: $true
### 1.2.7. ExportMailboxAccessRightsSelf
Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)
Default: $false
### 1.2.8. ExportMailboxAccessRightsInherited
Report inherited mailbox access rights (only works on-prem)
Default: $false
### 1.2.9. ExportMailboxFolderPermissions
This part of the report can take very long
Default: $true
### 1.2.10. ExportMailboxFolderPermissionsAnonymous
Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)
Default: $false
### 1.2.11. ExportMailboxFolderPermissionsDefault
Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)
Default: $false
### 1.2.12. ExportMailboxFolderPermissionsOwnerAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.
Default: $false
### 1.2.13. ExportMailboxFolderPermissionsMemberAtLocal
Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.
Default: $false
### 1.2.14. ExportMailboxFolderPermissionsExcludeFoldertype
List of Foldertypes to ignore.

Some known folder types are: Audits, Calendar, CalendarLogging, CommunicatorHistory, Conflicts, Contacts, ConversationActions, DeletedItems, Drafts, ExternalContacts, Files, GalContacts, ImContactList, Inbox, Journal, JunkEmail, LocalFailures, Notes, Outbox, QuickContacts, RecipientCache, RecoverableItemsDeletions, RecoverableItemsPurges, RecoverableItemsRoot, RecoverableItemsVersions, Root, RssSubscription, SentItems, ServerFailures, SyncIssues, Tasks, WorkingSet, YammerFeeds, YammerInbound, YammerOutbound, YammerRoot

Default: 'audits'
### 1.2.15. ExportSendAs
Export Send As permissions
Default: $true
### 1.2.16. ExportSendAsSelf
Export Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)
Default: $false
### 1.2.17. ExportSendOnBehalf
Export Send On Behalf permissions
Default: $true
### 1.2.18. ExportManagedBy
Only for distribution groups, and not to be confused with the "Manager" attribute
Default: $true
### 1.2.19. ExportLinkedMasterAccount
Export Linked Master Account
Only works on-prem
Default: $true
### 1.2.20. ExportTrustees
Include all trustees in permission report file, only valid or only invalid ones
Valid trustees are trustees which can be resolved to an Exchange recipient
Valid values: 'All', 'OnlyValid', 'OnlyInvalid'
Default: 'All'
### 1.2.21. ExportFile
Name (and path) of the permission report file
Default: '.\export\Export-RecipientPermissions_Result.csv'
### 1.2.22. ErrorFile
Name (and path) of the error log file
Set to $null or '' to disable debugging
Default: '.\export\Export-RecipientPermissions_Error.csv',
### 1.2.23. DebugFile
Name (and path) of the debug log file
Set to $null or '' to disable debugging
Default: ''
### 1.2.24. UpdateInverval
Interval to update the job progress
Updates are based von recipients done, not on duration
Number must be 1 or higher, low numbers mean bigger debug files
Default: 100
## 1.3. Runtime
The script can run many hours, depending on the number of recipients and the speed of the environments to check.
Exporting mailbox folder permissions takes even more time because of how Exchange is designed to query these permissions.
## 1.4. Requirements
The script needs to be run with an account that has read permissions to all recipients in the cloud as well as Active Directory and Exchange on premises. The script asks for credentials.
As the credentials are stored in the encrypted secure string file format and can be re-used, the script can be fully automated and run as a scheduled job.

Per default, the script uses multiple parallel threads, each one consuming one Exchange PowerShell session. Please watch CPU and RAM usage, as wel as your Exchange throttling policy:
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
Start the script with the '-debug' and '-verbose' parameters to get the maximum output for troubleshooting.
## 2.4. A permission is reported, but the trustee details (primary SMTP address etc.) are empty
When excluding a bug in the script, there are three possible reasons why a trustee does not have details like a primary SMTP address in the result file:
- The trustee is a valid Active Directory object, but not a valid Exchange recipient.  
Examples: 'NT AUTHORITY\SELF', 'Domain Admins', 'Enterprise Admins'
- The initial trustee no longer exists.  
Exchange does not check if trustees still exist and remove the according permissions in case of deletion - this would be a problem when restoring deleted Active Directory objects.  
As Exchange stores trustees in different formats, the trustee original identity can be a SID, an NT4-style logon name or just about any string.
- Multiple recipients share the same linked master account, user friendly name, distinguished name, GUID or primary SMTP address. As the search for this value returns multiple recipients, no details are shown.  
This should not happen when using the built-in Exchange tools, due to their built-in quality checks. It happens more often, when Exchange attributes are modified directly in Active Directory.  
Whe passing the '-Debug' PowerShell parameter, the script outputs recipients with non-unique attributes in the debug stream.  
## 2.5. Isn't a plural noun in the script name against PowerShell best practices?
Absolutely. PowerShell best practices recommend using singular nouns, but Export-RecipientPermissions contains a plural noun.

I intentionally decided not to follow the singular noun convention, as another language as PowerShell was initially used for coding and the name of the tool was already defined. If this was a commercial enterprise project, marketing would have overruled development.
## 2.6. Is there a roadmap for future versions?
There is no binding roadmap for future versions, although I maintain a list of ideas in the 'Contribution opportunities' chapter of '.\docs\CONTRIBUTING.html'.

Fixing issues has priority over new features, of course.

# 3. Get-DependentRecipients.ps1
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
-	Get-DependentRecipients_Output_InitialRecipients.csv  
	List of initial recipients.
-	Get-DependentRecipients_Output_AdditionalRecipients.csv  
	List of additional recipients.
-	Get-DependentRecipients_Output_AllRecipients.csv  
	List of all initial and additional recipients.
-	Get-DependentRecipients_Output_GML.gml  
	All recipients and their permissions in a graphical representation. The GML (Graph Modeling Language) file format used is human readable. Free tools like yWorks yEd Graph Editor, Gephi and others can be used to easily create visual representations from this file.  
-	Get-DependentRecipients_Output_Summary.txt  
	Number of initial recipients, number of additional recipients, number of total recipients, number of root cause mailbox permissions.

# 4. Compare-RecipientPermissions.ps1
The script can be found in '`.\sample code\Compare-RecipientPermissions`'.

Compare two result files from Export-RecipientPermissions.ps1 to see which permissions have changed over time

Changes are marked in the column 'Change' with
- 'Deleted' if a line exists in the old file but not in the new one
- 'New' if a line exists in the new file but not in the old one
- 'Unchanged' if a line exists as well in the new file as in the old one

# 5. Recommendations
Make sure you have the latest updates installed to avoid memory leaks and CPU spikes (PowerShell, .Net framework).

If possible, allow Export-RecipientPermissions.ps1 to use your on premises infrastructure. This will dramatically increase the initial enumeration of recipients.

Start the script from PowerShell, not from within the PowerShell ISE. This makes especially Get-DependentMailboxes.ps1 run faster due to a different default thread apartment mode.

When running the scripts as scheduled job, make sure to include the "-ExecutionPolicy Bypass" parameter.
Example: `powershell.exe -ExecutionPolicy Bypass -file "c:\path\Export-RecipientPermissions.ps1"`