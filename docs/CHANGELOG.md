<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="../src/logo/Export-RecipientPermissions%20Logo.png" width="400" title="Export-RecipientPermissions" alt="Export-RecipientPermissions"></a>**<br>Document, filter and compare Exchange permissions<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Changelog
<!--
Sample changelog entry
Remove leading spaces after pasting
## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/vX.X.X" target="_blank">vX.X.X</a> - YYYY-MM-DD
_Put Notice here_
_**Breaking:** Notice about breaking change_  
### Changed
- **Breaking:** XXX
### Added
### Removed
### Fixed
-->


## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v3.1.1" target="_blank">v3.1.1</a> - 2023-10-12
### Changed
- Update ExchangeOnlineManagement module to v3.4.0


## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v3.1.0" target="_blank">v3.1.0</a> - 2023-09-07
### Changed
- Update ExchangeOnlineManagement module to v3.3.0


## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v3.0.3" target="_blank">v3.0.3</a> - 2023-06-23
### Changed
- Update ExchangeOnlineManagement module to v3.2.0
### Fixed
- Indent of script output was not consistent across all tasks


## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v3.0.2" target="_blank">v3.0.2</a> - 2023-02-28
### Fixed
- ExportResourceDelegates was triggered by the ExportAcceptMessagesOnlyFrom parameter, not by ExportResourceDelegates
- ExpandGroups did not work because of assuming a wrong datatype which does not support the StartsWith method


## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v3.0.1" target="_blank">v3.0.1</a> - 2023-01-24
### Fixed
- Direct group members were only exported as GUIDs
- ExportGrantorsWithNoPermissions did not consider all distribution groups


## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v3.0.0" target="_blank">v3.0.0</a> - 2023-01-13
  _**Breaking:** Microsoft removes support for Remote PowerShell connections to Exchange Online starting June 1, 2023. See https://techcommunity.microsoft.com/t5/exchange-team-blog/announcing-deprecation-of-remote-powershell-rps-protocol-in/ba-p/3695597 for details.  
  Export-RecipientPermissions no longer uses Remote PowerShell to connect to Exchange Online and to Exchange on-premises. This brings some possibly breaking changes, which are detailed in the following release notes._
### Changed
- **Breaking:** Switching from Remote PowerShell session to local PowerShell session due Microsoft disabling Remote PowerShell in Exchange Online
  - Export-RecipientPermission will require more local resources (CPU, RAM, network) and will take longer to complete because operations previously handled on the server side now need to be handled on the client side
  - The variables '`$Grantor`' and '`$Trustee`' lose some sub attributes, so you may have to adopt your '`GrantorFilter`' and '`TrusteeFilter`' code:
    - '`.RecipientType.Value`' is now '`.RecipientType`'
    - '`.RecipientTypeDetails.Value`' is now '`.RecipientTypeDetails`'
    - '`.PrimarySmtpAddress`' no longer has the sub attributes .Local, .Domain and .Address
      - All the data formerly held in the sub attributes is still there, as .PrimarySmtpAddress is in the 'local@domain' format
    - '`.EmailAddresses`' (an array) no longer has the sub attributes .PrefixString, .IsPrimaryAddress, .SmtpAddress and .ProxyAddressString
      - All the data formerly held in the sub attributes is still there, as .EmailAddress is in the 'prefix:local@domain' format
    - On-prem only:
      - '`.Identity`' is now the canonical name (CN) only and no longer has the sub attributes .DomainId and .Parent
        - All the data formerly held in the sub attributes is still there, as .Identity is in the 'example.com/OU1/OU2/ObjectA' format  
  - Reduced the default value of '`ParallelJobsExchange`' from '`$ExchangeConnectionUriList.count * 3`' to '`$ExchangeConnectionUriList.count`' as local Exchange PowerShell sessions are not as stable as Remote PowerShell sessions 
- Adopted sample code and documentation to reflect changes in the '`$Grantor`' and '`$Trustee`' variables
- Use Get-EXO* cmdlets in Exchange Online where possible
- Upgrade to ExchangeOnlineManagement v3.1.0
### Added
- New export parameters: '`ExportModerators`', '`ExportRequireAllSendersAreAuthenticated`', '`ExportAcceptMessagesOnlyFrom`', '`ExportResourceDelegates`'. See '`README`' for details.
- New FAQ in '`README`': 'I receive an error message when connecting to Exchange on premises'
### Fixed
- When importing UserFriendlyNames, errors were not written to the error file because of a missing encoding variable
- Only the first ManagedBy entry for reach recipient was exported correctly, the following entries were missing trustee data 

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.3.1" target="_blank">v2.3.1</a> - 2022-11-28
### Added
- New FAQ in '`README`': 'How to export permissions for specific public folders?'
### Fixed
- Sample code '`compare.ps1`' now additionally outputs the original identity of a trustee and not only the primary SMTP address. This helps with permissions granted to 'Anonymous' and 'Default', as well as with recipients which have been deleted in the time between the old and the new export.
- Always include trustee groups in '`GrantorFilter`' when '`ExportDistributionGroups`' is set to '`OnlyTrustees`'

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.3.0" target="_blank">v2.3.0</a> - 2022-10-25
### Added
- When '`ExportFromOnPrem`' is set to '`$true`' and '`ExchangeConnectionUriList`' is not specified, '`ExchangeConnectionUriList`' defaults to '`http://<server>/powershell`' for each Exchange server with the mailbox server role
- New FAQs in '`README`': 'Which resources does a particular user or group have access to?', 'How to find distribution lists without members?'
- New sample code 'MemberOfRecurse.ps1'

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.2.1" target="_blank">v2.2.1</a> - 2022-09-27
### Fixed
- When ExportGrantorsWithNoPermissions is enabled and ExportGuids is disabled, empty management role groups were exported with no name and a trailing slash in the recipient type field
- Some major tasks did not show a timestamp, but rather the PowerShell code generating the timestamp
- ExpandGroups did not work for groups which were not part of the grantor filter
- Mailbox access rights now correctly show data for trustees identified via linked master account

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.2.0" target="_blank">v2.2.0</a> - 2022-09-21
### Changed
- Updated '`README`' to correctly document value names (.RecipientType.Value, .RecipientTypeDetails.Value) (<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues/14" target="_blank">#14</a>) (Thanks <a href="https://github.com/snurbnacnud" target="_blank">@snurbnacnud</a>!)
- Bump ExchangeOnlineMangement to v3.0.0
### Added
- New parameter '`ExportGrantorsWithNoPermissions`', see '`README`' for details
- New parameter '`ExportGuids`', see '`README`' for details
- New parameter '`ExportGroupMembersRecurse`', see '`README`' for details
- New FAQ in '`README`': Is there a GUI available?
- Parallelize combination of temporary result files to final result file
- Parallelize import of direct group members
- Export-RecipientPermissions now has a logo and an icon
### Fixed
- Incorrect escape of double quotes in CSV files

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.1.0" target="_blank">v2.1.0</a> - 2022-09-05
### Changed
- Importing recipients is now a multi-thred Exchange operation. Recipients are queried by lots of small queries to avoid problems with missing data in big environments.
### Added
- New parameter '`ExpandGroups`': If a trustee is a group, get the group's recursive members and export the granted permission for each individual member. See '`README`' for details and comparison to '`ExportDistributionGroupMembers`'.
- New parameter '`ExportDistributionGroupMembers`': Export recursive distribution group members, including nested groups and dynamic groups. See '`README`' for details and comparison to '`ExpandGroups`'.
- The new parameter '`ExportFileFilter`' allows filtering the final results before they are written to the export file. See '`README`' for details.
- Special mailboxes are now added to the recipients list. This includes Arbitration, AuditLog, AuxAuditLog, inactive, Migration, Monitoring, RemoteArchive and softdeleted mailboxes (some of them are only available in on-prem or cloud environments)
- Mailbox permissions exported from the cloud now include softdeleted and unresolved trustees

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.0.0" target="_blank">v2.0.0</a> - 2022-07-19
 _**Breaking:** See '`Changed`' section for breaking changes_  
### Changed
- **Breaking:** New default values for several parameters:
  - '`ExportMailboxFolderPermissions`': '`$false`'
  - '`ExportMailboxFolderPermissionsAnonymous`': '`$true`'
  - '`ExportMailboxFolderPermissionsDefault`': '`$true`'
- The GrantorFilter parameter can now only use the reference variable '`$Grantor`' and no longer '`$Recipient`'. This change has been announced with the release v1.5.0.
### Added
- The new parameter '`RecipientProperties`' controls which recipient properties are loaded and can then be used in '`GrantorFilter`' and '`TrusteeFilter`'. It also helps keep network traffic and memory usage low. See '`README`' for details.
- Mail-enabled public folders are now considered when exporting Send As and Send On Behalf permissions
- Export of public folder permissions. See '`README`' for details regarding the new parameters '`ExportPublicFolderPermissions`', '`ExportPublicFolderPermissionsAnonymous`', '`ExportPublicFolderPermissionsDefault`' and '`ExportPublicFolderPermissionsExcludeFoldertype`'.
- Export management role group permissions. See '`README`' for new parameter '`ExportManagementRoleGroupMembers`'.
- Export forwarders. See '`README`' for details regarding the '`ExportForwarders`' parameter.
- The sample file '`Export-RecipientPermissions_Result.csv`' shows a typical result file of Export-RecipientPermissions
- New FAQs 'Which permissions are required?' and 'Can the script resolve permissions granted to a group to it's individual members?' in '`README`'
### Fixed
- Export all Send As permissions, not only the one granted by the last recipient checked by each parallel job

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.5.1" target="_blank">v1.5.1</a> - 2022-06-27
### Fixed
- Make sure non-working Exchange Online connections are properly closed and do not remain active in the background

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.5.0" target="_blank">v1.5.0</a> - 2022-06-25
### Changed
- The GrantorFilter parameter can now use the new reference variable '$Grantor' in addition to '$Recipient'. Both reference variables have identical values. '$Recipient' is now marked as obsolete and may removed in a future release.
### Added
- New parameter 'TrusteeFilter': Filter the trustees included in the export. See 'README' for details.
- New sample code 'FiltersAndSidhistory.ps1' shows how to use TrusteeFilter and GrantorFilter to find permissions which may be affected by SIDHistory removal.
- The connection to the cloud now uses the Exchange Online PowerShell V2 module (a preview version is used which allows traditional remote PowerShell access to all cmdlets). See '.\README' for details about the required 'ExchangeOnlineConnnectionParameters' parameter.

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.4.0" target="_blank">v1.4.0</a> - 2022-05-24
### Added
- New parameter 'ExportLinkedMasterAccount': Export the configured linked master account of mailboxes as permission. See 'README' for details.
- Permissions granted to a trustee which has a linked master account are now resolved against the list of recipients
- New parameter 'ExportTrustees': Export all trustees, or only those can can or can not be resolved to a recipient. See 'README' for details.
- Non-unique LinkedMasterAccounts, UserFriendlyNames, DistinguishedNames, GUIDs and PrimarySmtpAddresses are shown in verbose stream. When one of these attributes is not unique, a trustee can not be matched against just one recipient, so the corresponding details in the report are empty.
- Added automatic retry of last command not only in case of an error related to the command itself, but also when the error is related to the underlying Exchange connection
- Added FAQs in 'README' file
### Fixed
- The mailbox GUID is no longer always blank in the error message indicating a connection error to a mailbox

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.3.0" target="_blank">v1.3.0</a> - 2022-03-06
### Added
- New parameter 'UseDefaultCredential'

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.2.0" target="_blank">v1.2.0</a> - 2022-03-03
### Added
- New parameter 'ExportMailboxFolderPermissionsExcludeFoldertype'
- Separate error file, configuration via 'ErrorFile' parameter
- Comment based help
### Changed
- Sample code directories now match the names of the contained scripts (Get-DependentRecipients, Compare-RecipientPermissions)
- Massive performance gains in sample code Get-DependentRecipients
- Detecting root mailbox folder now uses the 'FolterType' property
- The debug file is no longer enabled per default, in favor of the error file. This can be changed with the 'DebugFile' parameter.
- Encode CSV file in UTF8 with BOM instead of UTF8 without BOM, so that Excel detects the file format correctly

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.1.0" target="_blank">v1.1.0</a> - 2022-02-23
### Added
- Massive performance improvements, less RAM and network usage
- Documentation
- Command-line parameters
- Sample code: Permission changes over time
- Sample code: Get dependent recipients to support migrations to Exchange Online as not all permissions work cross-premises, or to graphically document existing permissions

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v1.0.0" target="_blank">v1.0.0</a> - 2019-04-04
_Initial release._

## v0.1.0 - 2018-01-01
_First lines of code were written as proof of concept, but never published._
