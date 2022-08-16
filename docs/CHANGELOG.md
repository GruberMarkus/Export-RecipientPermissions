<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">Export-RecipientPermissions</a>**<br>Document, filter and compare Exchange permissions: Mailbox Access Rights, Mailbox Folder permissions, Public Folder permissions, Send As, Send On Behalf, Managed By, Linked Master Accounts, Forwarders, Group members, Management Role Group members
<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

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

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/v2.x.x" target="_blank">v2.x.x</a> - YYYY-MM-DD
### Changed
- Importing recipients is now a multi-thred Exchange operation. Recipients are queried by lots of small queries to avoid problems with missing data in big environments.
### Added
- The new parameter '`ExpandGroups`' expands groups (including nested and dynamic groups) and exports the granted permission for each individual member. See '`README`' for details and comparison to '`ExportDistributionGroupMembers`'.
- The new parameter '`ExportDistributionGroupMembers`' exports distribution group members, including nested groups and dynamic groups. See '`README`' for details and comparison to '`ExpandGroups`'.
- The new parameter '`ExportFileFilter`' allows filtering the final results before they are written to the export file. See '`README`' for details.
- Special mailboxes are now added to the recipients list. This includes Arbitration, AuditLog, AuxAuditLog, inactive, Migration, Monitoring, RemoteArchive and softdeleted mailboxes (some of them are only available in on-prem or cloud environments)
- Mailbox permissions exported from the cloud now include softdeleted and unresolved trustees, as well as permissions granted to group mailboxes

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
