<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">Export-RecipientPermissions</a>**<br>Document Exchange mailbox access rights, folder permissions, "send as", "send on behalf", "managed by" and linked master accounts<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Changelog

## <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases/tag/vx.x.x" target="_blank">vx.x.x</a> - YYYY-MM-DD
###
- New default value '`$true`' for parameters '`ExportMailboxFolderPermissionsAnonymous`' and '`ExportMailboxFolderPermissionsDefault`'
### Added
- The new parameter '`RecipientProperties`' controls which recipient properties are loaded and can be used in '`GrantorFilter`' and '`TrusteeFilter`'. It also helps keep network traffic and memory usage low. See '`README`' for details.
- Mail-enabled public folders are now considered when exporting Send As and Send On Behalf permissions
- Support for export of public folder permissions. See '`README`' for details regarding the new parameters '`ExportPublicFolderPermissions`', '`ExportPublicFolderPermissionsAnonymous`', '`ExportPublicFolderPermissionsDefault`' and '`ExportPublicFolderPermissionsExcludeFoldertype`'.
- Support for export management role group permissiones. See '`README`' for new parameter '`ExportManagementRoleGroupMembers`'.
- Support for export of forwarders. See '`README`' for details regarding the '`ExportForwarders`' parameter.
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

## v0.1.0 - 2021-03-01
_First lines of code were written as proof of concept, but never published._
