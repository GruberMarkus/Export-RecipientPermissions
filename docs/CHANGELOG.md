<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">Export-RecipientPermissions</a>**<br>Document Exchange mailbox access rights, folder permissions, send as, send on behalf and managed by<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Changelog

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
