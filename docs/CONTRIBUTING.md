<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">Export-RecipientPermissions</a>**<br>Document, filter and compare Exchange permissions: Mailbox Access Rights, Mailbox Folder permissions, Public Folder permissions, Send As, Send On Behalf, Managed By, Linked Master Accounts, Forwarders, Group members, Management Role Group members
<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Welcome! <!-- omit in toc -->
Thank you very much for your interest in Export-RecipientPermissions.

If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, this document gives you a rough overview of the proposed process.  
I'm not a professional developer - if you are one and you notice something negative in the code or process, please let me know.
# Table of Contents <!-- omit in toc -->
- [1. Code of Conduct](#1-code-of-conduct)
- [2. Contribution opportunities](#2-contribution-opportunities)
  - [2.1. Sponsoring](#21-sponsoring)
  - [2.2. Code refactoring](#22-code-refactoring)
  - [2.3. Combine on-prem and cloud results](#23-combine-on-prem-and-cloud-results)
  - [2.4. Adopt Exchange Online connection method](#24-adopt-exchange-online-connection-method)
- [3. Branches](#3-branches)
- [4. Development process](#4-development-process)
- [5. Commit messages](#5-commit-messages)
  - [5.1. Commit message format](#51-commit-message-format)
  - [5.2. Type](#52-type)
  - [5.3. Body](#53-body)
  - [5.4. Footer](#54-footer)
- [6. Build process](#6-build-process)
# 1. Code of Conduct
When contributing to Export-RecipientPermissions, please make sure to follow the Code of Conduct ('CODE_OF_CONDUCT.html' in the same directory as this document).
# 2. Contribution opportunities
## 2.1. Sponsoring
Being Free and Open-Source Software, Export-RecipientPermissions can save a business thousands or even tens of thousand Euros/US-Dollars per year in comparison to commercial software.  
Please consider <a href="https://github.com/sponsors/GruberMarkus" target="_blank">sponsoring this project</a> to ensure continued support, testing and enhancements.

Dear businesses, please don't forget:
- Invest in the open-source projects you depend on. Contributors are working behind the scenes to make open-source better for everyone - give them the help and recognition they deserve.
- Sponsor the open-source software your team has built its business on. Fund the projects that make up your software supply chain to improve its performance, reliability, and stability.
## 2.2. Code refactoring
I'm not a professional developer, but a hobbyist scripter, and the code looks like that.

There are optimization opportunities in error handling, de-duplicating code with functions, applying PowerShell best practices, and more.
## 2.3. Combine on-prem and cloud results
Currently, Export-RecipientPermissions connects either to an on-prem instance or a cloud instance of Exchange.

A future release will be able to connect to both environments and combine these two data sources.
## 2.4. Adopt Exchange Online connection method
The script still connects to Exchange Online by using the classic New-PSSession construct.

Connect-ExchangeOnline clearly is the future. Unfortunately, it's connections to Exchange Online have shown to be error prone when used in parallel jobs, even for small environments.

As soon as these problems are solved, Export-RecipientPermissions will switch to Connect-ExchangeOnline.
# 3. Branches
The default branch is named '`main`'. It contains the source of the latest stable release.

Tags in the '`main`' branch mark releases - we release on tags (and therefore commits) and not on branches.
# 4. Development process
1. Create a new branch based on '`main`'
   - Hotfix: Give the branch a name starting with '`hotfix-`', e. g. '`hotfix-1.2.3`' or '`hotfix-issue13`'
   - New feature or vNext: Give the branch a name starting with '`develop-'`, e. g. '`develop-1.3.0`' or '`develop-vNext`'
2. Work on the code in the new branch. Every commit to the GitHub repository will trigger the build process and create a draft release.  
You can commit to the new branch as often as you like, and you don't have to care about commit messages during development and testing - the commit messages from the dev branch will not appear in the '`main`' branch, as we will do a "squash and merge" later.
3. When development is done, update the changelog.<br>Don't forget the link to the new release at the end of the file.
4. Create a pull request to incorporate the '`hotfix-`' or '`develop-`' changes into the '`main`' branch.
5. Discuss and review the pull request.
6. When applying the pull request, use `"squash and merge"` as it helps keep the commit history in the '`main`' branch clean and allows the developer to focus on, well, development.
7. After the pull request has been committed to the '`main`' branch, delete the now obsolete '`hotfix-`' or '`develop-`' branch.
8. If there are other '`hotfix-`' or '`develop-`' branches, they have to be rebased to the '`main`' branch which is now at least one commit ahead.
# 5. Commit messages
Commit messages should follow the <a href="https://www.conventionalcommits.org" target="_blank">Conventional Commits</a> standard.
## 5.1. Commit message format
```
<type>[optional scope]: <short description>
<blank line>
[optional body]
<blank line>
[optional footer]
```
## 5.2. Type
- Type is mandatory. 
- '`fix[optional scope]:`' A fix. Bumps SemVer patch version.
- '`feat[optional scope]:`' Introduces a new feature to the codebase. Bumps SemVer minor version.
- Other commit types other than '`fix:`' and '`feat:`' are allowed, e. g. '`build:`', '`chore:`', '`ci:`', '`docs:`', '`perf:`', '`refactor:`', '`revert:`', '`style:`' and '`test:`'.
- A scope may be provided to a commit's type, to provide additional contextual information and is contained within parenthesis, e.g. '`feat(parser): add ability to parse arrays`'.
## 5.3. Body
- Body is optional.
- Provide additional contextual information about the code changes. The body must begin one blank line after the description.
- '`BREAKING CHANGE:<blank>`' at the beginning of the optional body or footer section introduces a breaking API change (bumps SemVer major version). A breaking change can be part of commits of any type.
## 5.4. Footer
- Footer is optional.
- The footer should contain additional issue references about the code changes (such as the issues it fixes, e.g. '`Fixes [#13](https://github.com/GruberMarkus/Export-RecipientPermissions/issues/13) ([@GruberMarkus](https://github.com/GruberMarkus))`'.
- Text describing further details.
- '`BREAKING CHANGE:<blank>`' at the beginning of the optional body or footer section introduces a breaking API change (bumpgs SemVer major version). A commit of any type can contain a breaking change.
# 6. Build process
Every single commit in any branch or setting a tag starting with '`v`' triggers a build and the creation of a draft release.

The draft release includes the build artifact(s), the corresponding changelog entry and file hash information.

The build artifacts can be downloaded and go through the final test process.
- If these final tests are passed and the information in the draft build is correct, the build can directly be released.  
- If these final tests fail or the information in the draft release is wrong, delete the draft release and go on with with development process.

The build process is built on GitHub Actions workflows and currently only works in this environment.