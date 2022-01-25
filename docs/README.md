<!-- omit in toc -->
# <a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank">**Export-RecipientPermissions**</a><br>Document mailbox	access rights and folder permissions, "send as", "send on behalf" and "managed by"<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
**Signatures and OOF messages can be:**
- Generated from templates in DOCX or HTML file format  
- Customized with a broad range of variables, including photos, from Active Directory and other sources  
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific e-mail addresses, for every primary mailbox across all Outlook profiles  
- Assigned time ranges within which they are valid  
- Set as default signature for new e-mails, or for replies and forwards (signatures only)  
- Set as default OOF message for internal or external recipients (OOF messages only)  
- Set in Outlook Web for the currently logged in user  
- Centrally managed only or exist along user created signatures (signatures only)  
- Copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only)

Export-RecipientPermissions can be **executed by users on clients, or on a server without end user interaction**.  
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.  
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

**Sample templates** for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

**Simulation mode** allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.
  
The script is **designed to work in big and complex environments** (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works **on premises, in hybrid and cloud-only environments**.

It is **multi-client capable** by using different template paths, configuration files and script parameters.

Set-OutlookSignature requires **no installation on servers or clients**. You only need a standard file share on a server, and PowerShell and Office. 

A **documented implementation approach**, based on real life experiences implementing the script in a multi-client environment with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators.  
The implementatin approach is **suited for service providers as well as for clients**, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The script is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `'.\docs\LICENSE.txt'` for copyright and MIT license details.
<br><br>
**Dear businesses using Export-RecipientPermissions:**
- Being Free and Open-Source Software, Export-RecipientPermissions can save you thousands or even tens of thousand Euros/US-Dollars per year in comparison to commercial software.  
Please consider <a href="https://github.com/sponsors/GruberMarkus" target="_blank">sponsoring this project</a> to ensure continued support, testing and enhancements.
- Invest in the open-source projects you depend on. Contributors are working behind the scenes to make open-source better for everyone - give them the help and recognition they deserve.
- Sponsor the open-source software your team has built its business on. Fund the projects that make up your software supply chain to improve its performance, reliability, and stability.
# Table of Contents <!-- omit in toc -->
- [1. Requirements](#1-requirements)
- [2. Parameters](#2-parameters)
  - [2.1. SignatureTemplatePath](#21-signaturetemplatepath)
  - [2.2. SignatureIniPath](#22-signatureinipath)
  - [2.3. ReplacementVariableConfigFile](#23-replacementvariableconfigfile)
  - [2.4. GraphConfigFile](#24-graphconfigfile)
  - [2.5. TrustedDomainsToCheckForGroups](#25-trusteddomainstocheckforgroups)
  - [2.6. DeleteUserCreatedSignatures](#26-deleteusercreatedsignatures)
  - [2.7. DeleteScriptCreatedSignaturesWithoutTemplate](#27-deletescriptcreatedsignatureswithouttemplate)
  - [2.8. SetCurrentUserOutlookWebSignature](#28-setcurrentuseroutlookwebsignature)
  - [2.9. SetCurrentUserOOFMessage](#29-setcurrentuseroofmessage)
  - [2.10. OOFTemplatePath](#210-ooftemplatepath)
  - [2.11. OOFIniPath](#211-oofinipath)
  - [2.12. AdditionalSignaturePath](#212-additionalsignaturepath)
  - [2.13. AdditionalSignaturePathFolder](#213-additionalsignaturepathfolder)
  - [2.14. UseHtmTemplates](#214-usehtmtemplates)
  - [2.15. SimulateUser](#215-simulateuser)
  - [2.16. SimulateMailboxes](#216-simulatemailboxes)
  - [2.17. GraphCredentialFile](#217-graphcredentialfile)
  - [2.18. GraphOnly](#218-graphonly)
  - [2.19. CreateRTFSignatures](#219-creatertfsignatures)
  - [2.20. CreateTXTSignatures](#220-createtxtsignatures)
  - [2.21. EmbedImagesInHTML](#221-embedimagesinhtml)
- [3. Outlook signature path](#3-outlook-signature-path)
- [4. Mailboxes](#4-mailboxes)
- [5. Group membership](#5-group-membership)
- [6. Removing old signatures](#6-removing-old-signatures)
- [7. Error handling](#7-error-handling)
- [8. Run script while Outlook is running](#8-run-script-while-outlook-is-running)
- [9. Signature and OOF file format](#9-signature-and-oof-file-format)
  - [9.1. Signature and OOF file naming](#91-signature-and-oof-file-naming)
- [10. Tags and ini files](#10-tags-and-ini-files)
  - [10.1. Allowed tags](#101-allowed-tags)
  - [10.2. How to work with ini files](#102-how-to-work-with-ini-files)
- [11. Signature and OOF application order](#11-signature-and-oof-application-order)
- [12. Variable replacement](#12-variable-replacement)
  - [12.1. Photos from Active Directory](#121-photos-from-active-directory)
- [13. Outlook Web](#13-outlook-web)
- [14. Hybrid and cloud-only support](#14-hybrid-and-cloud-only-support)
  - [14.1. Basic Configuration](#141-basic-configuration)
  - [14.2. Advanced Configuration](#142-advanced-configuration)
  - [14.3. Authentication](#143-authentication)
- [15. Simulation mode](#15-simulation-mode)
- [16. FAQ](#16-faq)
  - [16.1. Where can I find the changelog?](#161-where-can-i-find-the-changelog)
  - [16.2. How can I contribute, propose a new feature or file a bug?](#162-how-can-i-contribute-propose-a-new-feature-or-file-a-bug)
  - [16.3. How is the account of a mailbox identified?](#163-how-is-the-account-of-a-mailbox-identified)
  - [16.4. How is the personal mailbox of the currently logged in user identified?](#164-how-is-the-personal-mailbox-of-the-currently-logged-in-user-identified)
  - [16.5. Which ports are required?](#165-which-ports-are-required)
  - [16.6. Why is Out of Office abbreviated OOF and not OOO?](#166-why-is-out-of-office-abbreviated-oof-and-not-ooo)
  - [16.7. Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates.](#167-should-i-use-docx-or-htm-as-file-format-for-templates-signatures-in-outlook-sometimes-look-different-than-my-templates)
  - [16.8. How can I log the script output?](#168-how-can-i-log-the-script-output)
  - [16.9. Can multiple script instances run in parallel?](#169-can-multiple-script-instances-run-in-parallel)
  - [16.10. How do I start the script from the command line or a scheduled task?](#1610-how-do-i-start-the-script-from-the-command-line-or-a-scheduled-task)
  - [16.11. How to create a shortcut to the script with parameters?](#1611-how-to-create-a-shortcut-to-the-script-with-parameters)
  - [16.12. What is the recommended approach for implementing the software?](#1612-what-is-the-recommended-approach-for-implementing-the-software)
  - [16.13. What is the recommended approach for custom configuration files?](#1613-what-is-the-recommended-approach-for-custom-configuration-files)
  - [16.14. Isn't a plural noun in the script name against PowerShell best practices?](#1614-isnt-a-plural-noun-in-the-script-name-against-powershell-best-practices)
  - [16.15. The script hangs at HTM/RTF export, Word shows a security warning!?](#1615-the-script-hangs-at-htmrtf-export-word-shows-a-security-warning)
  - [16.16. How to avoid empty lines when replacement variables return an empty string?](#1616-how-to-avoid-empty-lines-when-replacement-variables-return-an-empty-string)
  - [16.17. Is there a roadmap for future versions?](#1617-is-there-a-roadmap-for-future-versions)
  - [16.18. How to deploy signatures for "Send As", "Send On Behalf" etc.?](#1618-how-to-deploy-signatures-for-send-as-send-on-behalf-etc)
  - [16.19. Can I centrally manage and deploy Outook stationery with this script?](#1619-can-i-centrally-manage-and-deploy-outook-stationery-with-this-script)
  - [16.20. Why is membership in dynamic distribution groups and dynamic security groups not considered?](#1620-why-is-membership-in-dynamic-distribution-groups-and-dynamic-security-groups-not-considered)
    - [16.20.1. What's the alternative to dynamic groups?](#16201-whats-the-alternative-to-dynamic-groups)
  - [16.21. Why is no admin or user GUI available?](#1621-why-is-no-admin-or-user-gui-available)
  - [16.22. What about the new signature roaming feature Microsoft announced?](#1622-what-about-the-new-signature-roaming-feature-microsoft-announced)
    - [16.22.1. Please be aware of the following problem](#16221-please-be-aware-of-the-following-problem)
  
# 1. Requirements  
# 2. Parameters  

Export-RecipientPermissions.ps1
==================================================
Finds all recipients with a primary SMTP address in an on premises Exchange environment and/or an Exchange Online/Office 365 environment and documents their
	access rights,
	"send as" permissions,
	"send on behalf" permissions,
	the "managed by" attribute of mail-enabled groups,
	and the permissions set on mailbox folders.

The report is saved to the file 'Export-RecipientPermissions_Output.csv', which consists of the following columns:
	Grantor Primary SMTP: The primary SMTP address of the object granting a permission.
	Grantor Display Name: The display name of the object granting a permission.
	Grantor Recipient Type: The recipient type of the object granting a permission.
	Grantor Environment: Shows if the object granting a permission is held on-prem or in the cloud.
	Trustee Primary SMTP: The primary SMTP address of the object receiving a permission.
	Trustee Display Name: The display name of the object receiving a permission.
	Trustee Original Identity: The original identity string of the object receiing a permission.
	Trustee Recipient Type: The recipient type of the object receiving a permission.
	Trustee Environment: Shows if the object receiviing a permission is held on-prem or in the cloud.
	Permission(s): The permission(s) granted/received (e.g., FullAccess, SendAs, SendOnBehalf etc.)
	Folder Name: The name of the mailbox folder when exporting mailbox folder permissions.

Parameters like the name of the output file, the environment (on-prem/cloud) to check and the permissions to check can be configured directly in the top section of the file.

The script can run many hours or even days, depending on the number of recipients and the speed of the environments to check.
Exporting mailbox folder permissions takes even more time because of how Exchange is designed to store these permissions.

The script does not consider group memberships as this would take too much time to evaluate: User A grants Group B a permission. This permission is documented in the output file, but not the fact that User C is member of Group B and therefore has rights on User A's mailbox, too.

The script needs to be run with an account that has read permissions to all recipients in the cloud as well as Active Directory and Exchange on premises. The script asks for cloud credentials.
As the cloud credentials are stored in the encrypted secure string file format and can be re-used, the script can be fully automated and run as a scheduled job.

Per default, the script uses 10 background threads, each one consuming one Office 365 PowerShell session, in parallel to speed up data gathering. Per default, each of these threads works on 50 recipients one after the other. The number of parallel threads and the number of recipients per thread are configurable.


Get-DependentMailboxes.ps1
==================================================
Currently only the "full access" mailbox permission works cross-premises according to Microsoft. All other permissions, including the one to manage the members of distribution lists, only work when both, the grantor and the trustee, are hosted on the same environment.
There are environments where permissions work cross-premises, but there is no offical support from Microsoft.

This script takes a list of recipients and the output of Export-RecipientPermissions.ps1 to create a list of all mailboxes and distribution groups that have a grantor-trustee dependency beyond "full access" to each other.

The script not only considers situations where recipient A grants rights to recipient B, but the whole permission chain ("X-Z-A-B-C-D" etc.).

The script does not consider group memberships as this would take too much time to evaluate: User A grants Group B a permission. This permission is documented in the output file, but not the fact that User C is member of Group B and therefore has rights on User A's mailbox, too.

The script only considers dependencies between on-prem recipients, as it is only intended to be used to accelerate the move to the cloud.

The following outputs are created:
	Export-RecipientPermissions_Output_Modified.csv
	The original permission input file, reduced to the rows that have a connection with the recipient input file.
	Enhanced with information if a grantor or trustee are part of the initial recipient file or have to be migrated additionally.
	Enhanced with information which single permissions start permissions chains outside the initial recipients.

	Get-DependentRecipients_OriginalInput.csv
	The original recipient input file for documentation purposes.

	Get-DependentRecipients_Output_AdditionalRecipients.csv
	List of additional recipients. Format: "Primary SMTP address;Recipient type;Environment".

	Get-DependentRecipients_Output_AllRecipients.csv
	Lists of all initial and additional recipients, including their recipient type and environment. Format: "Primary SMTP address;Recipient type;Environment".

	Get-DependentRecipients_Output_AllRecipients.gml
	All recipients and their permissions in a graphical representation. The gml (Graph Modeling Language) file format used is human readable. Free tools like yEd Graph Editor, Gephi and others can be used to easily create visual representations from this file.
	You can use the file "OUs.csv" to have mailboxes grouped by OUs and their friendly names.

	Get-DependentRecipients_Output_Summary.txt
	Number of initial recipients, number of additional recipients, number of total recipients, number of root cause mailbox permissions.


Recommendations
==================================================
Make sure you have the latest updates installed to avoid memory leaks and CPU spikes (PowerShell, Exchange Management tools, .Net framework).

If possible, allow Export-RecipientPermissions.ps1 to use your on premises infrastructure. This will dramatically increase the initial enumeration of recipients.

Start the script from powershell.exe, not from within the PowerShell ISE. This makes especially Get-DependentMailboxes.ps1 run faster due to a different default thread apartment mode.

When running the scripts as scheduled job, make sure to include the "-ExecutionPolicy Bypass" parameter.
Example: powershell.exe -ExecutionPolicy Bypass -file "c:\path\Export-RecipientPermissions.ps1"
