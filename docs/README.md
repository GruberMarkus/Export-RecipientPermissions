<!-- omit in toc -->
# **<a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="" width="400" title="Export-RecipientPermissions" alt="Export-RecipientPermissions"></a>**<br>Document Exchange mailbox access rights, folder permissions, "send as", "send on behalf" and "managed by"<br><!--XXXRemoveWhenBuildingXXX<a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/badge/this%20release-XXXVersionStringXXX-informational" alt=""></a> XXXRemoveWhenBuildingXXX--><a href="https://github.com/GruberMarkus/Export-RecipientPermissions" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Export-RecipientPermissions" alt=""></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/v/release/GruberMarkus/Export-RecipientPermissions?display_name=tag&include_prereleases&sort=semver&label=latest%20release&color=informational" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a><br><a href="https://github.com/sponsors/GruberMarkus" target="_blank"><img src="https://img.shields.io/badge/sponsor-white?logo=githubsponsors" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Export-RecipientPermissions/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/network/members" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Export-RecipientPermissions/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Export-RecipientPermissions/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Export-RecipientPermissions" alt="" data-external="1"></a>  

# Features <!-- omit in toc -->
Finds all recipients with a primary SMTP address in an on premises Exchange environment and/or an Exchange Online/Office 365 environment and documents their
- access rights,
- folder permissions,
- "send as" permissions,
- "send on behalf" permissions,
- the "managed by" attribute of mail-enabled groups,

Easens the move to the cloud as permission dependencies beyond the supported cross-premises "full control" can easily be found and even be represented graphically (sample code included).

Compare exports from different times to detect permission changes (sample code included). 

# Table of Contents <!-- omit in toc -->
- [1. Export-RecipientPermissions.ps1](#1-export-recipientpermissionsps1)
	- [1.1. Output](#11-output)
	- [1.2. Parameters](#12-parameters)
	- [1.3. Runtime](#13-runtime)
	- [1.4. Group membership](#14-group-membership)
	- [1.5. Requirements](#15-requirements)
- [2. Get-DependentMailboxes.ps1](#2-get-dependentmailboxesps1)
- [3. Recommendations](#3-recommendations)

# 1. Export-RecipientPermissions.ps1
Finds all recipients with a primary SMTP address in an on premises Exchange environment and/or an Exchange Online/Office 365 environment and documents their
- access rights,
- folder permissions,
- "send as" permissions,
- "send on behalf" permissions,
- the "managed by" attribute of mail-enabled groups,
## 1.1. Output
The report is saved to the file 'Export-RecipientPermissions_Output.csv', which consists of the following columns:
- Grantor Primary SMTP: The primary SMTP address of the object granting a permission.
- Grantor Display Name: The display name of the object granting a permission.
- Grantor Recipient Type: The recipient type of the object granting a permission.
- Grantor Environment: Shows if the object granting a permission is held on-prem or in the cloud.
- Trustee Primary SMTP: The primary SMTP address of the object receiving a permission.
- Trustee Display Name: The display name of the object receiving a permission.
- Trustee Original Identity: The original identity string of the object receiing a permission.
- Trustee Recipient Type: The recipient type of the object receiving a permission.
- Trustee Environment: Shows if the object receiviing a permission is held on-prem or in the cloud.
- Permission(s): The permission(s) granted/received (e.g., FullAccess, SendAs, SendOnBehalf etc.)
- Folder Name: The name of the mailbox folder when exporting mailbox folder permissions.
## 1.2. Parameters
Parameters like the name of the output file, the environment (on-prem/cloud) to check and the permissions to check can be configured directly in the top section of the file.
## 1.3. Runtime
The script can run many hours or even days, depending on the number of recipients and the speed of the environments to check.
Exporting mailbox folder permissions takes even more time because of how Exchange is designed to store these permissions.
## 1.4. Group membership
The script optionally considers group memberships. This can take much time to evaluate.
## 1.5. Requirements
The script needs to be run with an account that has read permissions to all recipients in the cloud as well as Active Directory and Exchange on premises. The script asks for cloud credentials.
As the cloud credentials are stored in the encrypted secure string file format and can be re-used, the script can be fully automated and run as a scheduled job.

Per default, the script uses 10 background threads, each one consuming one Office 365 PowerShell session, in parallel to speed up data gathering. Per default, each of these threads works on 50 recipients one after the other. The number of parallel threads and the number of recipients per thread are configurable.

# 2. Get-DependentMailboxes.ps1
Currently only the "full access" mailbox permission works cross-premises according to Microsoft. All other permissions, including the one to manage the members of distribution lists, only work when both, the grantor and the trustee, are hosted on the same environment.
There are environments where permissions work cross-premises, but there is no offical support from Microsoft.

This script takes a list of recipients and the output of Export-RecipientPermissions.ps1 to create a list of all mailboxes and distribution groups that have a grantor-trustee dependency beyond "full access" to each other.

The script not only considers situations where recipient A grants rights to recipient B, but the whole permission chain ("X-Z-A-B-C-D" etc.).

The script optionally considers group membership. This can take too much time to evaluate.

The script only considers dependencies between on-prem recipients, as it is only intended to be used to accelerate the move to the cloud.

The following outputs are created:
- Export-RecipientPermissions_Output_Modified.csv  
	The original permission input file, reduced to the rows that have a connection with the recipient input file.  
	Enhanced with information if a grantor or trustee are part of the initial recipient file or have to be migrated additionally.
	Enhanced with information which single permissions start permissions chains outside the initial recipients.
-	Get-DependentRecipients_OriginalInput.csv  
	The original recipient input file for documentation purposes.
-	Get-DependentRecipients_Output_AdditionalRecipients.csv  
	List of additional recipients. Format: "Primary SMTP address;Recipient type;Environment".
-	Get-DependentRecipients_Output_AllRecipients.csv  
	Lists of all initial and additional recipients, including their recipient type and environment. Format: "Primary SMTP address;Recipient type;Environment".
-	Get-DependentRecipients_Output_AllRecipients.gml  
	All recipients and their permissions in a graphical representation. The gml (Graph Modeling Language) file format used is human readable. Free tools like yEd Graph Editor, Gephi and others can be used to easily create visual representations from this file.  
	You can use the file "OUs.csv" to have mailboxes grouped by OUs and their friendly names.
-	Get-DependentRecipients_Output_Summary.txt  
	Number of initial recipients, number of additional recipients, number of total recipients, number of root cause mailbox permissions.

# 3. Recommendations
Make sure you have the latest updates installed to avoid memory leaks and CPU spikes (PowerShell, Exchange Management tools, .Net framework).

If possible, allow Export-RecipientPermissions.ps1 to use your on premises infrastructure. This will dramatically increase the initial enumeration of recipients.

Start the script from powershell.exe, not from within the PowerShell ISE. This makes especially Get-DependentMailboxes.ps1 run faster due to a different default thread apartment mode.

When running the scripts as scheduled job, make sure to include the "-ExecutionPolicy Bypass" parameter.
Example: `powershell.exe -ExecutionPolicy Bypass -file "c:\path\Export-RecipientPermissions.ps1"`