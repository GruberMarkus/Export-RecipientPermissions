@{
RootModule = if($PSEdition -eq 'Core')
{
    '.\netCore\ExchangeOnlineManagementBeta.psm1'
}
else # Desktop
{
    '.\netFramework\ExchangeOnlineManagementBeta.psm1'
}
FunctionsToExport = @('Connect-ExchangeOnline', 'Connect-IPPSSession', 'Disconnect-ExchangeOnline')
ModuleVersion = '2.0.6'
GUID = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'
Author = 'Microsoft Corporation'
CompanyName = 'Microsoft Corporation'
Copyright = '(c) 2021 Microsoft. All rights reserved.'
Description = 'This is a Public Preview release of Exchange Online PowerShell V2 module.
Please check the documentation here - https://aka.ms/exops-docs.
For issues related to the module, contact Microsoft support.'
PowerShellVersion = '3.0'
CmdletsToExport = @('Get-ConnectionInformation','Get-EXOCasMailbox','Get-EXOMailbox','Get-EXOMailboxFolderPermission','Get-EXOMailboxFolderStatistics','Get-EXOMailboxPermission','Get-EXOMailboxStatistics','Get-EXOMobileDeviceStatistics','Get-EXORecipient','Get-EXORecipientPermission','Get-MyAnalyticsFeatureConfig','Get-UserBriefingConfig','Get-VivaInsightsSettings','Set-MyAnalyticsFeatureConfig','Set-UserBriefingConfig','Set-VivaInsightsSettings')
FileList = if($PSEdition -eq 'Core')
{
    @('.\netCore\Microsoft.Bcl.AsyncInterfaces.dll',
        '.\netCore\Microsoft.Exchange.Management.AdminApiProvider.dll',
        '.\netCore\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll',
        '.\netCore\Microsoft.Exchange.Management.RestApiClient.dll',
        '.\netCore\Microsoft.Identity.Client.dll',
        '.\netCore\Microsoft.IdentityModel.JsonWebTokens.dll',
        '.\netCore\Microsoft.IdentityModel.Logging.dll',
        '.\netCore\Microsoft.IdentityModel.Tokens.dll',
        '.\netCore\Microsoft.OData.Client.dll',
        '.\netCore\Microsoft.OData.Core.dll',
        '.\netCore\Microsoft.OData.Edm.dll',
        '.\netCore\Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.dll',
        '.\netCore\Microsoft.Spatial.dll',
        '.\netCore\Microsoft.Win32.Registry.AccessControl.dll',
        '.\netCore\Microsoft.Win32.SystemEvents.dll',
        '.\netCore\Newtonsoft.Json.dll',
        '.\netCore\System.CodeDom.dll',
        '.\netCore\System.Configuration.ConfigurationManager.dll',
        '.\netCore\System.Diagnostics.PerformanceCounter.dll',
        '.\netCore\System.DirectoryServices.dll',
        '.\netCore\System.Drawing.Common.dll',
        '.\netCore\System.IdentityModel.Tokens.Jwt.dll',
        '.\netCore\System.IO.Abstractions.dll',
        '.\netCore\System.Management.Automation.dll',
        '.\netCore\System.Management.dll',
        '.\netCore\System.Security.Cryptography.Pkcs.dll',
        '.\netCore\System.Security.Cryptography.ProtectedData.dll',
        '.\netCore\System.Security.Permissions.dll',
        '.\netCore\System.Text.Encodings.Web.dll',
        '.\netCore\System.Windows.Extensions.dll',
        '.\license.txt')
}
else # Desktop
{
    @('.\netFramework\Microsoft.Exchange.Management.AdminApiProvider.dll',
        '.\netFramework\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll',
        '.\netFramework\Microsoft.Exchange.Management.RestApiClient.dll',
        '.\netFramework\Microsoft.Identity.Client.dll',
        '.\netFramework\Microsoft.IdentityModel.Clients.ActiveDirectory.dll',
        '.\netFramework\Microsoft.IdentityModel.JsonWebTokens.dll',
        '.\netFramework\Microsoft.IdentityModel.Logging.dll',
        '.\netFramework\Microsoft.IdentityModel.Tokens.dll',
        '.\netFramework\Microsoft.OData.Client.dll',
        '.\netFramework\Microsoft.OData.Core.dll',
        '.\netFramework\Microsoft.OData.Edm.dll',
        '.\netFramework\Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.dll',
        '.\netFramework\Microsoft.Spatial.dll',
        '.\netFramework\Newtonsoft.Json.dll',
        '.\netFramework\System.IdentityModel.Tokens.Jwt.dll',
        '.\netFramework\System.IO.Abstractions.dll',
        '.\netFramework\System.Management.Automation.dll',
        '.\license.txt')
}

PrivateData = @{
    PSData = @{
    # Tags applied to this module. These help with module discovery in online galleries.
    Tags = 'Exchange', 'ExchangeOnline', 'EXO', 'EXOV2', 'Mailbox', 'Management'

    # Set to a prerelease string value if the release should be a prerelease.
    Prerelease = 'Preview8'

    ReleaseNotes = '
---------------------------------------------------------------------------------------------
Whats new in this release:

v2.0.6-Preview8 :
    1. Support for system-assigned and user-assigned Managed Identity from Azure Functions.
        - The -ManagedIdentity switch parameter, and the -Organization parameters need to be provided to indicate that a managed identity should be used. This will by default attempt to use a system-assigned managed identity.
        - For specifying a user-assigned managed identity, in addition to the parameters specified above, the AppID of the service principal corresponding to the user-assigned identity needs to be passed to the -ManagedIdentityAccountId.
    2. Support for formatted output data added.
	    - By default, the output now will be formatted similar to version 2.0.5. The -SkipLoadingFormatData switch parameter can be specified with Connect-ExchangeOnline to avoid loading the format data and execute Connect-ExchangeOnline faster. 
    3. Bug fixes in Connect-ExchangeOnline and Get-ConnectionInformation.

---------------------------------------------------------------------------------------------
Previous Releases:

v2.0.6-Preview7 :
    1. Support for system-assigned and user-assigned Managed Identity from Azure VMs and Virtual Machine Scale Sets.
        - The -ManagedIdentity switch parameter, and the -Organization parameters need to be provided to indicate that a managed identity should be used. This will by default attempt to use a system-assigned managed identity.
        - For specifying a user-assigned managed identity, in addition to the parameters specified above, the AppID of the service principal corresponding to the user-assigned identity needs to be passed to the -ManagedIdentityAccountId.
    2. Get-ConnectionInformation cmdlet introduced to get the information regarding all REST-based active connections to ExchangeOnline. This is similar to the Get-PSSession cmdlet which returns information on all the remote powershell sessions in the runspace.
    3. Bug fixes in Connect-ExchangeOnline.

v2.0.6-Preview6 :
    1. Compliance with Continuous Access Evaluation(CAE). If you have any Conditional Access policy enabled for your organization, you can now use RPS and the REST-based cmdlets in addition to the EXO cmdlets.
    2. Revamped error reporting framework for the REST-based cmdlets that maintains an individual log file per connection to EXO from the same powershell instance.
    3. Bug fixes in Connect-ExchangeOnline.

v2.0.6-Preview5 :
    1. General availability of Certificate Based Authentication feature which enables using Modern Authentication in Unattended Scripting or background automation scenarios for Connect-IPPSSession.
    2. Cmdlets for creating and assigning custom nudges to users that will show in their briefing emails have been introduced. These include the cmdlets ending in *-CustomNudge, *-CustomNudgeSettings, and *-CustomNudgeAssignment. As these cmdlets are still in development, they may not yet be enabled for your tenant.
    3. The Get-OwnerlessGroupPolicy and Set-OwnerlessGroupPolicy cmdlets that are available in version 2.0.5 have been deprecated and are no longer available. You can manage ownerless Microsoft 365 Groups in the Microsoft 365 admin center.
    4. Bug fixes in Connect-ExchangeOnline and Disconnect-ExchangeOnline.

v2.0.6-Preview4:
    1. Add new Feature DiscoverTryBuy in cmdlets Get-VivaInsightsSettings and Set-VivaInsightsSettings for Global/ExchangeOnline/Teams administrators to control user access of Discover Try Buy features in Viva Insights.

v2.0.6-Preview3 :
    1. This version contains all the * EXO * cmdlets along with 250 Remote PowerShell cmdlets which are REST API backed. These REST API backed cmdlets do not need PowerShell session and hence they do not need WinRM Basic Auth to be enabled and they work as-is without requiring any change in script.
    2. Use switch -RPSSession to use the default set of all ~900 RPS Cmdlets. Using RPSSession switch needs WinRM Basic Auth to be enabled on your client machine.
    3. For certain RPS Cmdlets, use switch -UseCustomRouting to route your requests directly to the required mailbox server and it may improve the performance of overall script execution. In case UseCustomRouting flag is passed, please pass the value of either of UserPrincipalName, SmtpAddress, MailboxGuid. Use this parameter as an experiment and share feedback to exocmdletpreview@service.microsoft.com.
        - This parameter is initially available for these 20 cmdlets and can be updated - Get-MailboxStatistics,Get-MailboxAutoReplyConfiguration,Get-MailboxMessageConfiguration,Get-MailboxPermission,Get-MailboxFolderStatistics,Get-MobileDeviceStatistics,Get-InboxRule,Get-MailboxRegionalConfiguration,Set-MailboxRegionalConfiguration,Get-UserPhoto,Set-UserPhoto,Remove-CalendarEvents,Set-Clutter,Get-MailboxCalendarFolder,Get-Clutter,Get-MailboxFolderPermission,Get-FocusedInbox,Set-FocusedInbox. For more accurate information, please check online documentation of EXO PowerShell V2 Module.

v2.0.5 :
    1. Manage ownerless Microsoft 365 groups through newly added cmdlets Get-OwnerlessGroupPolicy and Set-OwnerlessGroupPolicy.
    2. Add new cmdlets Get-VivaInsightsSettings and Set-VivaInsightsSettings for Global/ExchangeOnline/Teams administrators to control user access of Headspace features in Viva Insights.
 
---------------------------------------------------------------------------------------------
'
    LicenseUri='http://aka.ms/azps-license'
    }
}
}
