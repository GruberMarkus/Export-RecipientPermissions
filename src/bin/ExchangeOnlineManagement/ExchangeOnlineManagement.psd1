@{
RootModule = if($PSEdition -eq 'Core')
{
    '.\netCore\ExchangeOnlineManagement.psm1'
}
else # Desktop
{
    '.\netFramework\ExchangeOnlineManagement.psm1'
}
FunctionsToExport = @('Connect-ExchangeOnline', 'Connect-IPPSSession', 'Disconnect-ExchangeOnline')
ModuleVersion = '3.4.0'
GUID = 'B5ECED50-AFA4-455B-847A-D8FB64140A22'
Author = 'Microsoft Corporation'
CompanyName = 'Microsoft Corporation'
Copyright = '(c) 2021 Microsoft. All rights reserved.'
Description = 'This is a General Availability (GA) release of the Exchange Online Powershell V3 module. Exchange Online cmdlets in this module are REST-backed and do not require Basic Authentication to be enabled in WinRM. REST-based connections in Windows require the PowerShellGet module, and by dependency, the PackageManagement module.
Please check the documentation here - https://aka.ms/exov3-module.
For issues related to the module, contact Microsoft support.'
PowerShellVersion = '3.0'
CmdletsToExport = @('Add-VivaModuleFeaturePolicy','Get-ConnectionInformation','Get-DefaultTenantBriefingConfig','Get-DefaultTenantMyAnalyticsFeatureConfig','Get-EXOCasMailbox','Get-EXOMailbox','Get-EXOMailboxFolderPermission','Get-EXOMailboxFolderStatistics','Get-EXOMailboxPermission','Get-EXOMailboxStatistics','Get-EXOMobileDeviceStatistics','Get-EXORecipient','Get-EXORecipientPermission','Get-MyAnalyticsFeatureConfig','Get-UserBriefingConfig','Get-VivaInsightsSettings','Get-VivaModuleFeature','Get-VivaModuleFeatureEnablement','Get-VivaModuleFeaturePolicy','Remove-VivaModuleFeaturePolicy','Set-DefaultTenantBriefingConfig','Set-DefaultTenantMyAnalyticsFeatureConfig','Set-MyAnalyticsFeatureConfig','Set-UserBriefingConfig','Set-VivaInsightsSettings','Update-VivaModuleFeaturePolicy')
FileList = if($PSEdition -eq 'Core')
{
    @('.\netCore\Azure.Core.dll',
        '.\netCore\Microsoft.Bcl.AsyncInterfaces.dll',
        '.\netCore\Microsoft.Bcl.HashCode.dll',
        '.\netCore\Microsoft.Exchange.Management.AdminApiProvider.dll',
        '.\netCore\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll',
        '.\netCore\Microsoft.Exchange.Management.RestApiClient.dll',
        '.\netCore\Microsoft.Identity.Client.dll',
        '.\netCore\Microsoft.IdentityModel.Abstractions.dll',
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
        '.\netCore\msvcp140.dll',
        '.\netCore\Newtonsoft.Json.dll',
        '.\netCore\System.CodeDom.dll',
        '.\netCore\System.Configuration.ConfigurationManager.dll',
        '.\netCore\System.Diagnostics.PerformanceCounter.dll',
        '.\netCore\System.DirectoryServices.dll',
        '.\netCore\System.Drawing.Common.dll',
        '.\netCore\System.IdentityModel.Tokens.Jwt.dll',
        '.\netCore\System.Management.dll',
        '.\netCore\System.Memory.Data.dll',
        '.\netCore\System.Security.Cryptography.Pkcs.dll',
        '.\netCore\System.Security.Cryptography.ProtectedData.dll',
        '.\netCore\System.Security.Permissions.dll',
        '.\netCore\System.Windows.Extensions.dll',
        '.\netCore\vcruntime140_1.dll',
        '.\netCore\vcruntime140.dll',
        '.\license.txt')
}
else # Desktop
{
    @('.\netFramework\Microsoft.Exchange.Management.AdminApiProvider.dll',
        '.\netFramework\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll',
        '.\netFramework\Microsoft.Exchange.Management.RestApiClient.dll',
        '.\netFramework\Microsoft.Identity.Client.dll',
        '.\netFramework\Microsoft.IdentityModel.Abstractions.dll',
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
        '.\license.txt')
}

PrivateData = @{
    PSData = @{
    # Tags applied to this module. These help with module discovery in online galleries.
    Tags = 'Exchange', 'ExchangeOnline', 'EXO', 'EXOV2', 'EXOV3', 'Mailbox', 'Management'
    ReleaseNotes = '
---------------------------------------------------------------------------------------------
What is new in this release:

v3.4.0 :
    1.  Bug fixes in Connect-ExchangeOnline, Get-EXORecipientPermission and Get-EXOMailboxFolderPermission.
    2.  Support to use Constrained Language Mode(CLM) using SigningCertificate parameter.

---------------------------------------------------------------------------------------------
Previous Releases:

v3.3.0 :
    1.  Support to skip loading cmdlet help files with Connect-ExchangeOnline.
    2.  Global variable EXO_LastExecutionStatus can now be used to check the status of the last cmdlet that was executed.
    3.  Bug fixes in Connect-ExchangeOnline and Connect-IPPSSession.
    4.  Support of user controls enablement by policy for features that are onboarded to Viva feature access management.

v3.2.0 :
    1.  General Availability of new cmdlets:
        -  Updating Briefing Email Settings of a tenant (Get-DefaultTenantBriefingConfig and Set-DefaultTenantBriefingConfig)
        -  Updating Viva Insights Feature Settings of a tenant (Get-DefaultTenantMyAnalyticsFeatureConfig and Set-DefaultTenantMyAnalyticsFeatureConfig)
        -  View the features in Viva that support setting access management policies (Get-VivaModuleFeature)
        -  Create and manage Viva app feature policies
           -  Get-VivaModuleFeaturePolicy
           -  Add-VivaModuleFeaturePolicy
           -  Remove-VivaModuleFeaturePolicy
           -  Update-VivaModuleFeaturePolicy
        -  View whether or not a Viva feature is enabled for a specific user/group (Get-VivaModuleFeatureEnablement)

    2.  General Availability of REST based cmdlets for Security and Compliance PowerShell.
    3.  Support to get REST connection informations from Get-ConnectionInformation cmdlet and disconnect REST connections using Disconnect-ExchangeOnline cmdlet for specific connection(s).
    4.  Support to sign the temporary generated module with a client certificate to use the module in all PowerShell execution policies.
    5.  Bug fixes in Connect-ExchangeOnline.

v3.1.0 :
    1.  Support for providing an Access Token with Connect-ExchangeOnline.
    2.  Bug fixes in Connect-ExchangeOnline and Get-ConnectionInformation.
    3.  Bug fix in Connect-IPPSSession for connecting to Security and Compliance PowerShell using Certificate Thumbprint.

v3.0.0 :
    1.  General Availability of REST-backed cmdlets for Exchange Online which do not require WinRM Basic Authentication to be enabled.
    2.  General Availability of Certificate Based Authentication for Security and Compliance PowerShell cmdlets.
    3.  Support for System-Assigned and User-Assigned ManagedIdentities to connect to ExchangeOnline from Azure VMs, Azure Virtual Machine Scale Sets and Azure Functions.
    4.  Breaking changes
        -   Get-PSSession cannot be used to get information about the sessions created as PowerShell Remoting is no longer being used. The Get-ConnectionInformation cmdlet has been introduced instead, to get information about the existing connections to ExchangeOnline. Refer https://docs.microsoft.com/en-us/powershell/module/exchange/get-connectioninformation?view=exchange-ps for more information.
        -   Certain cmdlets that used to prompt for confirmation in specific scenarios will no longer have this prompt and the cmdlet will run to completion by default.
        -   The format of the error returned from a failed cmdlet execution has been slightly modified. The Exception contains some additional data such as the exception type, and the FullyQualifiedErrorId does not contain the FailureCategory. The format of the error is subject to further modifications.
        -   Deprecation of the Get-OwnerlessGroupPolicy and Set-OwnerlessGroupPolicy cmdlets.

v2.0.5 :
    1. Manage ownerless Microsoft 365 groups through newly added cmdlets Get-OwnerlessGroupPolicy and Set-OwnerlessGroupPolicy.
    2. Add new cmdlets Get-VivaInsightsSettings and Set-VivaInsightsSettings for Global/ExchangeOnline/Teams administrators to control user access of Headspace features in Viva Insights.

v2.0.4 :
    1. Manage EXO using Linux devices along with Browser based SSO Authentication for enhanced interactive management experience. No need to enter UserName and password everytime you run the PowerShell script.
    2. Manage EXO using Apple Macintosh devices. Supported versions of Apple MAC OS are Mojave, Catalina & Big Sur. Steps for installing PowerShell on MAC OS is documented here - https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-core-on-macos?view=powershell-7.1
    3. Real time policy & security enforcement in all user based authentication. Continuous Access Evaluation (CAE) has been enabled in EXO V2 Module. Read more about CAE here - https://techcommunity.microsoft.com/t5/azure-active-directory-identity/moving-towards-real-time-policy-and-security-enforcement/ba-p/1276933
    4. Use parameter InlineCredential to pass credentials of Non-MFA accounts on the go without the need of storing credentials in a variable
    5. More secure method to fetch access token using safe Reply URLs.
    6. Breaking change :- Change in cmdlet signature to configure MyAnalytics access for users in your tenant. Get/Set-UserAnalyticsConfig has been replaced by Get/Set-MyAnalyticsFeatureConfig Additionally, you can have more granular controls and configure access at feature level. For more steps read here - https://docs.microsoft.com/en-us/workplace-analytics/myanalytics/setup/configure-myanalytics

v2.0.3 :
    1. General availability of Certificate Based Authentication feature which enables using Modern Authentication in Unattended Scripting or background automation scenarios.
    2. Certificate Based Authentication accepts Certificate File directly from terminal thus enabling certificate files to be stored in Azure Key Vault and being fetched Just-In-Time for enhanced security. See parameter Certificate in Connect-ExchangeOnline.
    3. Connect with Exchange Online and Security Compliance Center simultaneously in a single PowerShell window.
    4. Ability to restrict the PowerShell cmdlets imported in a session using CommandName parameter, thus reducing memory footprint in case of high usage PowerShell applications.
    5. Get-ExoMailboxFolderPermission now supports ExternalDirectoryObjectID in the Identity parameter.
    6. Optimized latency of first V2 Cmdlet call. (Lab results show first call latency has been reduced from 8 seconds to ~1 seconds. Actual results will depend on result size and Tenant environment.)
 
v1.0.1 :
    1. This is the General Availability (GA) version of EXO PowerShell V2 Module. It is stable and ready for being used in production environments.
    2. Get-ExoMobileDeviceStatistics cmdlet now supports Identity parameter.
    3. Improved reliability of session auto-connect in certain cases where script was executing for ~50minutes and threw "Cmdlet not found" error due to a bug in auto-reconnect logic.
    4. Fixed data-type issues of two commonly used attributed "User" and "MailboxFolderUser" for easy migration of scripts.
    5. Enhanced support for filters as it now supports 4 more operators - endswith, contains, not and notlike support. Please check online documentation for attributes which are not supported in filter string.
 
---------------------------------------------------------------------------------------------
'
    LicenseUri='http://aka.ms/azps-license'
    }
}
}

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC5k4NLZSnwjwQa
# TWYsjsTna789xeUB/GReqnEPO4aw96CCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGeBMKdFbObl1GX6/sQ6++RV
# hLK25/NguTShkHhIDjRZMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAiPT46hqldGPEqOMVDTDDJT7T4edIvSBD+DOftudRcHmQLuJKWlYhu0tL
# GQxwrVU7Yqa+ZAUFdORKCJqav/Plz8AON0vEuX94zg5QIbovMCk5gq9JYxNic3VA
# 0+7w9GSG5BAQpEBDkoh/xC2d6dJTUn5eKGy8eAQtOs2sPIbEtOcZInCPC/C7VEJA
# YRIDys5N94pXzM0S0LszVGhzJQjVfFOSRUdpU8N09yY/zB8QajmoNipMjaiZ9WEI
# X+7wuqTaKcyFWIjEE56whPsiE+BTaoUGUxoPHrM1CAOyOpuyZhDPIBOcSdfSGkRy
# 8nAoOMoOCNTlWIGEYfNvBuGSYvU5MqGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCDRW9jKd7Ay2qFjclMqZxf+O/hTAZA+jEQnH8jseMOqogIGZQQZQYu2
# GBMyMDIzMDkyODA2MTcwMy40NDZaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODkwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAdMdMpoXO0AwcwABAAAB0zANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# MjRaFw0yNDAyMDExOTEyMjRaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODkwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQC0jquTN4g1xbhXCc8MV+dOu8Uqc3KbbaWti5vdsAWM
# 1D4fVSi+4NWgGtP/BVRYrVj2oVnnMy0eazidQOJ4uUscBMbPHaMxaNpgbRG9FEQR
# FncAUptWnI+VPl53PD6MPL0yz8cHC2ZD3weF4w+uMDAGnL36Bkm0srONXvnM9eNv
# nG5djopEqiHodWSauRye4uftBR2sTwGHVmxKu0GS4fO87NgbJ4VGzICRyZXw9+Rv
# vXMG/jhM11H8AWKzKpn0oMGm1MSMeNvLUWb31HSZekx/NBEtXvmdo75OV030NHgI
# XihxYEeSgUIxfbI5OmgMq/VDCQp2r/fy/5NVa3KjCQoNqmmEM6orAJ2XKjYhEJzo
# p4nWCcJ970U6rXpBPK4XGNKBFhhLa74TM/ysTFIrEXOJG1fUuXfcdWb0Ex0FAeTT
# r6gmmCqreJNejNHffG/VEeF7LNvUquYFRndiCUhgy624rW6ptcnQTiRfE0QL/gLF
# 41kA2vZMYzcc16EiYXQQBaF3XAtMduh1dpXqTPPQEO3Ms5/5B/KtjhSspMcPUvRv
# b35IWN+q+L+zEwiphmnCGFTuyOMqc5QE0ruGN3Mx0Vv6x/hcOmaXxrHQGpNKI5Pn
# 79Yk89AclqU2mXHz1ZHWp+KBc3D6VP7L32JlwxhJx3asa085xv0XPD58MRW1WaGv
# aQIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFNLHIIa4FAD494z35hvzCmm0415iMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBAYlhYoUQ+4aaQ54MFNfE6Ey8v4rWv+LtD
# RSjMM2X9g4uanA9cU7VitdpIPV/zE6v4AEhe/Vng2UAR5qj2SV3sz+fDqN6VLWUZ
# sKR0QR2JYXKnFPRVj16ezZyP7zd5H8IsvscEconeX+aRHF0xGGM4tDLrS84vj6Rm
# 0bgoWLXWnMTZ5kP4ownGmm0LsmInuu0GKrDZnkeTVmfk8gTTy8d1y3P2IYc2UI4i
# JYXCuSaKCuFeO0wqyscpvhGQSno1XAFK3oaybuD1mSoQxT9q77+LAGGQbiSoGlgT
# jQQayYsQaPcG1Q4QNwONGqkASCZTbzJlnmkHgkWlKSLTulOailWIY4hS1EZ+w+sX
# 0BJ9LcM142h51OlXLMoPLpzHAb6x22ipaAJ5Kf3uyFaOKWw4hnu0zWs+PKPd192n
# deK2ogWfaFdfnEvkWDDH2doL+ZA5QBd8Xngs/md3Brnll2BkZ/giZE/fKyolriR3
# aTAWCxFCXKIl/Clu2bbnj9qfVYLpAVQEcPaCfTAf7OZBlXmluETvq1Y/SNhxC6MJ
# 1QLCnkXSI//iXYpmRKT783QKRgmo/4ztj3uL9Z7xbbGxISg+P0HTRX15y4TReBbO
# 2RFNyCj88gOORk+swT1kaKXUfGB4zjg5XulxSby3uLNxQebE6TE3cAK0+fnY5UpH
# aEdlw4e7ijCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNN
# MIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjg5MDAtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBS
# x23cMcNB1IQws/LYkRXa7I5JsKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6L8RsjAiGA8yMDIzMDkyNzIwNDAx
# OFoYDzIwMjMwOTI4MjA0MDE4WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDovxGy
# AgEAMAcCAQACAkHLMAcCAQACAhdZMAoCBQDowGMyAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAHQplEc2YQZZ49ae4L/jJF4fAasQAmMCv732DuvwoHnM81or
# 1MGJ8/bpwJA17L5zYPFiOKxWokumpTyQwVyX1E44+vEXKfMQo/0SGDfiUVe8Vsi0
# 5M+YVV1EnowyQE4hOtf1R65FbaUwHVx7x29DWszTRNVvjXupj6T50ywbP44wP6nP
# 9ifboyXrqAjY6dI6GxnH6jWCH8B73ZFjh5GmIR3DL7TnQl3KCyP9Aoq0pFXGgd9i
# bG8OE+K16qzreQQH9C9pIguJxtRYvy7KDlKKg39be80IZb6MhGY/FbvLcWn3Fc/m
# zFgJ7ToJoeYOj8ZvXY1dYalXCl+syIONGw3f/u8xggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdMdMpoXO0AwcwABAAAB0zAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCBC0d02oQ0TQ9Wh0eN3+EuqB+kbUQ/RKjJrRMciT+Sh1jCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIJJm9OrE4O5PWA1KaFaztr9uP96r
# QgEn+tgGtY3xOqr1MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHTHTKaFztAMHMAAQAAAdMwIgQgGu76YSKOQH3GEzEv3/3ApyYW/m+h
# 04IvMI3otmZZwdEwDQYJKoZIhvcNAQELBQAEggIAgQoq/+m+03fJc8KyG4TuK77S
# 8biW1KYg4BEU0LfObBdn5m7cwQhz+2ziQHYUMacHxFJj/ek+soucXXcamh8eckb7
# SWxj2eSXWBgWREoapbF43voUFU1yMwQ3LaMpcGKC4mlWZIqyMRst0W2GvA+JTCKb
# E0TAhYRYvKXY82Z2HLlywooCjviOv693tPldy43h3g+mLooPoP5WUqy81v+6I11t
# +rx7H+aFeAFUE7FzhpKSwHbMT2+qhooG5kgpvdrRNh3/jSVkaj535y1iZjU4pvQF
# F8FpPPKU4aVSh1o9lK55cDCnm62ULRggklru3LKCGKG4X+YN+4QIsPo4P1tkqT0o
# Qxnx5fD+qilFqlAO5EfWQ22A+j/kaatVGNvXY/VQYQTAR85W9i358fApILm6trkd
# 4nwehcY2pRj1ICqUSqDnHGKUbvtUUWQHvez8ZRpufH6pQcwJzRuV/2XxRQJtnI9E
# geee+4QsZXgjjfbJ/Cy0cTRKHmmv+PE+h+AhXB3MawT+dEJID0hskMI25y/l24F4
# 3KsfhFOoq86iesoM+vWSwbWiE7PTR+1beSbENh0pL3SuGcpaatR3J3N0cmJohfHT
# EsQCF4+3S4G8sQu8Ok7oz3NWluCanUZyVJ8CKPRQynghY1zxHLE/zaT+0UpKkKGt
# kA5D3+WuC0Li8VSuH9M=
# SIG # End signature block
