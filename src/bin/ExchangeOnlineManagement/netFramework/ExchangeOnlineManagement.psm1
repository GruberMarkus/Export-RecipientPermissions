# Import the REST module so that the EXO* cmdlets are present before Connect-ExchangeOnline in the powershell instance.
$RestModule = "Microsoft.Exchange.Management.RestApiClient.dll"
$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestModule)
Import-Module $RestModulePath 

$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"
$ExoPowershellModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule)
Import-Module $ExoPowershellModulePath

############# Helper Functions Begin #############

    <#
    Details to be printed on the console when the Connect-ExchangeOnline function is run
    #>
    function Print-Details
    {
	    if ($UseRPSSession)
		{
			Write-Host -ForegroundColor Yellow ""
			Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------"
			Write-Host -ForegroundColor Yellow "The module allows access to all existing remote PowerShell (V1) cmdlets in addition to the 9 new, faster, and more reliable cmdlets."
			Write-Host -ForegroundColor Yellow ""
			Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
			Write-Host -ForegroundColor Yellow "|    Old Cmdlets                    |    New/Reliable/Faster Cmdlets       |"
			Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
			Write-Host -ForegroundColor Yellow "|    Get-CASMailbox                 |    Get-EXOCASMailbox                 |"
			Write-Host -ForegroundColor Yellow "|    Get-Mailbox                    |    Get-EXOMailbox                    |"
			Write-Host -ForegroundColor Yellow "|    Get-MailboxFolderPermission    |    Get-EXOMailboxFolderPermission    |"
			Write-Host -ForegroundColor Yellow "|    Get-MailboxFolderStatistics    |    Get-EXOMailboxFolderStatistics    |"
			Write-Host -ForegroundColor Yellow "|    Get-MailboxPermission          |    Get-EXOMailboxPermission          |"
			Write-Host -ForegroundColor Yellow "|    Get-MailboxStatistics          |    Get-EXOMailboxStatistics          |"
			Write-Host -ForegroundColor Yellow "|    Get-MobileDeviceStatistics     |    Get-EXOMobileDeviceStatistics     |"
			Write-Host -ForegroundColor Yellow "|    Get-Recipient                  |    Get-EXORecipient                  |"
			Write-Host -ForegroundColor Yellow "|    Get-RecipientPermission        |    Get-EXORecipientPermission        |"
			Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
			Write-Host -ForegroundColor Yellow ""
			Write-Host -ForegroundColor Yellow "To get additional information, run: Get-Help Connect-ExchangeOnline or check https://aka.ms/exops-docs"
			Write-Host -ForegroundColor Yellow ""
			Write-Host -ForegroundColor Yellow "Send your product improvement suggestions and feedback to exocmdletpreview@service.microsoft.com. For issues related to the module, contact Microsoft support. Don't use the feedback alias for problems or support issues."
			Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------"
			Write-Host -ForegroundColor Yellow ""
		
		}
		else
		{
	       Write-Host -ForegroundColor Yellow ""
		   Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------------------"
		   Write-Host -ForegroundColor Yellow "This version of EXO PowerShell V2 module contains new REST API backed cmdlets which doesn't require WinRM for Client-Server communication and therefore you can now run these cmdlets after turning off WinRM Basic Auth in your client machine thus making it more secure."
           Write-Host -ForegroundColor Yellow ""
           Write-Host -ForegroundColor Yellow "The cmdlets in the downloaded module are resilient to transient failures, handling retries and throttling errors inherently."
           Write-Host -ForegroundColor Yellow ""
           Write-Host -ForegroundColor Yellow "The service side execution of these cmdlets have been optimized to perform better than the RPS (V1) cmdlets."
           Write-Host -ForegroundColor Yellow ""
           Write-Host -ForegroundColor Yellow "Unlike the EXO* prefixed cmdlets, the cmdlets in this module supports full functional parity with the RPS (v1) cmdlets."
           Write-Host -ForegroundColor Yellow ""
           Write-Host -ForegroundColor Yellow "For more information check https://aka.ms/exov2-module"
           Write-Host -ForegroundColor Yellow ""
           Write-Host -ForegroundColor Yellow "Send your issues, product improvement suggestions and feedback to exocmdletpreview@service.microsoft.com"
		   Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------------------"
		   Write-Host -ForegroundColor Yellow ""
		
		}

    }

    <#
    Get the ExchangeOnlineManagement module version.
    Same function is present in the autogen module. Both the codes should be kept in sync.
    #>
    function Get-ModuleVersion
    {
        try
        {
            # Return the already computed version info if available.
            if ($script:ModuleVersion -ne $null -and $script:ModuleVersion -ne '')
            {
                Write-Verbose "Returning precomputed version info: $script:ModuleVersion"
                return $script:ModuleVersion;
            }

            $exoModule = Get-Module ExchangeOnlineManagement
            # Get the module version from the loaded module info.
            $script:ModuleVersion = $exoModule.Version.ToString()

            # Look for prerelease information from the corresponding module manifest.
            $exoModuleRoot = (Get-Item $exoModule.Path).Directory.Parent.FullName

            $exoModuleManifestPath = Join-Path -Path $exoModuleRoot -ChildPath ExchangeOnlineManagement.psd1
            $isExoModuleManifestPathValid = Test-Path -Path $exoModuleManifestPath
            if ($isExoModuleManifestPathValid -ne $true)
            {
                # Could be a local debug build import for testing. Skip extracting prerelease info for those.
                Write-Verbose "Module manifest path invalid, path: $exoModuleManifestPath, skipping extracting prerelease info"
                return $script:ModuleVersion
            }

            $exoModuleManifestContent = Get-Content -Path $exoModuleManifestPath
            $preReleaseInfo = $exoModuleManifestContent -match "Prerelease = '(.*)'"
            if ($preReleaseInfo -ne $null)
            {
                $script:ModuleVersion = "{0}-{1}" -f $exoModule.Version.ToString(),$preReleaseInfo[0].Split('=')[1].Trim().Trim("'")
            }

            Write-Verbose "Computed version info: $script:ModuleVersion"
            return $script:ModuleVersion
        }
        catch
        {
            return [string]::Empty
        }
    }

    <#
    .Synopsis Validates a given Uri
    #>
    function Test-Uri
    {
        [CmdletBinding()]
        [OutputType([bool])]
        Param
        (
            # Uri to be validated
            [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
            [string]
            $UriString
        )

        [Uri]$uri = $UriString -as [Uri]

        $uri.AbsoluteUri -ne $null -and $uri.Scheme -eq 'https'
    }

    <#
    .Synopsis Is Cloud Shell Environment
    #>
    function global:IsCloudShellEnvironment()
    {
        return [Microsoft.Exchange.Management.AdminApiProvider.Utility]::IsCloudShellEnvironment();
    }

    <#
    .Synopsis Override Get-PSImplicitRemotingSession function for reconnection
    #>
    function global:UpdateImplicitRemotingHandler()
    {
        # Remote Powershell Sessions created by the ExchangeOnlineManagement module are given a name that starts with "ExchangeOnlineInternalSession".
        # Only modules from such sessions should be modified here, to prevent modfification of RPS tmp_* modules created by running the New-PSSession cmdlet directly, or when connecting to exchange on-prem tenants.
        $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}

        if ($existingPSSession.count -gt 0) 
        {
            foreach ($session in $existingPSSession)
            {
                $module = Get-Module $session.CurrentModuleName
                if ($module -eq $null)
                {
                    continue
                }

                [bool]$moduleProcessed = $false
                [string] $moduleUrl = $module.Description
                [int] $queryStringIndex = $moduleUrl.IndexOf("?")

                if ($queryStringIndex -gt 0)
                {
                    $moduleUrl = $moduleUrl.SubString(0,$queryStringIndex)
                }

                if ($moduleUrl.EndsWith("/PowerShell-LiveId", [StringComparison]::OrdinalIgnoreCase) -or $moduleUrl.EndsWith("/PowerShell", [StringComparison]::OrdinalIgnoreCase))
                {
                    & $module { ${function:Get-PSImplicitRemotingSession} = `
                    {
                        param(
                            [Parameter(Mandatory = $true, Position = 0)]
                            [string]
                            $commandName
                        )

                        $shouldRemoveCurrentSession = $false;
                        # Clear any left over PS tmp modules
                        if (($script:PSSession -ne $null) -and ($script:PSSession.PreviousModuleName -ne $null) -and ($script:PSSession.PreviousModuleName -ne $script:MyModule.Name))
                        {
                            $null = Remove-Module -Name $script:PSSession.PreviousModuleName -ErrorAction SilentlyContinue
                            $script:PSSession.PreviousModuleName = $null
                        }

                        if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                        {
                            Set-PSImplicitRemotingSession `
                                (& $script:GetPSSession `
                                    -InstanceId $script:PSSession.InstanceId.Guid `
                                    -ErrorAction SilentlyContinue )
                        }
                        if ($script:PSSession -ne $null)
                        {
                            if ($script:PSSession.Runspace.RunspaceStateInfo.State -eq 'Disconnected')
                            {
                                # If we are handed a disconnected session, try re-connecting it before creating a new session.
                                Set-PSImplicitRemotingSession `
                                    (& $script:ConnectPSSession `
                                        -Session $script:PSSession `
                                        -ErrorAction SilentlyContinue)
                            }
                            else
                            {
                                # Import the module once more to ensure that Test-ActiveToken is present
                                Import-Module $global:_EXO_ModulePath -Cmdlet Test-ActiveToken;

                                # If there is no active token run the new session flow
                                $hasActiveToken = Test-ActiveToken -TokenExpiryTime $script:PSSession.TokenExpiryTime
                                $sessionIsOpened = $script:PSSession.Runspace.RunspaceStateInfo.State -eq 'Opened'
                                if (($hasActiveToken -eq $false) -or ($sessionIsOpened -ne $true))
                                {
                                    #If there is no active user token or opened session then ensure that we remove the old session
                                    $shouldRemoveCurrentSession = $true;
                                }
                            }
                        }
                        if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened') -or ($shouldRemoveCurrentSession -eq $true))
                        {
                            # Import the module once more to ensure that New-ExoPSSession is present
                            Import-Module $global:_EXO_ModulePath -Cmdlet New-ExoPSSession;

                            Write-PSImplicitRemotingMessage ('Creating a new Remote PowerShell session using Modern Authentication for implicit remoting of "{0}" command ...' -f $commandName)
                            $session = New-ExoPSSession -PreviousSession $script:PSSession

                            if ($session -ne $null)
                            {
                                if ($shouldRemoveCurrentSession -eq $true)
                                {
                                    Remove-PSSession $script:PSSession
                                }

                                # Import the latest session to ensure that the next cmdlet call would occur on the new PSSession instance.
                                if ([string]::IsNullOrEmpty($script:MyModule.ModulePrefix))
                                {
                                    $PSSessionModuleInfo = Import-PSSession $session -AllowClobber -DisableNameChecking -CommandName $script:MyModule.CommandName -FormatTypeName $script:MyModule.FormatTypeName
                                }
                                else
                                {
                                    $PSSessionModuleInfo = Import-PSSession $session -AllowClobber -DisableNameChecking -CommandName $script:MyModule.CommandName -FormatTypeName $script:MyModule.FormatTypeName -Prefix $script:MyModule.ModulePrefix
                                }

                                # Add the name of the module to clean up in case of removing the broken session
                                $session | Add-Member -NotePropertyName "CurrentModuleName" -NotePropertyValue $PSSessionModuleInfo.Name

                                $CurrentModule = Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $script:MyModule.ModulePrefix -PassThru
                                $CurrentModule | Add-Member -NotePropertyName "ModulePrefix" -NotePropertyValue $script:MyModule.ModulePrefix
                                $CurrentModule | Add-Member -NotePropertyName "CommandName" -NotePropertyValue $script:MyModule.CommandName
                                $CurrentModule | Add-Member -NotePropertyName "FormatTypeName" -NotePropertyValue $script:MyModule.FormatTypeName

                                $session | Add-Member -NotePropertyName "PreviousModuleName" -NotePropertyValue $script:MyModule.Name

                                UpdateImplicitRemotingHandler
                                $script:PSSession = $session
                            }
                        }
                        if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                        {
                            throw 'No session has been associated with this implicit remoting module'
                        }

                        return [Management.Automation.Runspaces.PSSession]$script:PSSession
                    }}
                }
            }
        }
    }

    <#
    .SYNOPSIS Extract organization name from UserPrincipalName
    #>
    function Get-OrgNameFromUPN
    {
        param([string] $UPN)
        $fields = $UPN -split '@'
        return $fields[-1]
    }

    <#
    .SYNOPSIS Get the command from the given module
    #>
    function global:Get-WrappedCommand
    {
        param(
        [string] $CommandName,
        [string] $ModuleName,
        [string] $CommandType)

        $cmd = (Get-Module $moduleName).ExportedFunctions[$CommandName]
        return $cmd
    }

    <#
    .Synopsis Writes a message to the console in Yellow.
    Call this method with caution since it uses Write-Host internally which cannot be suppressed by the user.
    #>
    function Write-Message
    {
        param([string] $message)
        Write-Host $message -ForegroundColor Yellow
    }

############# Helper Functions End #############

###### Begin Main ######

$EOPConnectionInProgress = $false
function Connect-ExchangeOnline 
{
    [CmdletBinding()]
    param(

        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = '',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = '',

        # Exchange Environment name
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment] $ExchangeEnvironmentName = 'O365Default',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # Prefix 
        [string] $Prefix = '',

        # Show Banner of Exchange cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true,

        #Cmdlets to Import for rps cmdlets , by default it would bring all
        [string[]] $CommandName = @("*"),

        #The way the output objects would be printed on the console
        [string[]] $FormatTypeName = @("*"),

        # Use Remote PowerShell Session based connection
        [switch] $UseRPSSession = $false
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            # Certificate
            $Certificate = New-Object System.Management.Automation.RuntimeDefinedParameter('Certificate', [System.Security.Cryptography.X509Certificates.X509Certificate2], $attributeCollection)
            $Certificate.Value = $null

            # Certificate Path 
            $CertificateFilePath = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateFilePath', [string], $attributeCollection)
            $CertificateFilePath.Value = ''

            # Certificate Password
            $CertificatePassword = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificatePassword', [System.Security.SecureString], $attributeCollection)
            $CertificatePassword.Value = $null

            # Certificate Thumbprint
            $CertificateThumbprint = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateThumbprint', [string], $attributeCollection)
            $CertificateThumbprint.Value = ''

            # Application Id
            $AppId = New-Object System.Management.Automation.RuntimeDefinedParameter('AppId', [string], $attributeCollection)
            $AppId.Value = ''

            # Organization
            $Organization = New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], $attributeCollection)
            $Organization.Value = ''

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")

            # Create a new attribute and valiate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

# EXO params start

            # Switch to track perfomance 
            $TrackPerformance = New-Object System.Management.Automation.RuntimeDefinedParameter('TrackPerformance', [bool], $attributeCollection)
            $TrackPerformance.Value = $false

            # Flag to enable or disable showing the number of objects written
            $ShowProgress = New-Object System.Management.Automation.RuntimeDefinedParameter('ShowProgress', [bool], $attributeCollection)
            $ShowProgress.Value = $false

            # Switch to enable/disable Multi-threading in the EXO cmdlets
            $UseMultithreading = New-Object System.Management.Automation.RuntimeDefinedParameter('UseMultithreading', [bool], $attributeCollection)
            $UseMultithreading.Value = $true

            # Pagesize Param
            $PageSize = New-Object System.Management.Automation.RuntimeDefinedParameter('PageSize', [uint32], $attributeCollection)
            $PageSize.Value = 1000

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to CmdInline parameters
            $InlineCredential = New-Object System.Management.Automation.RuntimeDefinedParameter('InlineCredential', [switch], $attributeCollection)
            $InlineCredential.Value = $false

# EXO params end
            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('Certificate', $Certificate)
            $paramDictionary.Add('CertificateFilePath', $CertificateFilePath)
            $paramDictionary.Add('CertificatePassword', $CertificatePassword)
            $paramDictionary.Add('AppId', $AppId)
            $paramDictionary.Add('Organization', $Organization)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('TrackPerformance', $TrackPerformance)
            $paramDictionary.Add('ShowProgress', $ShowProgress)
            $paramDictionary.Add('UseMultithreading', $UseMultithreading)
            $paramDictionary.Add('PageSize', $PageSize)
            if($PSEdition -eq 'Core')
            {
                $paramDictionary.Add('Device', $Device)
                $paramDictionary.Add('InlineCredential', $InlineCredential);
                # We do not want to expose certificate thumprint in Linux as it is not feasible there.
                if($IsWindows)
                {
                    $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
                }
            }
            else 
            {
                $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
            }

            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")

            # Create a new attribute and validate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

            # Switch to CmdInline parameters
            $InlineCredential = New-Object System.Management.Automation.RuntimeDefinedParameter('InlineCredential', [switch], $attributeCollection)
            $InlineCredential.Value = $false

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('InlineCredential', $InlineCredential)
            return $paramDictionary
        }
    }
    process {

        $startTime = Get-Date

        # Validate parameters
        if (($ConnectionUri -ne '') -and (-not (Test-Uri $ConnectionUri)))
        {
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri -ne '') -and (-not (Test-Uri $AzureADAuthorizationEndpointUri)))
        {
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }
        if (($Prefix -ne ''))
        {
            if ($Prefix -notmatch '^[a-z0-9]+$') 
            {
                throw "Use of any special characters in the Prefix string is not supported."
            }
            if ($Prefix -eq 'EXO') 
            {
                throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
            }
        }

        # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        try
        {
            $moduleVersion = Get-ModuleVersion
            Write-Verbose "ModuleVersion: $moduleVersion"

            # Generate a ConnectionId to use in all logs and to send in all server calls.
            $connectionContextID = [System.Guid]::NewGuid()

            $cmdletLogger = New-CmdletLogger -ExoModuleVersion $moduleVersion -LogDirectoryPath $LogDirectoryPath.Value -EnableErrorReporting:$EnableErrorReporting.Value -ConnectionId $connectionContextID -IsRpsSession:$UseRPSSession.IsPresent
            $logFilePath = $cmdletLogger.GetCurrentLogFilePath()

            $cmdletLogger.InitLog($connectionContextID)
            $cmdletLogger.LogStartTime($connectionContextID, $startTime)
            $cmdletLogger.LogCmdletName($connectionContextID, "Connect-ExchangeOnline");
            $cmdletLogger.LogCmdletParameters($connectionContextID, $PSBoundParameters);

            if ($isCloudShell -eq $false)
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value `
                -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring `
                -DelegatedOrg $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value `
                -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $CertificateThumbprint.Value -AppId $AppId.Value `
                -Organization $Organization.Value -Device:$Device.Value -InlineCredential:$InlineCredential.Value -CommandName $CommandName `
                -FormatTypeName $FormatTypeName -Prefix $Prefix -PageSize $PageSize.Value -ExoModuleVersion:$moduleVersion -Logger $cmdletLogger `
                -ConnectionId $connectionContextID -IsRpsSession $UseRPSSession.IsPresent -EnableErrorReporting:$EnableErrorReporting.Value
            }
            else
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -Credential $Credential.Value -PSSessionOption $PSSessionOption `
                -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -InlineCredential:$InlineCredential.Value `
                -DelegatedOrg $DelegatedOrganization -CommandName $CommandName -FormatTypeName $FormatTypeName -Prefix $prefix -ExoModuleVersion:$moduleVersion `
                -Logger $cmdletLogger -ConnectionId $connectionContextID -IsRpsSession $UseRPSSession.IsPresent -EnableErrorReporting:$EnableErrorReporting.Value
            }

            if ($isCloudShell -eq $false)
            {
                $global:_EXO_EnableErrorReporting = $EnableErrorReporting.Value;
            }

            if ($ShowBanner -eq $true)
            {
                try
                {
                    $BannerContent = Get-EXOBanner -ConnectionContext:$ConnectionContext -IsRPSSession:$UseRPSSession.IsPresent
                    Write-Host -ForegroundColor Yellow $BannerContent
                }
                catch
                {
                    Write-Verbose "Failed to fetch banner content from server, using the default banner content. Reason: $_"
                    $cmdletLogger.LogGenericError($connectionContextID, $_);

                    Print-Details;
                }
            }

            if (($ConnectionUri -ne '') -and ($AzureADAuthorizationEndpointUri -eq ''))
            {
                Write-Information "Using ConnectionUri:'$ConnectionUri', in the environment:'$ExchangeEnvironmentName'."
            }
            if (($AzureADAuthorizationEndpointUri -ne '') -and ($ConnectionUri -eq ''))
            {
                Write-Information "Using AzureADAuthorizationEndpointUri:'$AzureADAuthorizationEndpointUri', in the environment:'$ExchangeEnvironmentName'."
            }
            
            if ($EnableErrorReporting.Value -eq $true -and $UseRPSSession -eq $false)
            {
                Write-Message ("Writing cmdlet logs to " + $logFilePath)
            }

            $ImportedModuleName = '';
            $LogModuleDirectoryPath = [System.IO.Path]::GetTempPath();

            if ($UseRPSSession -eq $true)
            {
                $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
                $ModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule);

                Import-Module $ModulePath;

                $global:_EXO_ModulePath = $ModulePath;

                $PSSession = New-ExoPSSession -ConnectionContext $ConnectionContext

                if ($PSSession -ne $null)
                {
                    if ([string]::IsNullOrEmpty($Prefix))
                    {
                        $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking -CommandName $CommandName -FormatTypeName $FormatTypeName
                    }
                    else
                    {
                        $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking -CommandName $CommandName -FormatTypeName $FormatTypeName -Prefix $Prefix
                    }
                    # Add the name of the module to clean up in case of removing the broken session
                    $PSSession | Add-Member -NotePropertyName "CurrentModuleName" -NotePropertyValue $PSSessionModuleInfo.Name

                    # Import the above module globally. This is needed as with using psm1 files, 
                    # any module which is dynamically loaded in the nested module does not reflect globally.
                    $CurrentModule = Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $Prefix -PassThru
                    $CurrentModule | Add-Member -NotePropertyName "ModulePrefix" -NotePropertyValue $Prefix
                    $CurrentModule | Add-Member -NotePropertyName "CommandName" -NotePropertyValue $CommandName
                    $CurrentModule | Add-Member -NotePropertyName "FormatTypeName" -NotePropertyValue $FormatTypeName

                    UpdateImplicitRemotingHandler

                    # Import the REST module
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    $RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings;

                    $ImportedModuleName = $PSSessionModuleInfo.Name;
                }
            }
            else
            {
                # Download the new web based EXOModule
                $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext;
                if ($null -ne $ImportedModule)
                {
                    $ImportedModuleName = $ImportedModule.Name;
                    $LogModuleDirectoryPath = $ImportedModule.ModuleBase

                    Write-Verbose "AutoGen EXOModule created at  $($ImportedModule.ModuleBase)"
                    if ($null -ne $HelpFileNames -and $HelpFileNames -is [array] -and $HelpFileNames.Count -gt 0)
                    {
                        Get-HelpFiles -HelpFileNames $HelpFileNames -ConnectionContext $ConnectionContext -ImportedModule $ImportedModule -EnableErrorReporting:$EnableErrorReporting.Value
                    }
                    else
                    {
                        Write-Warning "Get-help cmdlet might not work. Please reconnect if you want to access that functionality."
                    }
                }
                else
                {
                    throw "Module could not be correctly formed. Please run Connect-ExchangeOnline again."
                }
            }

            # If we are configured to collect telemetry, add telemetry wrappers in case of an RPS connection
            if ($EnableErrorReporting.Value -eq $true)
            {
                if ($UseRPSSession -eq $true)
                {
                    $FilePath = Add-EXOClientTelemetryWrapper -Organization (Get-OrgNameFromUPN -UPN $UserPrincipalName.Value) -PSSessionModuleName $ImportedModuleName -LogDirectoryPath $LogDirectoryPath.Value -LogModuleDirectoryPath $LogModuleDirectoryPath
                    Write-Message("Writing telemetry records for this session to " + $FilePath[0] );
                    $global:_EXO_TelemetryFilePath = $FilePath[0]
                    Import-Module $FilePath[1] -DisableNameChecking -Global

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }

                $endTime = Get-Date
                $cmdletLogger.LogEndTime($connectionContextID, $endTime);
                $cmdletLogger.CommitLog($connectionContextID);

                if ($EOPConnectionInProgress -eq $false)
                {
                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }
            }
            else 
            {
                if ($EOPConnectionInProgress -eq $false)
                {
                    # Set the AppSettings disabling the logging
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -EnableErrorReporting $false
                }
            }
        }
        catch
        {
            # If telemetry is enabled, log errors generated from this cmdlet also.
            # If telemetry is not enabled, calls to cmdletLogger will be a no-op.
            $errorCountAtProcessEnd = $global:Error.Count 
            $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
            for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
            {
                $cmdletLogger.LogGenericError($connectionContextID, $global:Error[$i]);
            }

            $cmdletLogger.CommitLog($connectionContextID);

            if ($EnableErrorReporting.Value -eq $true -and $UseRPSSession -eq $true)
            {
                if ($global:_EXO_TelemetryFilePath -eq $null)
                {
                    $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath -LogDirectoryPath $LogDirectoryPath.Value

                    # Import the REST module
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    $RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings;

                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ConnectionUri $ConnectionUri -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }

                # Log errors which are encountered during Connect-ExchangeOnline execution. 
                Write-Message ("Writing Connect-ExchangeOnline error log to " + $global:_EXO_TelemetryFilePath)
                Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
            }

            if ($_.Exception -ne $null)
            {
                if ($_.Exception.InnerException -ne $null)
                {
                    throw $_.Exception.InnerException;
                }
                else
                {
                    throw $_.Exception;
                }
            }
            else
            {
                throw $_;
            }
        }
    }
}

function Connect-IPPSSession
{
    [CmdletBinding()]
    param(
        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations',

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Prefix 
        [string] $Prefix = '',

        #Cmdlets to Import, by default it would bring all
        [string[]] $CommandName = @("*"),

        #The way the output objects would be printed on the console
        [string[]] $FormatTypeName = @("*")
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            # Certificate
            $Certificate = New-Object System.Management.Automation.RuntimeDefinedParameter('Certificate', [System.Security.Cryptography.X509Certificates.X509Certificate2], $attributeCollection)
            $Certificate.Value = $null

            # Certificate Path 
            $CertificateFilePath = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateFilePath', [string], $attributeCollection)
            $CertificateFilePath.Value = ''

            # Certificate Password
            $CertificatePassword = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificatePassword', [System.Security.SecureString], $attributeCollection)
            $CertificatePassword.Value = $null

            # Certificate Thumbprint
            $CertificateThumbprint = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateThumbprint', [string], $attributeCollection)
            $CertificateThumbprint.Value = ''

            # Application Id
            $AppId = New-Object System.Management.Automation.RuntimeDefinedParameter('AppId', [string], $attributeCollection)
            $AppId.Value = ''

            # Organization
            $Organization = New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], $attributeCollection)
            $Organization.Value = ''

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('Certificate', $Certificate)
            $paramDictionary.Add('CertificateFilePath', $CertificateFilePath)
            $paramDictionary.Add('CertificatePassword', $CertificatePassword)
            $paramDictionary.Add('AppId', $AppId)
            $paramDictionary.Add('Organization', $Organization)
            if($PSEdition -eq 'Core')
            {
                # We do not want to expose certificate thumprint in Linux as it is not feasible there.
                if($IsWindows)
                {
                    $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
                }
            }
            else 
            {
                $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
            }

            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            return $paramDictionary
        }
    }
    process 
    {
        try
        {
            $EOPConnectionInProgress = $true
            if ($isCloudShell -eq $false)
            {
                Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$false -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $CertificateThumbprint.Value -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$true
            }
            else
            {
                Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -ShowBanner:$false -DelegatedOrganization $DelegatedOrganization -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$true
            }
        }
        finally
        {
            $EOPConnectionInProgress = $false
        }
    }
}

function Disconnect-ExchangeOnline
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    param()

    process
    {
        if ($PSCmdlet.ShouldProcess(
            "Running this cmdlet clears all active sessions created using Connect-ExchangeOnline or Connect-IPPSSession.",
            "Press(Y/y/A/a) if you want to continue.",
            "Running this cmdlet clears all active sessions created using Connect-ExchangeOnline or Connect-IPPSSession. "))
        {

            # Keep track of error count at beginning.
            $errorCountAtStart = $global:Error.Count;
            $startTime = Get-Date

            try
            {
                # Get all the connection contexts so that the logger can be initialized.
                $connectionContexts = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetAllConnectionContexts()
                $disconnectCmdletId = [System.Guid]::NewGuid().ToString()

                # denotes if any of the connections is an RPS session.
                # This is used to Push-EXOTelemetryRecord in case any RPS connections are present.
                $rpsConnectionWithErrorReportingExists = $false
                
                foreach ($context in $connectionContexts)
                {
                    $context.Logger.InitLog($disconnectCmdletId);
                    $context.Logger.LogStartTime($disconnectCmdletId, $startTime);
                    $context.Logger.LogCmdletName($disconnectCmdletId, "Disconnect-ExchangeOnline");
                    if ($context.IsRpsSession -and $context.ErrorReportingEnabled)
                    {
                        $rpsConnectionWithErrorReportingExists = $true
                    }
                }

                # Import the module once more to ensure that Test-ActiveToken is present
                $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
                $ModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule);
                Import-Module $ModulePath -Cmdlet Clear-ActiveToken;

                $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}

                if ($existingPSSession.count -gt 0) 
                {
                    for ($index = 0; $index -lt $existingPSSession.count; $index++)
                    {
                        $session = $existingPSSession[$index]
                        Remove-PSSession -session $session

                        Write-Information "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"

                        # Remove any active access token from the cache
                        Clear-ActiveToken -TokenProvider $session.TokenProvider

                        # Remove any previous modules loaded because of the current PSSession
                        if ($session.PreviousModuleName -ne $null)
                        {
                            if ((Get-Module $session.PreviousModuleName).Count -ne 0)
                            {
                                $null = Remove-Module -Name $session.PreviousModuleName -ErrorAction SilentlyContinue
                            }

                            $session.PreviousModuleName = $null
                        }

                        # Remove any leaked module in case of removal of broken session object
                        if ($session.CurrentModuleName -ne $null)
                        {
                            if ((Get-Module $session.CurrentModuleName).Count -ne 0)
                            {
                                $null = Remove-Module -Name $session.CurrentModuleName -ErrorAction SilentlyContinue
                            }
                        }
                    }
                }

                # Remove all the AutoREST modules from this instance of powershell if created
                $existingAutoRESTModules = Get-Module "tmpEXO_*"
                foreach ($module in $existingAutoRESTModules)
                {
                    $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                }

                Write-Information "Disconnected successfully !"

                # Remove all the Wrapped modules from this instance of powershell if created	
                $existingWrappedModules = Get-Module "EXOCmdletWrapper-*"	
                foreach ($module in $existingWrappedModules)	
                {
                    $null = Remove-Module -Name $module -ErrorAction SilentlyContinue	
                }

                # Remove all ConnectionContexts
                # this internally clears all the active tokens in ConnectionContexts
                [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveAllConnectionContexts()

                if ($rpsConnectionWithErrorReportingExists)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }
            }
            catch
            {
                # If telemetry is enabled for any of the connections, log errors generated from this cmdlet also. 
                $errorCountAtProcessEnd = $global:Error.Count 

                $endTime = Get-Date
                foreach ($context in $connectionContexts)
                {
                    $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
                    for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
                    {
                        $cmdletLogger.LogGenericError($disconnectCmdletId, $global:Error[$i]);                        
                    }

                    $context.Logger.LogEndTime($disconnectCmdletId, $endTime);
                    $context.Logger.CommitLog($disconnectCmdletId);
                    $logFilePath = $context.Logger.GetCurrentLogFilePath();
                }

                if ($rpsConnectionWithErrorReportingExists)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    # Log errors which are encountered during Disconnect-ExchangeOnline execution. 
                    Write-Message ("Writing Disconnect-ExchangeOnline errors to " + $global:_EXO_TelemetryFilePath)

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
                }

                throw $_
            }

            $endTime = Get-Date
            foreach ($context in $connectionContexts)
            {
                if ($context.Logger -ne $null)
                {
                    $context.Logger.LogEndTime($disconnectCmdletId, $endTime);
                    $context.Logger.CommitLog($disconnectCmdletId);
                }
            }
        }
    }
}

# SIG # Begin signature block
# MIIntwYJKoZIhvcNAQcCoIInqDCCJ6QCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB8GMSlLIrWkzXU
# MqnxI6xeesBh4hEWquQYRbUniIqUFaCCDYEwggX/MIID56ADAgECAhMzAAACUosz
# qviV8znbAAAAAAJSMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjEwOTAyMTgzMjU5WhcNMjIwOTAxMTgzMjU5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDQ5M+Ps/X7BNuv5B/0I6uoDwj0NJOo1KrVQqO7ggRXccklyTrWL4xMShjIou2I
# sbYnF67wXzVAq5Om4oe+LfzSDOzjcb6ms00gBo0OQaqwQ1BijyJ7NvDf80I1fW9O
# L76Kt0Wpc2zrGhzcHdb7upPrvxvSNNUvxK3sgw7YTt31410vpEp8yfBEl/hd8ZzA
# v47DCgJ5j1zm295s1RVZHNp6MoiQFVOECm4AwK2l28i+YER1JO4IplTH44uvzX9o
# RnJHaMvWzZEpozPy4jNO2DDqbcNs4zh7AWMhE1PWFVA+CHI/En5nASvCvLmuR/t8
# q4bc8XR8QIZJQSp+2U6m2ldNAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUNZJaEUGL2Guwt7ZOAu4efEYXedEw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDY3NTk3MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAFkk3
# uSxkTEBh1NtAl7BivIEsAWdgX1qZ+EdZMYbQKasY6IhSLXRMxF1B3OKdR9K/kccp
# kvNcGl8D7YyYS4mhCUMBR+VLrg3f8PUj38A9V5aiY2/Jok7WZFOAmjPRNNGnyeg7
# l0lTiThFqE+2aOs6+heegqAdelGgNJKRHLWRuhGKuLIw5lkgx9Ky+QvZrn/Ddi8u
# TIgWKp+MGG8xY6PBvvjgt9jQShlnPrZ3UY8Bvwy6rynhXBaV0V0TTL0gEx7eh/K1
# o8Miaru6s/7FyqOLeUS4vTHh9TgBL5DtxCYurXbSBVtL1Fj44+Od/6cmC9mmvrti
# yG709Y3Rd3YdJj2f3GJq7Y7KdWq0QYhatKhBeg4fxjhg0yut2g6aM1mxjNPrE48z
# 6HWCNGu9gMK5ZudldRw4a45Z06Aoktof0CqOyTErvq0YjoE4Xpa0+87T/PVUXNqf
# 7Y+qSU7+9LtLQuMYR4w3cSPjuNusvLf9gBnch5RqM7kaDtYWDgLyB42EfsxeMqwK
# WwA+TVi0HrWRqfSx2olbE56hJcEkMjOSKz3sRuupFCX3UroyYf52L+2iVTrda8XW
# esPG62Mnn3T8AuLfzeJFuAbfOSERx7IFZO92UPoXE1uEjL5skl1yTZB3MubgOA4F
# 8KoRNhviFAEST+nG8c8uIsbZeb08SeYQMqjVEmkwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIZjDCCGYgCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAlKLM6r4lfM52wAAAAACUjAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgW3ejCmlF
# MtrEUQ4Zs2uOdEM0Ig7MtwMH+WRDbbc/j7kwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQAp6ClgOnug/52SIOxLX8oYmkhzjr47JnUja7Gkdwwm
# sWhcZ6tS9IL9YgCt87q+dOd85qMI/Tm3EH/QiWd1BRQ77A5ZdGFN42q+R6jTQ/Dy
# 75OB1pJ0wbq/+jCT2VQD1qqfCcXith0bbTbRQF4TP30Wyh2qp/pQiLEuJ8RpZhx/
# YAacO8SzoJNIwg/IZWC9DX8EG07pgYqTy5ofRelRFTEo23M8W2xPj7oh4KLTqjjm
# BBDnFugV/+0lniZ6gt4YrBimrNISrOODI5cSrxVkT/Kl0kh6/+A1p6TXQJ0Bh+im
# jjQbm2fvpg2EKaF02jSjimMoprMZVmbrQnLXnuOWMS4foYIXFjCCFxIGCisGAQQB
# gjcDAwExghcCMIIW/gYJKoZIhvcNAQcCoIIW7zCCFusCAQMxDzANBglghkgBZQME
# AgEFADCCAVkGCyqGSIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIGVqIsrtQpVW6qTMc54PIKrnvc45tsjR+grwBONf
# 4kckAgZibEQ4hmEYEzIwMjIwNTE0MDIzMzEwLjE3NFowBIACAfSggdikgdUwgdIx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1p
# Y3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhh
# bGVzIFRTUyBFU046MkFENC00QjkyLUZBMDExJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2WgghFlMIIHFDCCBPygAwIBAgITMwAAAYZ45RmJ+CRL
# zAABAAABhjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMDAeFw0yMTEwMjgxOTI3MzlaFw0yMzAxMjYxOTI3MzlaMIHSMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
# SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOjJBRDQtNEI5Mi1GQTAxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAwI3G2Wpv
# 6B4IjAfrgfJpndPOPYO1Yd8+vlfoIxMW3gdCDT+zIbafg14pOu0t0ekUQx60p7Pa
# dH4OjnqNIE1q6ldH9ntj1gIdl4Hq4rdEHTZ6JFdE24DSbVoqqR+R4Iw4w3GPbfc2
# Q3kfyyFyj+DOhmCWw/FZiTVTlT4bdejyAW6r/Jn4fr3xLjbvhITatr36VyyzgQ0Y
# 4Wr73H3gUcLjYu0qiHutDDb6+p+yDBGmKFznOW8wVt7D+u2VEJoE6JlK0EpVLZus
# dSzhecuUwJXxb2uygAZXlsa/fHlwW9YnlBqMHJ+im9HuK5X4x8/5B5dkuIoX5lWG
# jFMbD2A6Lu/PmUB4hK0CF5G1YaUtBrME73DAKkypk7SEm3BlJXwY/GrVoXWYUGEH
# yfrkLkws0RoEMpoIEgebZNKqjRynRJgR4fPCKrEhwEiTTAc4DXGci4HHOm64EQ1g
# /SDHMFqIKVSxoUbkGbdKNKHhmahuIrAy4we9s7rZJskveZYZiDmtAtBt/gQojxbZ
# 1vO9C11SthkrmkkTMLQf9cDzlVEBeu6KmHX2Sze6ggne3I4cy/5IULnHZ3rM4ZpJ
# c0s2KpGLHaVrEQy4x/mAn4yaYfgeH3MEAWkVjy/qTDh6cDCF/gyz3TaQDtvFnAK7
# 0LqtbEvBPdBpeCG/hk9l0laYzwiyyGY/HqMCAwEAAaOCATYwggEyMB0GA1UdDgQW
# BBQZtqNFA+9mdEu/h33UhHMN6whcLjAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAx
# MCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3Rh
# bXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMA0GCSqGSIb3DQEBCwUAA4ICAQDD7mehJY3fTHKC4hj+wBWB8544
# uaJiMMIHnhK9ONTM7VraTYzx0U/TcLJ6gxw1tRzM5uu8kswJNlHNp7RedsAiwviV
# QZV9AL8IbZRLJTwNehCwk+BVcY2gh3ZGZmx8uatPZrRueyhhTTD2PvFVLrfwh2li
# DG/dEPNIHTKj79DlEcPIWoOCUp7p0ORMwQ95kVaibpX89pvjhPl2Fm0CBO3pXXJg
# 0bydpQ5dDDTv/qb0+WYF/vNVEU/MoMEQqlUWWuXECTqx6TayJuLJ6uU7K5QyTkQ/
# l24IhGjDzf5AEZOrINYzkWVyNfUOpIxnKsWTBN2ijpZ/Tun5qrmo9vNIDT0lobgn
# ulae17NaEO9oiEJJH1tQ353dhuRi+A00PR781iYlzF5JU1DrEfEyNx8CWgERi90L
# KsYghZBCDjQ3DiJjfUZLqONeHrJfcmhz5/bfm8+aAaUPpZFeP0g0Iond6XNk4YiY
# bWPFoofc0LwcqSALtuIAyz6f3d+UaZZsp41U4hCIoGj6hoDIuU839bo/mZ/AgESw
# GxIXs0gZU6A+2qIUe60QdA969wWSzucKOisng9HCSZLF1dqc3QUawr0C0U41784K
# o9vckAG3akwYuVGcs6hM/SqEhoe9jHwe4Xp81CrTB1l9+EIdukCbP0kyzx0WZzte
# eiDN5rdiiQR9mBJuljCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUw
# DQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhv
# cml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg
# 4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aO
# RmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41
# JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5
# LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL
# 64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9
# QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj
# 0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqE
# UUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0
# kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435
# UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB
# 3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTE
# mr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwG
# A1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNV
# HSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNV
# HQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo
# 0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29m
# dC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5j
# cmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDAN
# BgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4
# sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th54
# 2DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRX
# ud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBew
# VIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0
# DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+Cljd
# QDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFr
# DZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFh
# bHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7n
# tdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+
# oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6Fw
# ZvKhggLUMIICPQIBATCCAQChgdikgdUwgdIxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046MkFENC00Qjky
# LUZBMDExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoB
# ATAHBgUrDgMCGgMVAAGu2DRzWkKljmXySX1korHL4fMnoIGDMIGApH4wfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQEFBQACBQDmKTbWMCIY
# DzIwMjIwNTE0MDM1ODE0WhgPMjAyMjA1MTUwMzU4MTRaMHQwOgYKKwYBBAGEWQoE
# ATEsMCowCgIFAOYpNtYCAQAwBwIBAAICB64wBwIBAAICEV8wCgIFAOYqiFYCAQAw
# NgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgC
# AQACAwGGoDANBgkqhkiG9w0BAQUFAAOBgQCrelynmZ8zOl0MQImkHwwEGc9us2B3
# /3I0jBc00VTiw55ZaY4+drd2+dGr5ESofWEA4jcpu2BGlIb+AHuo256UWgm5W/oa
# q4qkSUt7cQivHjtuXUFqb54lhn5UUAsQ0QR/j3erGivpLj9HpmL+45hsh+WnCBoP
# sk3m/mW+8HGj5jGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAABhnjlGYn4JEvMAAEAAAGGMA0GCWCGSAFlAwQCAQUAoIIBSjAa
# BgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIC3RGVPj
# 7nn1+YsxlwOd3c8EdtDT0uBoYOBvsRRZvABEMIH6BgsqhkiG9w0BCRACLzGB6jCB
# 5zCB5DCBvQQgGpmI4LIsCFTGiYyfRAR7m7Fa2guxVNIw17mcAiq8Qn4wgZgwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAYZ45RmJ+CRLzAAB
# AAABhjAiBCAaMlDaNsIjvQwjBYWeFngI27RvdgVfTcvNFACIHYqE8DANBgkqhkiG
# 9w0BAQsFAASCAgChcZkwRc4ywPBEknpTkCZ6SBeUlFCyRNNeqwTxaO1lxcrpvnL4
# 1KvIHknAGb4ize91uzeDTX0fIlsGLAO96Ut5/ahpd9P5C056lZlNF76kCqLRA4sB
# ESWgmtAFfB3hMOyr7x/hzeLbnC6+epSiBeFzZvtuqGcFCFvdx0ZSHkJpHD8P2KiG
# miENUbb6MGUxCDvZ28TT3MEpuAPoZ7GuMy/pDlBHGZdUoKYoRqAljEOpKo9mlgez
# ioujgPo/ava8QvnS9NRZRPf1JDObhMSQnQAby8mcCeh9rlj/MoApf1TGj4QSqSMH
# yuUQNke+bWK4TCTKZK2QCIMxOc6kAcNBMG9MTNmlfwEa4nidER2QGl+QZkgqXSmh
# SkPPyU7jJqbmsLAkXDhbO9CAiYC83jRGT3+vHlHp5fZuLlqI6IZ/h2GGRy4o3nO+
# 0FPY7TWsNPu19Y0tGj/vvxdzqjjBPqPXNnSZ+uG7pzy3ww88JfFcnXGPfOliTDk2
# vuKRzyGDt44G1jqxk1rcbhQDWGRZNTw7PA/Z2U8BdS2y2m5tvqfbTEGg3HYOa0FN
# zi0QX545u6cAPucj4SxWXDwy3mSk7yUs+MIt1P813DPnuJMDAgw5VVM9g6V4ILP9
# 58qR4MFFe119ze88Yz2yhDtHv1ia5sUm7W7/HnksGilTD7hkVLIZeNGp9Q==
# SIG # End signature block
