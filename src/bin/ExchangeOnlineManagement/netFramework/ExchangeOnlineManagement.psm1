# Import the REST module so that the EXO* cmdlets are present before Connect-ExchangeOnline in the powershell instance.
$RestModule = "Microsoft.Exchange.Management.RestApiClient.dll"
$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestModule)
Import-Module $RestModulePath 

$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"
$ExoPowershellModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule)
Import-Module $ExoPowershellModulePath

#Keep track of Execution status of last cmdlet
$global:EXO_LastExecutionStatus = $true;

############# Helper Functions Begin #############

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
            
            # Check for ExchangeOnlineManagementBeta in case the psm1 is loaded directly
            if ($exoModule -eq $null)
            {
               $exoModule = (Get-Command -Name Connect-ExchangeOnline).Module
            }

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
        [switch] $UseRPSSession = $false,

        # Use to Skip Exchange Format file loading into current shell.
        [switch] $SkipLoadingFormatData = $false,

        # Use to skip downloading and loading the help files into the current connection.
        [switch] $SkipLoadingCmdletHelp = $false,

        # Externally provided access token
        [string] $AccessToken = '',

        # Client certificate to sign the temp module
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $SigningCertificate = $null
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

            # Switch to use Managed Identity flow. 
            $ManagedIdentity = New-Object System.Management.Automation.RuntimeDefinedParameter('ManagedIdentity', [switch], $attributeCollection)
            $ManagedIdentity.Value = $false

            # ManagedIdentityAccountId to be used in case of User Assigned Managed Identity flow
            $ManagedIdentityAccountId = New-Object System.Management.Automation.RuntimeDefinedParameter('ManagedIdentityAccountId', [string], $attributeCollection)
            $ManagedIdentityAccountId.Value = ''

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
            $paramDictionary.Add('ManagedIdentity', $ManagedIdentity)
            $paramDictionary.Add('ManagedIdentityAccountId', $ManagedIdentityAccountId)
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
        $global:EXO_LastExecutionStatus = $true;
        $startTime = Get-Date

        # Validate parameters
        if (($ConnectionUri -ne '') -and (-not (Test-Uri $ConnectionUri)))
        {
            $global:EXO_LastExecutionStatus = $false;
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri -ne '') -and (-not (Test-Uri $AzureADAuthorizationEndpointUri)))
        {
            $global:EXO_LastExecutionStatus = $false;
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }
        if (($Prefix -ne ''))
        {
            if ($Prefix -notmatch '^[a-z0-9]+$') 
            {
                $global:EXO_LastExecutionStatus = $false;
                throw "Use of any special characters in the Prefix string is not supported."
            }
            if ($Prefix -eq 'EXO') 
            {
                $global:EXO_LastExecutionStatus = $false;
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
            
            if ($EnableErrorReporting.Value -eq $true -and $UseRPSSession -eq $false)
            {
                Write-Message ("Writing cmdlet logs to " + $logFilePath)
            }

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
                -ConnectionId $connectionContextID -IsRpsSession $UseRPSSession.IsPresent -EnableErrorReporting:$EnableErrorReporting.Value `
                -ManagedIdentity:$ManagedIdentity.Value -ManagedIdentityAccountId $ManagedIdentityAccountId.Value -AccessToken $AccessToken
            }
            else
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -Credential $Credential.Value -PSSessionOption $PSSessionOption `
                -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -InlineCredential:$InlineCredential.Value `
                -DelegatedOrg $DelegatedOrganization -CommandName $CommandName -FormatTypeName $FormatTypeName -Prefix $prefix -ExoModuleVersion:$moduleVersion `
                -Logger $cmdletLogger -ConnectionId $connectionContextID -IsRpsSession $UseRPSSession.IsPresent -EnableErrorReporting:$EnableErrorReporting.Value -AccessToken $AccessToken
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
                    Write-Verbose "Failed to fetch banner content from server. Reason: $_"
                    $cmdletLogger.LogGenericError($connectionContextID, $_);
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

            if ($SkipLoadingFormatData -eq $true -and $UseRPSSession -eq $true)
            {
                Write-Warning "In RPS Session SkipLoadingFormatData switch will be ignored. Use SkipLoadingFormatData Switch only in Non RPS session."
            }

            if ($SkipLoadingCmdletHelp -eq $true -and $UseRPSSession -eq $true)
            {
                Write-Warning "In RPS Session SkipLoadingCmdletHelp switch will be ignored. Use SkipLoadingCmdletHelp Switch only in Non RPS session."
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
                if ($SigningCertificate -ne $null)
                {
                    $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext -SkipLoadingFormatData:$SkipLoadingFormatData -SigningCertificate $SigningCertificate;
                }
                else
                {
                    $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext -SkipLoadingFormatData:$SkipLoadingFormatData;
                }
                if ($null -ne $ImportedModule)
                {
                    $ImportedModuleName = $ImportedModule.Name;
                    $LogModuleDirectoryPath = $ImportedModule.ModuleBase

                    Write-Verbose "AutoGen EXOModule created at  $($ImportedModule.ModuleBase)"

                    if ($SkipLoadingCmdletHelp -eq $true)
                    {
                        $cmdletLogger.LogGenericInfo($connectionContextID, "Skipping cmdlet help data");
                    }
                    elseif ($null -ne $HelpFileNames -and $HelpFileNames -is [array] -and $HelpFileNames.Count -gt 0)
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
                    $global:EXO_LastExecutionStatus = $false;
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

            $global:EXO_LastExecutionStatus = $false;

            if ($_.Exception -ne $null)
            {
                # Connect-ExchangeOnline Failed, Remove ConnectionContext from Map.
                if ([Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveConnectionContextUsingConnectionId($connectionContextID))
                {
                    Write-Verbose "ConnectionContext Removed"
                }

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
        [string[]] $FormatTypeName = @("*"),

        # Use Remote PowerShell Session based connection
        [switch] $UseRPSSession = $false,

        # Show Banner of scc cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true
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
                $certThumbprint = $CertificateThumbprint.Value
                # Will pass CertificateThumbprint if it is not null or not empty
                if($certThumbprint)
                {
                    Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $certThumbprint -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$UseRPSSession
                }
                else
                {
                    Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$UseRPSSession
                }
            }
            else
            {
                Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$UseRPSSession
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
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High', DefaultParameterSetName='DefaultParameterSet')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='ConnectionId', ValueFromPipelineByPropertyName=$true)]
        [string[]] $ConnectionId,
        [Parameter(Mandatory=$true, ParameterSetName='ModulePrefix')]
        [string[]] $ModulePrefix
    )

    process
    {
        $global:EXO_LastExecutionStatus = $true;
        $disconnectConfirmationMessage = ""
        Switch ($PSCmdlet.ParameterSetName)
        {
            'ConnectionId'
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageByConnectionId($ConnectionId)
                break
            }
            'ModulePrefix'
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageByModulePrefix($ModulePrefix)
                break
            }
            Default
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageWithInbuilt()
            }
        }
	
        if ($PSCmdlet.ShouldProcess(
            $disconnectConfirmationMessage,
            "Press(Y/y/A/a) if you want to continue.",
            $disconnectConfirmationMessage))
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
                        # If the connectionId of the session being cleared is equal to AppSettings.ConnectionId, this means connection to EXO cmdlets will break.
                        if ($session.ConnectionContext.ConnectionId -ieq [Microsoft.Exchange.Management.AdminApiProvider.AppSettings]::ConnectionId)
                        {
                            Clear-ActiveToken -TokenProvider $session.TokenProvider -IsSessionUsedByInbuiltCmdlets:$true
                        }
                        else
                        {
                            Clear-ActiveToken -TokenProvider $session.TokenProvider -IsSessionUsedByInbuiltCmdlets:$false
                        }

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

                $modulesToRemove = $null
                Switch ($PSCmdlet.ParameterSetName)
                {
                    'ConnectionId'
                    {
                        # Call GetModulesToRemoveByConnectionId in this scenario
                        $modulesToRemove = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetModulesToRemoveByConnectionId($ConnectionId)
                        break
                    }
                    'ModulePrefix'
                    {
                        # Call GetModulesToRemoveByModulePrefix in this scenario
                        $modulesToRemove = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetModulesToRemoveByModulePrefix($ModulePrefix)
                        break
                    }
                    Default
                    {
                        # Remove all the AutoREST modules from this instance of powershell if created
                        $existingAutoRESTModules = Get-Module "tmpEXO_*"
                        foreach ($module in $existingAutoRESTModules)
                        {  
                            $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                        }

                        # The below call to remove all connection contexts could be removed as we already have an OnRemove event hooked to the module. Work Item 3461604 to investigate this
                        # Remove all ConnectionContexts
                        # this internally clears all the active tokens in ConnectionContexts
                        [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveAllConnectionContexts()
                    }
                }

                if ($modulesToRemove -ne $null -and $modulesToRemove.Count -gt 0)
                {
                    $null = Remove-Module $modulesToRemove -ErrorAction SilentlyContinue
                }

                Write-Information "Disconnected successfully !"

                # Remove all the Wrapped modules from this instance of powershell if created
                $existingWrappedModules = Get-Module "EXOCmdletWrapper-*"
                foreach ($module in $existingWrappedModules)
                {
                    $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                }

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
                $global:EXO_LastExecutionStatus = $false;

                $endTime = Get-Date
                foreach ($context in $connectionContexts)
                {
                    $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
                    for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
                    {
                        $context.Logger.LogGenericError($disconnectCmdletId, $global:Error[$i]);
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
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvIhSIfXkVGj5V
# lNDLPVcwUEwXzD8ADCBpladzOocCvaCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIOYwaG/xNn4W559cudpJEXyS
# CPBi5V2s1JODcM3glFQ0MEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAhgCqjFdv2Iyy/L8NLYsBai7zseLeFDjkG2q0FwYs3AenWZts2Y0aXqfM
# H3Mn8UZDd+4ZFPpC9peOgFsXQ2AwK4kbJeVMHj2R8pO45ee01xbTUB/Gr3X/szh1
# IkNF7yrq5UM6rhwq6LMcJdRMv5UChyC4dn2m1rcOrwGHtWbvW2WkusoRAVJhipPq
# z4HQ1KJNyOVKwfmk9KBbV3wmGDAnquy3TJtDZyju1j+rszDcLsKXrsoQaoy3NBfy
# 5sf3j7HO3YMMZon4llCa6q+AC7RLE3gfOCR+tL+93hRvs8NETaV0zWLeAnNAAe5i
# lf1kfuR7YYzzDsWeSz8UovESkIzMNqGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCCMOuHT2zoY18jSzNzVw182oan2Jon7+RIjIIiOCZ3LzQIGZNTJNVW/
# GBMyMDIzMDkwMjIzMzExOC42MThaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OEQwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAc1VByrnysGZHQABAAABzTANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# MDVaFw0yNDAyMDExOTEyMDVaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OEQwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDTOCLVS2jmEWOqxzygW7s6YLmm29pjvA+Ch6VL7HlT
# L8yUt3Z0KIzTa2O/Hvr/aJza1qEVklq7NPiOrpBAIz657LVxwEc4BxJiv6B68a8D
# QiF6WAFFNaK3WHi7TfxRnqLohgNz7vZPylZQX795r8MQvX56uwjj/R4hXnR7Na4L
# lu4mWsml/wp6VJqCuxZnu9jX4qaUxngcrfFT7+zvlXClwLah2n0eGKna1dOjOgyK
# 00jYq5vtzr5NZ+qVxqaw9DmEsj9vfqYkfQZry2JO5wmgXX79Ox7PLMUfqT4+8w5J
# kdSMoX32b1D6cDKWRUv5qjiYh4o/a9ehE/KAkUWlSPbbDR/aGnPJLAGPy2qA97YC
# BeeIJjRKURgdPlhE5O46kOju8nYJnIvxbuC2Qp2jxwc6rD9M6Pvc8sZIcQ10YKZV
# YKs94YPSlkhwXwttbRY+jZnQiDm2ZFjH8SPe1I6ERcfeYX1zCYjEzdwWcm+fFZml
# JA9HQW7ZJAmOECONtfK28EREEE5yzq+T3QMVPhiEfEhgcYsh0DeoWiYGsDiKEuS+
# FElMMyT456+U2ZRa2hbRQ97QcbvaAd6OVQLp3TQqNEu0es5Zq0wg2CADf+QKQR/Y
# 6+fGgk9qJNJW3Mu771KthuPlNfKss0B1zh0xa1yN4qC3zoE9Uq6T8r7G3/OtSFms
# 4wIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFKGT+aY2aZrBAJVIZh5kicokfNWaMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBSqG3ppKIU+i/EMwwtotoxnKfw0SX/3T16
# EPbjwsAImWOZ5nLAbatopl8zFY841gb5eiL1j81h4DiEiXt+BJgHIA2LIhKhSscd
# 79oMbr631DiEqf9X5LZR3V3KIYstU3K7f5Dk7tbobuHu+6fYM/gOx44sgRU7YQ+Y
# TYHvv8k4mMnuiahJRlU/F2vavcHU5uhXi078K4nSRAPnWyX7gVi6iVMBBUF4823o
# PFznEcHup7VNGRtGe1xvnlMd1CuyxctM8d/oqyTsxwlJAM5F/lDxnEWoSzAkad1n
# WvkaAeMV7+39IpXhuf9G3xbffKiyBnj3cQeiA4SxSwCdnx00RBlXS6r9tGDa/o9R
# S01FOABzKkP5CBDpm4wpKdIU74KtBH2sE5QYYn7liYWZr2f/U+ghTmdOEOPkXEcX
# 81H4dRJU28Tj/gUZdwL81xah8Kn+cB7vM/Hs3/J8tF13ZPP+8NtX3vu4NrchHDJY
# gjOi+1JuSf+4jpF/pEEPXp9AusizmSmkBK4iVT7NwVtRnS1ts8qAGHGPg2HPa4b2
# u9meueUoqNVtMhbumI1y+d9ZkThNXBXz2aItT2C99DM3T3qYqAUmvKUryVSpMLVp
# se4je5WN6VVlCDFKWFRH202YxEVWsZ5baN9CaqCbCS0Ea7s9OFLaEM5fNn9m5s69
# lD/ekcW2qTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
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
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjhEMDAtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBo
# qfem2KKzuRZjISYifGolVOdyBKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6J5CDDAiGA8yMDIzMDkwMjIzMjE0
# OFoYDzIwMjMwOTAzMjMyMTQ4WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDonkIM
# AgEAMAcCAQACAjHhMAcCAQACAhIAMAoCBQDon5OMAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAG9zvQQ+oME1XyTEL3IPb0PWstui14Y0x63D6iM6TmKnQR52
# 5WX9Bg9Fi7w9u3PpB+tdiME2LQc6FpVnKaIPZ0dTEcyO8d4Bqpmyj941HCPlu2DM
# lGwLpxJ0UJror36/udg0N6F6I8ZIpHC6yTHDx8Z/AXykQgXO0nL8iVhR0MwtmyYy
# kdDSGSHhKmPIw9DrczkjIOzT/9Uy+6uWZYUEK2FHEpT9g/EVzHDFjrDVYJ786Sr2
# CbQqa39mNGrwPk3/8M6O4VfIAGWamzCXQdgfnh/ZZ3mQYDSioS5UdrBsGMiETVri
# VG0fK5Xd9phzZwB/bU7VH7Mnq9PBoplv9a29JUkxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAc1VByrnysGZHQABAAABzTAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCD0L/6262giWefIMpEbOeWovvYNExwqJ3nYjm0T0mPyRTCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIOJmpfitVr1PZGgvTEdTpStUc6GN
# h7LNroQBKwpURpkKMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHNVQcq58rBmR0AAQAAAc0wIgQg67ajXSgZdhKqmoyWcUNtkElWAK2D
# 9ecV1wZkKvELLm8wDQYJKoZIhvcNAQELBQAEggIAxT3LRIaYtCUYi5GebLgIX8Rd
# KJ/LQeWQY/koV9pObD4yfjuF0S/Mq+nJV+X6hRyCq6QODEguTwYlAwcVvVuefuyJ
# lNH8WSuE/ztXE8J5rCo/yhNDHDec1Vywsg+31dcEUkF/aW/q2JxOPzS8Vl6XOvae
# PjVhickQvLFO5UTd6tBipDaR0j0Yh6i0OrHzcDS5IH+z4PEZmuLP98+foDqpsWIM
# c0cazL/BXxLBFUqiQyG1+MrUEQuQKCRrgu1Kzebkj4S6Q9QsCm37ED8QwCbCO2Fo
# EGA6eQ/zjyZT6WHkisEoTGBzMPo+OpT/G96QALucXd0NMt/WEliA1pwz+jCu2VQh
# N376QFAjAaAD8AW22v1Zp9qrKeMHntXfnFS+cegGDY/diEqLJugcY2EfPzTe910n
# Qu2QxsRAdXugYlBcm9HaAMITVJj9mRQUzwTN+NmHl84MaMnkCM1ojOOSyPqd1Ui/
# 74OynoFdHzFNVCXpFSKnoXTkrn4zFaCnpjd7imu4sXTz+6sW1zL4pNPSbQZAGmMN
# mz9rpYed8kWgl3i8WoDdHG3JL637ngvkC3MDoQvFTshT9kXtMolb4UeqhSm+ywXv
# jhc/6pusnGo0DdpP96PNpZwZPMdMFkJcAGP3wbbRV5H84o7Cy8l/OoDjYMOLuHQj
# Ol0ZMRkruTaf2/8P/qM=
# SIG # End signature block
