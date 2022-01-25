<#
.SYNOPSIS
Export-RecipientPermissions XXXVersionStringXXX
Document mailbox access rights and folder permissions, "send as", "send on behalf" and "managed by"

.DESCRIPTION
Signatures and OOF messages can be:
- Generated from templates in DOCX or HTML file format
- Customized with a broad range of variables, including photos, from Active Directory and other sources
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific e-mail addresses, for every primary mailbox across all Outlook profiles
- Assigned time ranges within which they are valid
- Set as default signature for new e-mails, or for replies and forwards (signatures only)
- Set as default OOF message for internal or external recipients (OOF messages only)
- Set in Outlook Web for the currently logged in user
- Centrally managed only or exist along user created signatures (signatures only)
- Copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only)

Export-RecipientPermissions can be executed by users on clients, or on a server without end user interaction.
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

Sample templates for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

Simulation mode allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.

The script is designed to work in big and complex environments (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works on premises, in hybrid and cloud-only environments.

It is multi-client capable by using different template paths, configuration files and script parameters.

Set-OutlookSignature requires no installation on servers or clients. You only need a standard file share on a server, and PowerShell and Office on the client.

A documented implementation approach, based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators.
The implementatin approach is suited for service providers as well as for clients, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The script is Free and Open-Source Software (FOSS). It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see '.\docs\LICENSE.txt' for copyright and MIT license details.

.LINK
Github: https://github.com/GruberMarkus/Export-RecipientPermissions

.PARAMETER SignatureTemplatePath
Path to centrally managed signature templates.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
Default value: '.\templates\Signatures DOCX'

.PARAMETER SignatureIniPath
Path to ini file containing signature template tags
This is an alternative to file name tags
See '.\templates\sample signatures ini file.ini' for a sample file with further explanations.
Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
The currently logged in user needs at least read access to the path
Default value: ''

.PARAMETER ReplacementVariableConfigFile
Path to a replacement variable config file.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
Default value: '.\config\default replacement variables.txt'

.PARAMETER GraphConfigFile
Path to a Graph variable config file.
Local and remote paths are supported
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signature')
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/config/default graph config.ps1' or '\\server.domain@SSL\SignatureSite\config\default graph config.ps1'
The currently logged in user needs at least read access to the path
Default value: '.\config\default graph config.ps1'

.PARAMETER TrustsToCheckForGroups
List of trusted domains to check for group membership across trusts.
If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.
If a string starts with a minus or dash ("-domain-a.local"), the domain after the dash or minus is removed from the list.
Subdomains of trusted domains are always considered.
Default value: '*'

.PARAMETER DeleteUserCreatedSignatures
Shall the script delete signatures which were created by the user itself?
Default value: $false

.PARAMETER DeleteScriptCreatedSignaturesWithoutTemplate
Shall the script delete signatures which were created by the script before but are no longer available as template?
default value: $true

.PARAMETER SetCurrentUserOutlookWebSignature
Shall the script set the Outlook Web signature of the currently logged in user?
If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used.
Default value: $true

.PARAMETER SetCurrentUserOOFMessage
Shall the script set the Out of Office (OOF) auto reply message of the currently logged in user?
If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used.
Default value: $true

.PARAMETER OOFTemplatePath
Path to centrally managed signature templates.
Local and remote paths are supported.
Local paths can be absolute ('C:\OOF templates') or relative to the script path ('.\templates\Out of Office').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/OOFTemplates' or '\\server.domain@SSL\SignatureSite\OOFTemplates'
The currently logged in user needs at least read access to the path.
Default value: '.\templates\Out of Office DOCX'

.PARAMETER OOFIniPath
Path to ini file containing signature template tags
This is an alternative to file name tags
See '.\templates\sample OOF ini file.ini' for a sample file with further explanations.
Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
The currently logged in user needs at least read access to the path
Default value: ''

.PARAMETER AdditionalSignaturePath
An additional path that the signatures shall be copied to.
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.
This way, the user can easily copy-paste the preferred preconfigured signature for use in an e-mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.
Local and remote paths are supported.
Local paths can be absolute ('C:\Outlook signatures') or relative to the script path ('.\Outlook signatures').
WebDAV paths are supported (https only): 'https://server.domain/User' or '\\server.domain@SSL\User'
The currently logged in user needs at least write access to the path.
If the folder or folder structure does not exist, it is created.
Default value: "$([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures'))"

.PARAMETER AdditionalSignaturePathFolder
An optional folder or folder structure below AdditionalSignaturePath.
This parameter is available for compatibility with versions before 2.2.1. Starting with 2.2.1, you can pass a full path via the parameter AdditionalSignaturePath, so AdditionalSignaturePathFolder is no longer needed.
If the folder or folder structure does not exist, it is created.
Default value: ''

.PARAMETER UseHtmTemplates
With this parameter, the script searches for templates with the extension .htm instead of .docx.
Each format has advantages and disadvantages, please see "Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates." for a quick overview.
Default value: $false

.PARAMETER SimulateUser
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged in user.
Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

.PARAMETER SimulateMailboxes
SimulateMailboxes is optional for simulation mode, although highly recommended.
It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.


.PARAMETER GraphCredentialFile
Path to file containing Graph credential which should be used as alternative to other token acquisition methods
Makes only sense in combination with '.\sample code\SimulateAndDeploy.ps1', do not use this parameter for other scenarios
See '.\sample code\SimulateAndDeploy.ps1' for an example how to create this file
Default value: $null

.PARAMETER GraphOnly
Try to connect to Microsoft Graph only, ignoring any local Active Directory.
The default behavior is to try Active Directory first and fall back to Graph.
Default value: $false

.PARAMETER CreateRTFSignatures
Should signatures be created in RTF format?
Default value: $true

.PARAMETER CreateTXTSignatures
Should signatures be created in TXT format?
Default value: $true

.PARAMETER EmbedImagesInHTML
Should images be embedded into HTML files?
Outlook 2016 and newer can handle images embedded directly into an HTML file as BASE64 string ('<img src="data:image/[...]"').
Outlook 2013 and earlier can't handle these embedded images when composing HTML e-mails (there is no problem receiving such e-mails, or when composing RTF or TXT e-mails).
When setting EmbedimagesInHTML to $false, consider setting the Outlook registry value "Send Pictures With Document" to 1 to ensure that images are sent to the recipient (see https://support.microsoft.com/en-us/topic/inline-images-may-display-as-a-red-x-in-outlook-704ae8b5-b9b6-d784-2bdf-ffd96050dfd6 for details).
Default value: $true

.INPUTS
None. You cannot pipe objects to Export-RecipientPermissions.ps1.

.OUTPUTS
Export-RecipientPermissions.ps1 writes the current activities, warnings and error messages to the standard output stream.

.EXAMPLE
Run Export-RecipientPermissions with default values and sample templates
PS> .\Export-RecipientPermissions.ps1

.EXAMPLE
Use custom signature templates
PS> .\Export-RecipientPermissions.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates'

.EXAMPLE
Use custom signature templates, ignore trust to internal-test.example.com
PS> .\Export-RecipientPermissions.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -TrustsToCheckForGroups '*', '-internal-test.example.com'

.EXAMPLE
Use custom signature templates, only check domains/trusts internal-test.example.com and company.b.com
PS> .\Export-RecipientPermissions.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -TrustsToCheckForGroups 'internal-test.example.com', 'company.b.com'

.EXAMPLE
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. See '.\docs\README.html' for details.
PowerShell.exe -Command "& '\\server\share\directory\Export-RecipientPermissions.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1' "

.EXAMPLE
Please see '.\docs\README.html' and https://github.com/GruberMarkus/Export-RecipientPermissions for more details.

.NOTES
Script : Export-RecipientPermissions
Version: XXXVersionStringXXX
Web    : https://github.com/GruberMarkus/Export-RecipientPermissions
License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)
#>


[CmdletBinding(PositionalBinding = $false)]


Param(
    # Environments to consider: Office 365 (Exchange Online) and/or Exchange on premises
    [boolean]$ExportFromOnPrem = $true, # Highly recommended to enable this for fast initial recipient enumeration
    [boolean]$ExportFromCloud = $false,

    # Permission types to export
    [boolean]$ExportAccessRights = $true, # Rights like "FullAccess" and "ReadAccess" to the entire mailbox
    [boolean]$ExportFullAccessPerTrustee = $true, # Additionally export a list showing who has full access to which mailbox
    [boolean]$ExportSendAs = $true, # Send as
    [boolean]$ExportSendOnBehalf = $true, # Send on behalf
    [boolean]$ExportManagedBy = $true, # Only valid for groups
    [boolean]$ExportFolderPermissions = $false, # Export permissions set on specific mailbox folders. This will take very long.
    [boolean]$ResolveGroups = $false, # Resolve trustee groups to individual members (recursively)

    # Name of the permission export file
    [string]$ExportFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Output.csv",

    # Name of the error file
    [string]$ErrorFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Errors.txt",

    # Name of the transcript file
    [string]$TranscriptFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Transcript.txt",

    # Name of temporary recipient file
    [string]$TempRecipientFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Recipients.csv",

    # Folder to additionally store files created when $ExportFullAccessPerTrustee = $true. This folder must already exist at runtime. Set to "" when not needed.
    [string]$TargetFolder = "\\server.domain\share\folder",

    # Parallelization
    # Watch RAM and CPU usage
    [int]$NumberOfJobsParallel = 30, # Each job is a separate session towards Exchange on-prem and Office 365, so watch your maximum concurreny settings
    [int]$RecipientsPerJob = 100, # More recipients save time as jobs run longer, but the risk of a problem with the O365 connection is higher

    # User name and password are stored in secure string format
    [string]$CredentialPasswordFile = ".\Export-RecipientPermissions_CredentialPassword.txt",
    [string]$CredentialUsernameFile = ".\Export-RecipientPermissions_CredentialUsername.txt"
)


#
# Do not change anything from here on.
#

Function Pause($M = "Press any key to continue . . . ") { If ($psISE) { $S = New-Object -ComObject "WScript.Shell"; $B = $S.Popup("Click OK to continue.", 0, "Script Paused", 0); Return }; Write-Host -NoNewline $M; $I = 16, 17, 18, 20, 91, 92, 93, 144, 145, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183; While ($Null -Eq $K.VirtualKeyCode -Or $I -Contains $K.VirtualKeyCode) { $K = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") }; Write-Host }

$script:SessionCloud = $null

function Connect-ExchangeOnPrem {
    $Stoploop = $false
    [int]$Retrycount = 0
    do {
        try {
            #Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
            #. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
            #Connect-ExchangeServer -auto -ErrorAction Stop
            ###
            $env:tmp = 'c:\alexclude\PowerShell.temp'
            Get-ChildItem $env:tmp -Directory | Where-Object { $_.LastWriteTime -le (Get-Date).adddays(-2) } | Remove-Item -Force -Recurse
            Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://W02-EX10.sv-services.at/PowerShell/ -Authentication Kerberos) -DisableNameChecking
            Set-AdServerSettings -ViewEntireForest $True
            ###
            Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop | Out-Null
            $Stoploop = $true
        } catch {
            if ($Retrycount -le 3) {
                Write-Host "Could not connect to Exchange on-prem. Trying again in 70 seconds."
                Start-Sleep -Seconds 70
                $Retrycount = $Retrycount + 1
            } else {
                Write-Host "Could not connect to Exchange on-prem after three retires. Exiting."
                exit
                $Stoploop = $true
            }
        }
    } While ($Stoploop -eq $false)
}

function Connect-ExchangeOnline {
    $Stoploop = $false
    [int]$Retrycount = 0
    do {
        try {
            $test = $null
            $test = (Get-PSSession | Where-Object { ($_.name -like "O365Session") -and ($_.state -like "opened") })
            if ($null -eq $test) {
                $CloudUser = Get-Content $CredentialUsernameFile
                $CloudPassword = Get-Content $CredentialPasswordFile | ConvertTo-SecureString
                $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CloudUser, $CloudPassword
                $script:SessionCloud = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -Name "O365Session" -ErrorAction Stop
                (Invoke-Command -Session $script:SessionCloud -ScriptBlock { Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop }) | Out-Null
            }
            $Stoploop = $true
        } catch {
            if ($Retrycount -le 3) {
                Write-Host "Could not connect to Exchange Online. Trying again in 70 seconds."
                Start-Sleep -Seconds 70
                $Retrycount = $Retrycount + 1
            } else {
                Write-Host "Could not connect to Exchange Online after three retires. Exiting."
                exit
                $Stoploop = $true
            }
        }
    } While ($Stoploop -eq $false)
}

function write-HostColored() {
    <#
    .SYNOPSIS
    A wrapper around write-Host that supports selective coloring of
    substrings.

    .DESCRIPTION
    In addition to accepting a default foreground and background color,
    you can embed one or more color specifications in the string to write,
    using the following syntax:
    #<fgcolor>[:<bgcolor>]#<text>#

    <fgcolor> and <bgcolor> must be valid [ConsoleColor] values, such as 'green' or 'white' (case does not matter).
    Everything following the color specification up to the next '#' or, impliclitly, the end of the string
    is written in that color.

    Note that nesting of color specifications is not supported.
    As a corollary, any token that immediately follows a color specification is treated
    as text to write, even if it happens to be a technically valid color spec too.
    This allows you to use, e.g., 'The next word is #green#green#.', without fear
    of having the second '#green' be interpreted as a color specification as well.

    .PARAMETER ForegroundColor
    Specifies the default text color for all text portions
    for which no embedded foreground color is specified.

    .PARAMETER BackgroundColor
    Specifies the default background color for all text portions
    for which no embedded background color is specified.

    .PARAMETER NoNewline
    Output the specified string withpout a trailing newline.

    .NOTES
    While this function is convenient, it will be slow with many embedded colors, because,
    behind the scenes, write-Host must be called for every colored span.

    .EXAMPLE
    write-HostColored "#green#Green foreground.# Default colors. #blue:white#Blue on white."

    .EXAMPLE
    '#black#Black on white (by default).#Blue# Blue on white.' | Write-HostColored -BackgroundColor White

    #>
    [CmdletBinding(ConfirmImpact = 'None', SupportsShouldProcess = $false, SupportsTransactions = $false)]
    param(
        [parameter(Position = 0, ValueFromPipeline = $true)]
        [string[]] $Text
        ,
        [switch] $NoNewline
        ,
        [ConsoleColor] $BackgroundColor = $host.UI.RawUI.BackgroundColor
        ,
        [ConsoleColor] $ForegroundColor = $host.UI.RawUI.ForegroundColor
    )

    begin {
        # If text was given as an operand, it'll be an array.
        # Like write-Host, we flatten the array into a single string
        # using simple string interpolation (which defaults to separating elements with a space,
        # which can be changed by setting $OFS).
        if ($null -ne $Text) {
            $Text = "$Text"
        }
    }

    process {
        if ($Text) {

            # Start with the foreground and background color specified via
            # -ForegroundColor / -BackgroundColor, or the current defaults.
            $curFgColor = $ForegroundColor
            $curBgColor = $BackgroundColor

            # Split message into tokens by '#'.
            # A token between to '#' instances is either the name of a color or text to write (in the color set by the previous token).
            $tokens = $Text.split("#")

            # Iterate over tokens.
            $prevWasColorSpec = $false
            foreach ($token in $tokens) {

                if (-not $prevWasColorSpec -and $token -match '^([a-z]+)(:([a-z]+))?$') {
                    # a potential color spec.
                    # If a token is a color spec, set the color for the next token to write.
                    # Color spec can be a foreground color only (e.g., 'green'), or a foreground-background color pair (e.g., 'green:white')
                    try {
                        $curFgColor = [ConsoleColor]  $matches[1]
                        $prevWasColorSpec = $true
                    } catch {}
                    if ($matches[3]) {
                        try {
                            $curBgColor = [ConsoleColor]  $matches[3]
                            $prevWasColorSpec = $true
                        } catch {}
                    }
                    if ($prevWasColorSpec) {
                        continue
                    }
                }

                $prevWasColorSpec = $false

                if ($token) {
                    # A text token: write with (with no trailing line break).
                    # !! In the ISE - as opposed to a regular PowerShell console window,
                    # !! $host.UI.RawUI.ForegroundColor and $host.UI.RawUI.ForegroundColor inexcplicably
                    # !! report value -1, which causes an error when passed to write-Host.
                    # !! Thus, we only specify the -ForegroundColor and -BackgroundColor parameters
                    # !! for values other than -1.
                    $argsHash = @{}
                    if ([int] $curFgColor -ne -1) { $argsHash += @{ 'ForegroundColor' = $curFgColor } }
                    if ([int] $curBgColor -ne -1) { $argsHash += @{ 'BackgroundColor' = $curBgColor } }
                    Write-Host -NoNewline @argsHash $token
                }

                # Revert to default colors.
                $curFgColor = $ForegroundColor
                $curBgColor = $BackgroundColor

            }
        }
        # Terminate with a newline, unless suppressed
        if (-not $NoNewLine) { Write-Host }
    }
}

$error.clear()

Set-Location $PSScriptRoot

Clear-Host

$TargetFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TargetFolder)
$Exportfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Exportfile)
$ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
$TranscriptFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TranscriptFile)
$TempRecipientFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TempRecipientFile)
$CredentialPasswordFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CredentialPasswordFile)
$CredentialUsernameFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CredentialUsernameFile)
New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $TempRecipientFile) | Out-Null
if (Test-Path $Exportfile) { (Remove-Item $ExportFile -Force) }
if (Test-Path $Errorfile) { (Remove-Item $ErrorFile -Force) }
if (Test-Path $TranscriptFile) { (Remove-Item $TranscriptFile -Force) }
if (Test-Path $TempRecipientFile) { (Remove-Item $TempRecipientFile -Force) }
if (($ExportFullAccessPerTrustee -eq $true) -and ($ExportAccessRights -eq $true)) {
    if (Test-Path ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee")) {
        Remove-Item ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee") -Force -Recurse
    }
    New-Item -ItemType Directory -Force -Path ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee") | Out-Null
}

Start-Transcript -Path ($TranscriptFile + "_temp") -Force

if ($ExportFromCloud -eq $true) {
    if ((Test-Path $CredentialUsernameFile) -and (Test-Path $CredentialPasswordFile)) { } else {
        Write-Host 'Please enter cloud user name for later use.'
        Read-Host | Out-File $CredentialUsernameFile
        Write-Host 'Please enter cloud admin password for later use.'
        Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File $CredentialPasswordFile
    }
}

# Test on-prem connection
if (($ExportFromOnPrem -eq $true)) {
    Try {
        Connect-ExchangeOnPrem
        Write-Host 'On-prem connection working.'
    } Catch {
        Write-Host 'On-prem connection does not work. Error executing ''Get-Recipient -ResultSize 1''. Exiting.'
        Write-Host 'Please start the script on an Exchange server with appropriate permissions.'
        $ExportFromOnPrem = $false
        exit
    }
}

if (($ExportFromCloud -eq $true)) {
    Try {
        Write-Host "Connecting to Exchange Online."
        Connect-ExchangeOnline
        Write-Host 'Cloud connection working.'
    } Catch {
        Write-Host 'Cloud connection does not work. Error executing ''Get-Recipient -ResultSize 1''. Exiting.'
        $ExportFromCloud = $false
        exit
    }
}

# Export list of objects
if ($ExportFromOnPrem -eq $true) {
    Write-Host 'Enumerating on-prem recipients. This may take a long time.'
} else {
    if ($ExportFromCloud -eq $true) {
        Write-Host 'Enumerating cloud recipients. This may take a long time.'
    } else {
        Write-Host 'Neither on-prem nor cloud connection configured or possible. Exiting.'
        exit
    }
}


if ($ExportFromOnPrem -eq $true) {
    get-recipient -recipienttype MailUniversalSecurityGroup, DynamicDistributionGroup, UserMailbox, MailUniversalDistributionGroup, MailUniversalSecurityGroup, MailNonUniversalGroup, MailUser -resultsize 200 -WarningAction silentlyContinue | Select-Object DistinguishedName | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    <#
    get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype PublicFolder -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype DynamicDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype UserMailbox -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailUniversalDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailNonUniversalGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailUser -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailContact -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    #>
} else {
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock { get-recipient -recipienttype MailUniversalSecurityGroup, DynamicDistributionGroup, UserMailbox, MailUniversalDistributionGroup, MailUniversalSecurityGroup, MailNonUniversalGroup, MailUser -resultsize unlimited -WarningAction silentlyContinue | Select-Object DistinguishedName }) | Select-Object DistinguishedName | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    <#
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype PublicFolder -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype DynamicDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype UserMailbox -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailNonUniversalGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUser -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailContact -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    #>
}


if (($ExportFromCloud -eq $true)) {
    Write-Host 'Disconnecting from cloud services.'
    Remove-PSSession $script:SessionCloud
    #if ((test-path (Split-Path -Path $script:SessionCloudPath.path)) -eq $true) {
    #    Remove-Item (Split-Path -Path $script:SessionCloudPath.path) -Recurse -Force
    #}
}

# Import list of objects
$Recipients = (Import-Csv $TempRecipientFile)
$RecipientCount = ($Recipients | Measure-Object).count
$count = 1



$Batch = 0
for ($RecipientStartID = 0; $RecipientStartID -lt $RecipientCount; $RecipientStartID += $RecipientsPerJob) {
    $RecipientEndID = ($RecipientStartID + $RecipientsPerJob - 1)
    $Batch++
}

Write-Host "$RecipientCount recipients found. Reading permissions in $Batch batches of $RecipientsPerJob recipients each."
Write-Host "Up to $NumberOfJobsParallel of $Batch batches will run in parallel. Output is updated at completion of a single batch."

Get-Job | Remove-Job -Force
$Batch = 1
for ($RecipientStartID = 0; $RecipientStartID -lt $RecipientCount; $RecipientStartID += $RecipientsPerJob) {
    $RecipientEndID = ($RecipientStartID + $RecipientsPerJob - 1)
    $running = @(Get-Job -State running)
    foreach ($x in (Get-Job -State Completed)) {
        if (Test-Path ($Exportfile + '_temp' + $x.name)) { (Get-Content ($Exportfile + '_temp' + $x.name)) | Write-HostColored }
    }
    if ($running.Count -ge $NumberOfJobsParallel) {
        # wait and receive
        while ($true) {
            if (@(Get-Job -State running).count -lt $NumberOfJobsParallel) {
                foreach ($x in (Get-Job -State Completed)) {
                    $temp = $null
                    $TempPath = $null
                    # show temp job output file, delete output file
                    $TempPath = ($Exportfile + '_temp' + $x.Name)
                    if (Test-Path $TempPath) {
                        $temp = Get-Content $TempPath
                        $temp | Write-HostColored
                        Remove-Item $TempPath -Force
                    }
                    $temp = $null
                    $TempPath = $null

                    # append temp error file and delete temp file
                    $TempPath = ($Errorfile + '_temp' + $x.Name)
                    $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
                    if (Test-Path $TempPath) {
                        $temp = Get-Content $TempPath
                        $temp | Out-File $Errorfile -Append -Force
                        Remove-Item $TempPath -Force
                    }
                    $temp = $null
                    $TempPath = $null

                    # append temp transcript file and delete temp file
                    $TempPath = ($Transcriptfile + '_temp' + $x.Name)
                    $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
                    if (Test-Path $TempPath) {
                        $temp = Get-Content $TempPath
                        $temp | Out-File $Transcriptfile -Append -Force
                        Remove-Item $TempPath -Force
                    }

                    # append temp export file and delete temp file
                    $TempPath = ($Exportfile + '_temp' + $x.Name)
                    $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
                    if (Test-Path $TempPath) {
                        $temp = Import-Csv $TempPath -Delimiter ";"
                        $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                        Remove-Item $TempPath -Force
                    }
                    $temp = $null
                    $TempPath = $null
                    Remove-Job -Job $x -Force
                }
                break
            } else {
                [System.GC]::Collect() # garbage collection
                Start-Sleep -s 5
            }
        }
    }
    Start-Job {
        param(
            $RecipientStartID,
            $RecipientEndID,
            $Exportfile,
            $ErrorFile,
            $TempRecipientFile,
            $ExportFromOnPrem,
            $ExportFromCloud,
            $CredentialPasswordFile,
            $CredentialUsernameFile,
            $ExportAccessRights,
            $ExportSendAs,
            $ExportSendOnBehalf,
            $ExportManagedBy,
            $ExportFolderPermissions,
            $ExportFullAccessPerTrustee,
            $TranscriptFile,
            $ResolveGroups
        )
        Start-Sleep -s (Get-Random -Minimum 0 -Maximum 20)
        Set-Location $PSScriptRoot
        $Exportfile = $Exportfile + '_temp' + $RecipientStartID
        $ErrorFile = $ErrorFile + '_temp' + $RecipientStartID
        $TranscriptFile = $TranscriptFile + '_temp' + $RecipientStartID
        $Exportfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Exportfile)
        $ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
        $TranscriptFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TranscriptFile)
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile) | Out-Null
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile) | Out-Null
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile) | Out-Null
        if (Test-Path $Exportfile) { (Remove-Item $ExportFile -Force) }
        if (Test-Path $Errorfile) { (Remove-Item $ErrorFile -Force) }
        Start-Transcript -Path $TranscriptFile -Force
        Write-Host ("RecipientStartID: " + $RecipientStartID)
        Write-Host ("RecipientEndID: " + $RecipientEndID)
        Write-Host ("Time: " + (Get-Date))

        $script:BatchSessionCloud = $null
        function Connect-ExchangeOnPrem {
            $Stoploop = $false
            [int]$Retrycount = 0
            do {
                try {
                    #Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
                    #. $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                    #Connect-ExchangeServer -auto -ErrorAction Stop
                    ###
                    $env:tmp = 'c:\alexclude\PowerShell.temp'
                    Get-ChildItem $env:tmp -Directory | Where-Object { $_.LastWriteTime -le (Get-Date).adddays(-2) } | Remove-Item -Force -Recurse
                    Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://W02-EX10.sv-services.at/PowerShell/ -Authentication Kerberos) -DisableNameChecking
                    Set-AdServerSettings -ViewEntireForest $True
                    ###
                    Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop | Out-Null
                    $Stoploop = $true
                } catch {
                    if ($Retrycount -le 3) {
                        Write-Host ("Time: " + (Get-Date))
                        Write-Host "Could not connect to Exchange on-prem. Trying again in 70 seconds."
                        Start-Sleep -Seconds 70
                        $Retrycount = $Retrycount + 1
                    } else {
                        Write-Host ("Time: " + (Get-Date))
                        Write-Host "Could not connect to Exchange on-prem after three retires. Exiting."
                        exit
                        $Stoploop = $true
                    }
                }
            } While ($Stoploop -eq $false)
        }

        function Connect-ExchangeOnline {
            $Stoploop = $false
            [int]$Retrycount = 0
            do {
                try {
                    $test = $null
                    $test = (Get-PSSession | Where-Object { ($_.name -like "O365BatchSession") -and ($_.state -like "opened") })
                    if ($null -eq $test) {
                        $CloudUser = Get-Content $CredentialUsernameFile
                        $CloudPassword = Get-Content $CredentialPasswordFile | ConvertTo-SecureString
                        $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CloudUser, $CloudPassword
                        $script:BatchSessionCloud = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -Name "O365BatchSession" -ErrorAction Stop
                    }
                    $Stoploop = $true
                } catch {
                    if ($Retrycount -le 3) {
                        Write-Host ("Time: " + (Get-Date))
                        Write-Host "Could not connect to Office 365. Trying again in 70 seconds."
                        Start-Sleep -Seconds 70
                        $Retrycount = $Retrycount + 1
                    } else {
                        Write-Host ("Time: " + (Get-Date))
                        Write-Host "Could not connect to Office 365 after three retires. Exiting."
                        exit
                        $Stoploop = $true
                    }
                }
            } While ($Stoploop -eq $false)
        }


        filter get_member_recurse {
            if ($_) {
                $tempObject = get-recipient -identity $_.tostring()
                if ($tempObject.RecipientType -ilike "*group") {
                    Get-DistributionGroupMember $tempObject.identity | get_member_recurse
                } else {
                    $tempObject
                }
            }
        }


        $Recipients = (Import-Csv $TempRecipientFile)
        $RecipientCount = $Recipients.count
        $Count = $RecipientStartID + 1
        #$BatchID = ($RecipientStartID / ($RecipientEndID - $RecipientStartID + 1)) + 1
        if (($ExportFromCloud -eq $true)) { Connect-ExchangeOnline }
        if (($ExportFromOnPrem -eq $true)) { Connect-ExchangeOnPrem }
        $ErrorCount = 0
        for ($RecipientStartID; $RecipientStartID -le $RecipientEndID; $RecipientStartID++) {
            Write-Host "Time: $(Get-Date); RecipientID: $RecipientStartID; '$($Recipients[$RecipientStartID].DistinguishedName)'"
            if ($RecipientStartID -ge $Recipients.length) { break }
            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
            if ($ExportFromOnPrem -eq $true) {
                $Recipient = get-recipient $Recipients[$RecipientStartID].DistinguishedName -resultsize 1
            } else {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $Recipient = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient $($args[0]) -resultsize 1 } -ArgumentList $Recipients[$RecipientStartID].DistinguishedName)
            }
            $GrantorDisplayName = ""
            $GrantorPrimarySMTP = ""
            $GrantorRecipientType = ""
            $GrantorRecipientTypeDetails = ""
            $GrantorLegacyExchangeDN = ""
            $GrantorOU = ""
            $ALias = ""
            $x = $null
            if ($ExportFromOnPrem -eq $true) {
                if ($Recipient.RecipientTypeDetails -like "Remote*") { $GrantorCloudOrOnPrem = 'Cloud' } else { $GrantorCloudOrOnPrem = 'On-Prem' }
                if ($Recipient.RecipientTypeDetails -like "*Group") { $GrantorCloudOrOnPrem = 'On-Prem' }
            } else {
                if ($Recipient.RecipientTypeDetails -like "Remote*") { $GrantorCloudOrOnPrem = 'On-Prem' } else { $GrantorCloudOrOnPrem = 'Cloud' }
                if ($Recipient.RecipientTypeDetails -like "*Group") { $GrantorCloudOrOnPrem = 'Cloud' }
            }
            $GrantorDisplayName = $Recipient.DisplayName.tostring()
            $GrantorPrimarySMTP = $Recipient.PrimarySMTPAddress.tostring()
            $GrantorRecipientType = $Recipient.RecipientType.tostring()
            $GrantorRecipientTypeDetails = $Recipient.RecipientTypeDetails.tostring()
            $GrantorOU = $Recipient.OrganizationalUnit.tostring()
            $Alias = $Recipient.name.tostring()
            if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $RecipientTemp = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient $($args[0]) -resultsize 10 } -ArgumentList $GrantorPrimarySMTP)
                if ($recipientTemp.count) {
                    foreach ($x in $recipientTemp) {
                        if ($x.recipienttypedetails -like "*mailbox*") {
                            $GrantorDN = $x.DistinguishedName.tostring()
                        }
                    }
                } else {
                    $GrantorDN = $recipienttemp.DistinguishedName.tostring()
                }
            } else {
                $GrantorDN = $Recipient.DistinguishedName.tostring()
            }
            $x = $null
            $recipientTemp = $null

            $Text = ("{0:000000}/{1:000000}: " -f $count, $RecipientCount) + $GrantorPrimarySMTP + ', ' + $GrantorRecipienttype + '/' + $GrantorRecipientTypeDetails + ', ' + $GrantorCloudOrOnPrem

            if (($Recipient.Recipienttype -eq "PublicFolder") -or ($Recipient.Recipienttype -eq "MailContact")) { $Text += (", recipient type $GrantorRecipientType not supported."); continue }


            # Access Rights (full access etc.)
            if ($ExportAccessRights -eq $true) {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                if (($GrantorRecipientType -NotMatch 'group')) {
                    $Text += ', AccessRights'
                    if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                        try {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            $TrusteeRightsQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-MailboxPermission -identity $($args[0]) -resultsize unlimited -wa stop -ea stop } -ArgumentList $GrantorDN) | Where-Object { ($_.IsInherited -eq $false) -and ($_.user -inotlike 'NT AUTHORITY\*') }
                            $GrantorLegacyExchangeDN = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-Mailbox -identity $($args[0]) -resultsize 1 -wa stop -ea stop | Select-Object LegacyExchangeDN } -ArgumentList $GrantorDN).LegacyExchangeDN
                        } catch {
                        }
                    } else {
                        try {
                            $TrusteeRightsQuery = (get-mailboxpermission -Identity $GrantorDN -resultsize unlimited -wa stop -ea stop | Where-Object { ($_.IsInherited -eq $false) -and ($_.user -inotlike "NT AUTHORITY\*") })
                            $GrantorLegacyExchangeDN = (Get-Mailbox -identity $GrantorDN -resultsize 1 -wa stop -ea stop).LegacyExchangeDN
                        } catch {
                        }
                    }
                    if ($ResolveGroups) {
                        $TrusteeIdentityOriginal = @($TrusteeRightsQuery.user | get_member_recurse | Select-Object @{Name = 'Trustee'; Expression = { $_.identity } })
                    } else {
                        $TrusteeIdentityOriginal = @($TrusteeRightsQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.user } })
                    }
                    if ($error.count -eq 0) {
                        foreach ($TrusteeIdentity in $TrusteeIdentityOriginal) {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            try {
                                $TrusteeIdentityQuery = (get-recipient ($TrusteeIdentity.trustee.tostring()) -resultsize 1 -wa stop -ea stop)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'Cloud' } else { $TrusteeCloudOrOnPrem = 'On-Prem' }
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'On-Prem' }
                            } catch {
                                try {
                                    if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                                    if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                                        $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $TrusteeIdentity.trustee.tostring())
                                    } else {
                                        $TrusteeIdentityQuery = get-recipient $TrusteeIdentity.trustee.tostring() -resultsize 1 -wa stop -ea stop
                                    }
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'On-Prem' } else { $TrusteeCloudOrOnPrem = 'Cloud' }
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'Cloud' }
                                } catch {
                                    continue
                                }
                            }
                            $error.clear()
                            $TrusteeRecipientType = $null
                            $TrusteeRecipientTypeDetails = $null
                            $TrusteeDisplayName = $null
                            $TrusteePrimarySMTP = $null
                            $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                            $TrusteeRightsQuery | Where-Object { ($_.user -like $TrusteeIdentity.trustee.ToString()) } | Select-Object @{name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } }, @{name = 'Grantor Display Name'; expression = { $GrantorDisplayName } }, @{name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } }, @{name = 'Grantor Environment'; expression = { $GrantorCloudOrOnPrem } }, @{Name = 'Trustee Primary SMTP'; Expression = { $TrusteePrimarySMTP } }, @{Name = 'Trustee Display Name'; Expression = { $TrusteeDisplayName } }, @{Name = 'Trustee Original Identity'; Expression = { $TrusteeIdentity.trustee.ToString() } }, @{name = 'Trustee Recipient Type'; expression = { $TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails } }, @{name = 'Trustee Environment'; expression = { $TrusteeCloudOrOnPrem } }, @{name = 'Permission(s)'; expression = { [string]::join(', ', @($_.AccessRights)) } }, @{Name = 'Folder Name'; Expression = { '' } }, @{Name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } }, @{Name = 'Grantor OU'; Expression = { $GrantorOU } }, @{Name = 'Trustee OU'; Expression = { $TrusteeOU } } | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                        }
                        if ($error) {
                            $ErrorCount++
                            $Text += ' #white:red#ERROR#'
                            "==============================" | Out-File $ErrorFile -Append
                            ("{0:000000}/{1:000000}: {2}, AccessRights" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append
                            for ($e = ($error.count - 1); $e -ge 0; $e--) {
                                $error[$e] | Out-File $ErrorFile -Append
                                "" | Out-File $ErrorFile -Append
                            }
                            "" | Out-File $ErrorFile -Append; "" | Out-File $ErrorFile -Append
                        }
                        $ErrorActionPreference = "Continue"
                        $WarningPreference = "Continue"
                        $error.clear()
                    }
                }
                $GrantorLegacyExchangeDN = ""
            }


            # Send As
            if ($ExportSendAs -eq $true) {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', SendAs'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    try {
                        if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                        $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-RecipientPermission -identity $($args[0]) -resultsize unlimited -wa stop -ea stop } -ArgumentList $GrantorDN) | Where-Object { ($_.Trustee -inotlike 'NT AUTHORITY\*') -and ($_.AccessRights -like '*SendAs*') }
                    } catch {
                    }
                } else {
                    try {
                        $TrusteeIdentityQuery = (Get-ADPermission -identity $GrantorDN -wa stop -ea stop | Where-Object { ($_.user -notlike 'NT AUTHORITY\*') -and ($_.ExtendedRights -like '*Send-As*') } | Select-Object *, @{Name = "trustee"; Expression = { $_."identity" } })
                    } catch {
                    }
                }
                if ($ResolveGroups) {
                    $TrusteeIdentityOriginal = @($TrusteeIdentityQuery.trustee | get_member_recurse | Select-Object @{Name = 'Trustee'; Expression = { $_.identity } })
                } else {
                    $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.identity } })
                }

                if ($error.count -eq 0) {
                    foreach ($TrusteeIdentity in $TrusteeIdentityOriginal.trustee) {
                        if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                        $TrusteeIdentity = $TrusteeIdentity.tostring()
                        try {
                            $TrusteeIdentityQuery = (get-recipient $TrusteeIdentity -resultsize 1 -wa stop -ea stop)
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'Cloud' } else { $TrusteeCloudOrOnPrem = 'On-Prem' }
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'On-Prem' }
                        } catch {
                            try {
                                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient ($($args[0])) -resultsize 1 -wa stop -ea stop } -ArgumentList $TrusteeIdentity)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'On-Prem' } else { $TrusteeCloudOrOnPrem = 'Cloud' }
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'Cloud' }
                            } catch {
                                continue
                            }
                        }
                        $error.clear()
                        $TrusteeIdentityOriginal = $TrusteeIdentity
                        $TrusteeRecipientType = $null
                        $TrusteeRecipientTypeDetails = $null
                        $TrusteeDisplayName = $null
                        $TrusteePrimarySMTP = $null
                        $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                        $TrusteeIdentityQuery | Select-Object @{name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } }, @{name = 'Grantor Display Name'; expression = { $GrantorDisplayName } }, @{name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } }, @{name = 'Grantor Environment'; expression = { $GrantorCloudOrOnPrem } }, @{Name = 'Trustee Primary SMTP'; Expression = { $TrusteePrimarySMTP } }, @{Name = 'Trustee Display Name'; Expression = { $TrusteeDisplayName } }, @{Name = 'Trustee Original Identity'; Expression = { $TrusteeIdentityOriginal } }, @{name = 'Trustee Recipient Type'; expression = { $TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails } }, @{name = 'Trustee Environment'; expression = { $TrusteeCloudOrOnPrem } }, @{Name = 'Permission(s)'; Expression = { 'SendAs' } }, @{Name = 'Folder Name'; Expression = { '' } }, @{Name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } }, @{Name = 'Grantor OU'; Expression = { $GrantorOU } }, @{Name = 'Trustee OU'; Expression = { $TrusteeOU } } | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | Out-File $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, SendAs" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | Out-File $ErrorFile -Append
                        "" | Out-File $ErrorFile -Append
                    }
                    "" | Out-File $ErrorFile -Append; "" | Out-File $ErrorFile -Append
                }
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
            }


            # Send On Behalf
            if (($ExportSendOnBehalf -eq $true)) {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', SendOnBehalf'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    if (($GrantorRecipientType -match 'group') -and ($GrantorRecipientType -notmatch 'DynamicDistributionGroup')) {
                        if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                        try {
                            $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-distributiongroup -identity $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $GrantorDN) | Where-Object { $_.GrantSendOnBehalfto -ne '' }
                        } catch {
                        }
                        $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.GrantSendonBehalfto } })
                    } else {
                        if (($GrantorRecipientType -like 'DynamicDistributionGroup')) {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            try {
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $GrantorDN) | Where-Object { $_.GrantSendOnBehalfto -ne '' }
                            } catch {
                            }
                            $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.GrantSendonBehalfto } })
                        } else {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            try {
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-Mailbox -identity $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $GrantorDN) | Where-Object { $_.GrantSendOnBehalfto -ne '' }
                            } catch {
                            }
                        }
                    }
                } else {
                    if (($GrantorRecipientType -match 'group') -and ($GrantorRecipientType -notmatch 'DynamicDistributionGroup')) {
                        try { $TrusteeIdentityQuery = (Get-distributiongroup -identity $GrantorDN -resultsize 1 -wa stop -ea stop | Where-Object { $_.GrantSendOnBehalfto -ne '' }) } catch {}
                        $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.GrantSendonBehalfto } })
                    } else {
                        if (($GrantorRecipientType -like 'DynamicDistributionGroup')) {
                            try { $TrusteeIdentityQuery = (Get-recipient -identity $GrantorDN -resultsize 1 -wa stop -ea stop | Where-Object { $_.GrantSendOnBehalfto -ne '' }) } catch {}
                            $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.GrantSendonBehalfto } })
                        } else {
                            try { $TrusteeIdentityQuery = (Get-Mailbox -identity $GrantorDN -resultsize 1 -wa stop -ea stop | Where-Object { $_.GrantSendOnBehalfto -ne '' }) } catch {}
                        }
                    }
                }
                if ($ResolveGroups) {
                    $TrusteeIdentityOriginal = @($TrusteeIdentityQuery.GrantSendOnBehalfTo | get_member_recurse | Select-Object @{Name = 'Trustee'; Expression = { $_.identity } })
                } else {
                    $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.GrantSendonBehalfto } })
                }

                if ($error.count -eq 0) {
                    foreach ($TrusteeIdentity in $TrusteeIdentityOriginal.trustee) {
                        if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                        try {
                            $TrusteeIdentityQuery = (get-recipient ($TrusteeIdentity) -resultsize 1 -wa stop -ea stop)
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'Cloud' } else { $TrusteeCloudOrOnPrem = 'On-Prem' }
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'On-Prem' }
                        } catch {
                            try {
                                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $TrusteeIdentity)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'On-Prem' } else { $TrusteeCloudOrOnPrem = 'Cloud' }
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'Cloud' }
                            } catch {
                                continue
                            }
                        }
                        $error.clear()
                        $TrusteeIdentityOriginal = $TrusteeIdentity
                        $TrusteeRecipientType = $null
                        $TrusteeRecipientTypeDetails = $null
                        $TrusteeDisplayName = $null
                        $TrusteePrimarySMTP = $null
                        $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                        $TrusteeIdentityQuery | Select-Object @{name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } }, @{name = 'Grantor Display Name'; expression = { $GrantorDisplayName } }, @{name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } }, @{name = 'Grantor Environment'; expression = { $GrantorCloudOrOnPrem } }, @{Name = 'Trustee Primary SMTP'; Expression = { $TrusteePrimarySMTP } }, @{Name = 'Trustee Display Name'; Expression = { $TrusteeDisplayName } }, @{Name = 'Trustee Original Identity'; Expression = { $TrusteeIdentityOriginal } }, @{name = 'Trustee Recipient Type'; expression = { $TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails } }, @{name = 'Trustee Environment'; expression = { $TrusteeCloudOrOnPrem } }, @{Name = 'Permission(s)'; Expression = { 'SendOnBehalf' } }, @{Name = 'Folder Name'; Expression = { '' } }, @{Name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } }, @{Name = 'Grantor OU'; Expression = { $GrantorOU } }, @{Name = 'Trustee OU'; Expression = { $TrusteeOU } } | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | Out-File $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, SendOnBehalf" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | Out-File $ErrorFile -Append
                        "" | Out-File $ErrorFile -Append
                    }
                    "" | Out-File $ErrorFile -Append; "" | Out-File $ErrorFile -Append
                }
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
            }


            # Managed By
            if (($ExportManagedBy -eq $true) -and ($GrantorRecipientType -Match 'group')) {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', ManagedBy'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    try {
                        if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                        $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-Recipient -identity $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $GrantorDN) | Where-Object { $_.ManagedBy -ne '' }
                    } catch {
                    }
                } else {
                    try {
                        $TrusteeIdentityQuery = (Get-Recipient -identity $GrantorDN -resultsize 1 -wa stop -ea stop | Where-Object { $_.ManagedBy -ne '' } | Select-Object *, @{Name = "trustee"; Expression = { $_."user" } })
                    } catch {
                    }
                }
                if ($ResolveGroups) {
                    $TrusteeIdentityOriginal = @($TrusteeIdentityQuery.ManagedBy | get_member_recurse | Select-Object @{Name = 'Trustee'; Expression = { $_.identity } })
                } else {
                    $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | Select-Object @{Name = 'Trustee'; Expression = { $_.ManagedBy } })
                }

                if ($error.count -eq 0) {
                    foreach ($TrusteeIdentity in $TrusteeIdentityOriginal.trustee) {
                        if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                        $trusteeidentity = $trusteeidentity.tostring()
                        try {
                            $TrusteeIdentityQuery = (get-user ($TrusteeIdentity) -resultsize 1 -wa stop -ea stop)
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'Cloud' } else { $TrusteeCloudOrOnPrem = 'On-Prem' }
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'On-Prem' }
                        } catch {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            try {
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $TrusteeIdentity)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'On-Prem' } else { $TrusteeCloudOrOnPrem = 'Cloud' }
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'Cloud' }
                            } catch {
                                continue
                            }
                        }
                        $error.clear()
                        $TrusteeIdentityOriginal = $TrusteeIdentity
                        $TrusteeRecipientType = $null
                        $TrusteeRecipientTypeDetails = $null
                        $TrusteeDisplayName = $null
                        $TrusteePrimarySMTP = $null
                        $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                        $TrusteeIdentityQuery | Select-Object @{name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } }, @{name = 'Grantor Display Name'; expression = { $GrantorDisplayName } }, @{name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } }, @{name = 'Grantor Environment'; expression = { $GrantorCloudOrOnPrem } }, @{Name = 'Trustee Primary SMTP'; Expression = { $TrusteePrimarySMTP } }, @{Name = 'Trustee Display Name'; Expression = { $TrusteeDisplayName } }, @{Name = 'Trustee Original Identity'; Expression = { $TrusteeIdentityOriginal } }, @{name = 'Trustee Recipient Type'; expression = { $TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails } }, @{name = 'Trustee Environment'; expression = { $TrusteeCloudOrOnPrem } }, @{Name = 'Permission(s)'; Expression = { 'ManagedBy' } }, @{Name = 'Folder Name'; Expression = { '' } }, @{Name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } }, @{Name = 'Grantor OU'; Expression = { $GrantorOU } }, @{Name = 'Trustee OU'; Expression = { $TrusteeOU } } | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | Out-File $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, ManagedBy" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | Out-File $ErrorFile -Append
                        "" | Out-File $ErrorFile -Append
                    }
                    "" | Out-File $ErrorFile -Append; "" | Out-File $ErrorFile -Append
                }
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
            }


            # Folder permissions
            if (($ExportFolderPermissions -eq $true) -and ($GrantorRecipientType -NotMatch 'group')) {
                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', Folders'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                    try {
                        $Folders = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-MailboxFolderStatistics -identity $($args[0]) } -ArgumentList $GrantorDN) | ForEach-Object { $_.folderpath } | ForEach-Object { $_.replace('/', '\') }
                    } catch {
                    }
                } else {
                    try {
                        $Folders = Get-MailboxFolderStatistics -identity $GrantorDN | ForEach-Object { $_.folderpath } | ForEach-Object { $_.replace('/', '\') }
                    } catch {
                    }
                }
                $FolderCount = 1
                if ($error.count -eq 0) {
                    ForEach ($Folder in $Folders) {
                        $FolderKey = $Alias + ':' + $Folder
                        $Permissions = $null
                        if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            $Permissions = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { Get-MailboxFolderPermission -identity $($args[0]) -wa silentlycontinue -ea silentlycontinue } -ArgumentList $FolderKey) | Where-Object { $_.user.usertype -inotlike 'Default' -and $_.user.usertype -inotlike 'Anonymous' -and $_.user.displayname -inotlike $Recipient.DisplayName }
                        } else {
                            $Permissions = Get-MailboxFolderPermission -identity $FolderKey -wa silentlycontinue -ea silentlycontinue | Where-Object { $_.user.usertype -inotlike 'Default' -and $_.user.usertype -inotlike 'Anonymous' -and $_.user.displayname -inotlike $Recipient.DisplayName }
                        }
                        if ($permissions -eq $null) { continue }
                        foreach ($TrusteeIdentity in
                            $(if ($ResolveGroups) {
                                    ($Permissions.user.adrecipient.identity | get_member_recurse).displayname
                                } else {
                                    $permissions.user.adrecipient.identity
                                })
                        ) {
                            if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                            $trusteeidentity = $trusteeidentity.tostring()
                            try {
                                $TrusteeIdentityQuery = (get-recipient ($TrusteeIdentity) -resultsize 1 -wa stop -ea stop)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'Cloud' } else { $TrusteeCloudOrOnPrem = 'On-Prem' }
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'On-Prem' }
                            } catch {
                                if ($ExportFromCloud -eq $true) { Connect-ExchangeOnline }
                                try {
                                    $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock { get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop } -ArgumentList $TrusteeIdentity)
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") { $TrusteeCloudOrOnPrem = 'On-Prem' } else { $TrusteeCloudOrOnPrem = 'Cloud' }
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") { $TrusteeCloudOrOnPrem = 'Cloud' }
                                } catch {
                                    continue
                                }
                            }
                            $error.clear()
                            $TrusteeIdentityOriginal = $TrusteeIdentity
                            $TrusteeRecipientType = $null
                            $TrusteeRecipientTypeDetails = $null
                            $TrusteeDisplayName = $null
                            $TrusteePrimarySMTP = $null
                            $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                            $TrusteeIdentityQuery | Select-Object @{name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } }, @{name = 'Grantor Display Name'; expression = { $GrantorDisplayName } }, @{name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } }, @{name = 'Grantor Environment'; expression = { $GrantorCloudOrOnPrem } }, @{Name = 'Trustee Primary SMTP'; Expression = { $TrusteePrimarySMTP } }, @{Name = 'Trustee Display Name'; Expression = { $TrusteeDisplayName } }, @{Name = 'Trustee Original Identity'; Expression = { $TrusteeIdentityOriginal } }, @{name = 'Trustee Recipient Type'; expression = { $TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails } }, @{name = 'Trustee Environment'; expression = { $TrusteeCloudOrOnPrem } }, @{Name = 'Permission(s)'; Expression = { [string]::join(', ', @($Permissions | Where-Object { $_.User -like $trusteeidentity }).accessrights) } }, @{Name = 'Folder Name'; Expression = { $Folder } }, @{Name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } }, @{Name = 'Grantor OU'; Expression = { $GrantorOU } }, @{Name = 'Trustee OU'; Expression = { $TrusteeOU } } | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                        }
                        $FolderCount++
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | Out-File $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, FolderPermissions" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | Out-File $ErrorFile -Append
                        "" | Out-File $ErrorFile -Append
                    }
                    "" | Out-File $ErrorFile -Append; "" | Out-File $ErrorFile -Append
                    $ErrorActionPreference = "Continue"
                    $WarningPreference = "Continue"
                    $error.clear()
                }
            }
            $count++
            $text | Out-File ($Exportfile + "_Job") -Append -Force
            [System.GC]::Collect() # garbage collection
        }
        if (($ExportFromCloud -eq $true)) {
            Remove-PSSession $script:BatchSessionCloud
        }
        Write-Host "Done."
    } -Name ("$RecipientStartID" + "_Job") -ArgumentList $RecipientStartID, $RecipientEndID, $Exportfile, $ErrorFile, $TempRecipientFile, $ExportFromOnPrem, $ExportFromCloud, $CredentialPasswordFile, $CredentialUsernameFile, $ExportAccessRights, $ExportSendAs, $ExportSendOnBehalf, $ExportManagedBy, $ExportFolderPermissions, $ExportFullAccessPerTrustee, $TranscriptFile, $ResolveGroups | Out-Null
    $Batch = $Batch + 1
}

# Wait for all remaining jobs to complete and results are ready to be received
while ($true) {
    foreach ($x in (Get-Job -State Completed)) {
        $temp = $null
        $TempPath = $null
        # show temp job output file, delete output file
        $TempPath = ($Exportfile + '_temp' + $x.Name)
        if (Test-Path $TempPath) {
            $temp = Get-Content $TempPath
            $temp | Write-HostColored
            Remove-Item $TempPath -Force

        }
        $temp = $null
        $TempPath = $null

        # append temp error file and delete temp file
        $TempPath = ($Errorfile + '_temp' + $x.Name)
        $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
        if (Test-Path $TempPath) {
            $temp = Get-Content $TempPath
            $temp | Out-File $Errorfile -Append -Force
            Remove-Item $TempPath -Force
        }
        $temp = $null
        $TempPath = $null

        # append temp transcript file and delete temp file
        $TempPath = ($Transcriptfile + '_temp' + $x.Name)
        $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
        if (Test-Path $TempPath) {
            $temp = Get-Content $TempPath
            $temp | Out-File $Transcriptfile -Append -Force
            Remove-Item $TempPath -Force
        }
        $temp = $null
        $TempPath = $null

        # append temp export file and delete temp file
        $TempPath = ($Exportfile + '_temp' + $x.Name)
        $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
        if (Test-Path $TempPath) {
            $temp = Import-Csv $TempPath -Delimiter ";"
            $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
            Remove-Item $TempPath -Force
        }
        $temp = $null
        $TempPath = $null
        Remove-Job -Job $x -Force
    }
    [System.GC]::Collect() # garbage collection
    Start-Sleep -s 5

    # end loop when no more completed jobs and no more running jobs are left
    if ((@(Get-Job -State running).count -eq 0) -and (@(Get-Job -State completed).count -eq 0)) { break }
}
if (($ExportAccessRights -eq $true) -and ($ExportFullAccessPerTrustee -eq $true)) {
    Write-Host 'Creating full access permission files per trustee.'
    $AllowedChars = @("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    $PrimarySMTPAddressesToIgnore = @("xxx@domain.com", "yyy@domain.com") #List of primary SMTP addresses to ignore (service account, for example). Wildcards are not allowed.
    $RecipientPermissions = Import-Csv $ExportFile -Delimiter ';' | Select-Object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name', 'Permission(s)', 'Grantor Environment', 'Trustee Environment' | Sort-Object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name', 'Permission(s)', 'Grantor Environment', 'Trustee Environment'
    for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
        if (($RecipientPermissions[$x].'Permission(s)' -like "*FullAccess*") -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne $RecipientPermissions[$x].'Grantor Primary SMTP') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -notin $PrimarySMTPAddressesToIgnore) -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -notin $PrimarySMTPAddressesToIgnore) -and ($RecipientPermissions[$x].'Grantor Environment' -ne $RecipientPermissions[$x].'Trustee Environment')) {
            if ($AllowedChars.Contains($RecipientPermissions[$x].'Trustee Primary SMTP'.substring(0, 1).tolower()) -eq $true) {
                $FileName = 'prefix_' + $RecipientPermissions[$x].'Trustee Primary SMTP'.substring(0, 1).tolower() + '.csv'
            } else {
                $FileName = 'prefix__.csv'
            }

            $RecipientPermissions[$x].'Trustee Primary SMTP' = $RecipientPermissions[$x].'Trustee Primary SMTP'.ToLower()
            $RecipientPermissions[$x].'Grantor Primary SMTP' = $RecipientPermissions[$x].'Grantor Primary SMTP'.ToLower()
            $RecipientPermissions[$x].'Grantor LegacyExchangeDN' = $RecipientPermissions[$x].'Grantor LegacyExchangeDN'
            $RecipientPermissions[$x].'Grantor Display Name' = $RecipientPermissions[$x].'Grantor Display Name'
            $RecipientPermissions[$x] | Select-Object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name' | Export-Csv ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee\' + $FileName) -Append -Force -NoTypeInformation -Delimiter ";"
        }
    }

    if ($TargetFolder -ne "") {
        if (Test-Path $TargetFolder) {
            Get-ChildItem ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee\") -Filter 'prefix_*.csv' -File | ForEach-Object {
                $x = Import-Csv $_.fullname -Delimiter ";" | Select-Object * -Unique
                $x | Export-Csv $_.fullname -NoTypeInformation -Force -Delimiter ';'
                $x = $null
                if (Test-Path ($TargetFolder + "\" + $_.Name)) {
                    # File exists at target, compare MD5 hashes with source.
                    if ((Get-FileHash $_.FullName -Algorithm MD5).hash -eq (Get-FileHash ($TargetFolder + '\' + $_.Name) -Algorithm MD5).hash) {
                        # MD5 hashes are equal, file does not need to be copied
                    } else {
                        # MD5 hashes are not equal, file needs to be copied.
                        Copy-Item $_.fullname $TargetFolder -Force
                    }
                } else {
                    # File does not exist at target, copy file.
                    Copy-Item $_.fullname $TargetFolder -Force
                }
            }

            Get-ChildItem $TargetFolder -Filter 'prefix_*.csv' -File | ForEach-Object {
                if (-not (Test-Path (((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee\' + $_.Name)))) {
                    # File does not exist at source. Delete at target.
                    Remove-Item $_.FullName -Force
                }
            }
        } else {
            Write-Host "Folder $TargetFolder does not exist."
        }
    }
}

Write-Host 'Cleaning output file.'
$RecipientPermissions = Import-Csv $ExportFile -Delimiter ';' | Select-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission(s)', 'Folder Name', 'Grantor OU', 'Trustee OU' | Sort-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission(s)', 'Folder Name', 'Grantor OU', 'Trustee OU'
$RecipientPermissions | Export-Csv $ExportFile -NoTypeInformation -Force -Delimiter ';'

if (Test-Path $TempRecipientFile) { (Remove-Item $TempRecipientFile -Force) }

Stop-Transcript
$TempPath = ($Transcriptfile + '_temp')
if (Test-Path $TempPath) {
    $temp = Get-Content $TempPath
    $temp | Out-File $TranscriptFile -Append -Force
    Remove-Item $TempPath -Force
}
$temp = $null
$TempPath = $null

Write-Host 'Script completed.'