[CmdletBinding(PositionalBinding = $false)]


Param(
    # Environments to consider: Microsoft 365 (Exchange Online) or Exchange on premises
    [boolean]$ExportFromOnPrem = $false, # $true exports from on-prem, $false from Exchange Online
    [string]$ExchangeConnectionUri = 'https://outlook.office365.com/powershell-liveid/',
    # Permission types to export
    [boolean]$ExportAccessRights = $true, # Rights like "FullAccess" and "ReadAccess" to the entire mailbox
    [boolean]$ExportFullAccessPerTrustee = $false, # Additionally export a list which user has full access to which mailbox (legacyExchangeDN) for use with tools such as OutlookRedemption
    [boolean]$ExportSendAs = $true, # Send As
    [boolean]$ExportSendOnBehalf = $true, # Send On Behalf
    [boolean]$ExportManagedby = $true, # Only valid for groups
    [boolean]$ExportFolderPermissions = $true, # Export permissions set on specific mailbox folders. This will take very long.
    [boolean]$ResolveGroups = $false, # Resolve trustee groups to individual members (recursively)

    # Name of the permission export file
    [string]$ExportFile = 'c:\temp\Export-RecipientPermissions_Output\Export-RecipientPermissions_Output.csv',

    # Name of the error file
    [string]$ErrorFile = 'c:\temp\Export-RecipientPermissions_Output\Export-RecipientPermissions_Errors.txt',

    # Name of the transcript file
    [string]$TranscriptFile = 'c:\temp\Export-RecipientPermissions_Output\Export-RecipientPermissions_Transcript.txt',

    # Name of temporary recipient file
    [string]$TempRecipientFile = 'c:\temp\Export-RecipientPermissions_Output\Export-RecipientPermissions_Recipients.csv',

    # Folder to additionally store files created when $ExportFullAccessPerTrustee = $true. This folder must already exist at runtime. Set to "" when not needed.
    [string]$TargetFolder = '',

    # Parallelization
    # Watch RAM and CPU usage
    [int]$NumberOfJobsParallel = 30, # Each job is a separate session towards Exchange on-prem and Office 365, so watch your maximum concurreny settings
    [int]$RecipientsPerJob = 100, # More recipients save time as jobs run longer, but the risk of a problem with the O365 connection is higher

    # User name and password are stored in secure string format
    [string]$CredentialPasswordFile = 'c:\temp\Export-RecipientPermissions_CredentialPassword.txt',
    [string]$CredentialUsernameFile = 'c:\temp\Export-RecipientPermissions_CredentialUsername.txt'
)


#
# Do not change anything from here on.
#

$script:ExchangeSession = $null


function ConnectExchangeAndKeepAlive {
    $Stoploop = $false
    [int]$Retrycount = 0

    while ($Stoploop -eq $false) {
        if ($ExportFromOnPrem -eq $false) {
            try {
                if (-not (Get-PSSession | Where-Object { ($_.name -like 'ExchangeSession') -and ($_.state -like 'opened') })) {
                    $CloudUser = Get-Content $CredentialUsernameFile
                    $CloudPassword = Get-Content $CredentialPasswordFile | ConvertTo-SecureString
                    $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CloudUser, $CloudPassword
                    $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $UserCredential -Authentication Basic -AllowRedirection -Name 'ExchangeSession' -ErrorAction Stop
                    Import-PSSession -Session $script:ExchangeSession -DisableNameChecking -AllowClobber | Out-Null
                    Get-Recipient -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction stop > $null
                }
                $Stoploop = $true
            } catch {
                if ($Retrycount -le 3) {
                    Write-Host 'Could not connect to Exchange Online. Trying again in 70 seconds.'
                    Start-Sleep -Seconds 70
                    $Retrycount = $Retrycount + 1
                } else {
                    Write-Host 'Could not connect to Exchange Online after three retires. Exiting.'
                    exit
                    $Stoploop = $true
                }
            }
        } else {
            try {
                if (-not (Get-PSSession | Where-Object { ($_.name -like 'ExchangeSession') -and ($_.state -like 'opened') })) {
                    $env:tmp = 'c:\alexclude\PowerShell.temp'
                    Get-ChildItem $env:tmp -Directory | Where-Object { $_.LastWriteTime -le (Get-Date).adddays(-2) } | Remove-Item -Force -Recurse
                    $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Authentication Kerberos
                    Import-PSSession -Session $script:ExchangeSession -DisableNameChecking -AllowClobber -Name 'ExchangeSession' | Out-Null
                    Set-AdServerSettings -ViewEntireForest $True
                    (Get-Recipient -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction stop)  > $null
                    $Stoploop = $true
                }
            } catch {
                if ($Retrycount -le 3) {
                    Write-Host 'Could not connect to Exchange on-prem. Trying again in 70 seconds.'
                    Start-Sleep -Seconds 70
                    $Retrycount = $Retrycount + 1
                } else {
                    Write-Host 'Could not connect to Exchange on-prem after three retires. Exiting.'
                    exit
                    $Stoploop = $true
                }
            }
        }
    }
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
            $tokens = $Text.split('#')

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
(New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile))  > $null
(New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile))  > $null
(New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile))  > $null
(New-Item -ItemType Directory -Force -Path (Split-Path -Path $TempRecipientFile))  > $null
if (Test-Path $Exportfile) { (Remove-Item $ExportFile -Force) }
if (Test-Path $Errorfile) { (Remove-Item $ErrorFile -Force) }
if (Test-Path $TranscriptFile) { (Remove-Item $TranscriptFile -Force) }
if (Test-Path $TempRecipientFile) { (Remove-Item $TempRecipientFile -Force) }
if (($ExportFullAccessPerTrustee -eq $true) -and ($ExportAccessRights -eq $true)) {
    if (Test-Path ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee')) {
        Remove-Item ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee') -Force -Recurse
    }
    (New-Item -ItemType Directory -Force -Path ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee'))  > $null
}

Start-Transcript -Path ($TranscriptFile + '_temp') -Force

if ($ExportFromOnPrem -eq $false) {
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
        ConnectExchangeAndKeepAlive
        Write-Host 'On-prem connection working.'
    } Catch {
        Write-Host 'On-prem connection does not work. Error executing ''Get-Recipient -ResultSize 1''. Exiting.'
        Write-Host 'Please start the script on an Exchange server with appropriate permissions.'
        exit 1
    }
} else {
    Try {
        Write-Host 'Connecting to Exchange Online.'
        ConnectExchangeAndKeepAlive
        Write-Host 'Cloud connection working.'
    } Catch {
        Write-Host 'Cloud connection does not work. Error executing ''Get-Recipient -ResultSize 1''. Exiting.'
        exit 1
    }
}

# Export list of objects
if ($ExportFromOnPrem -eq $true) {
    Write-Host 'Enumerating on-prem recipients. This may take a long time.'
} else {
    Write-Host 'Enumerating cloud recipients. This may take a long time.'
}


if ($ExportFromOnPrem -eq $true) {
    Get-recipient -RecipientType MailUniversalSecurityGroup, DynamicDistributionGroup, UserMailbox, MailUniversalDistributionGroup, MailUniversalSecurityGroup, MailNonUniversalGroup, MailUser -ResultSize unlimited -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    <#
    Get-recipient -RecipientType MailUniversalSecurityGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType PublicFolder -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType DynamicDistributionGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType UserMailbox -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailUniversalDistributionGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailUniversalSecurityGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailNonUniversalGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailUser -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailContact -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    #>
} else {
    Get-recipient -RecipientType MailUniversalSecurityGroup, DynamicDistributionGroup, UserMailbox, MailUniversalDistributionGroup, MailUniversalSecurityGroup, MailNonUniversalGroup, MailUser -ResultSize unlimited -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    <#
    Get-recipient -RecipientType MailUniversalSecurityGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType PublicFolder -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType DynamicDistributionGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType UserMailbox -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailUniversalDistributionGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailUniversalSecurityGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailNonUniversalGroup -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailUser -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    Get-recipient -RecipientType MailContact -ResultSize 1000 -WarningAction SilentlyContinue | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    #>
}


if (($ExportFromOnPrem -eq $false)) {
    Write-Host 'Disconnecting from cloud services.'
    Remove-PSSession $script:ExchangeSession
    #if ((test-path (Split-Path -Path $script:ExchangeSessionPath.path)) -eq $true) {
    #    Remove-Item (Split-Path -Path $script:ExchangeSessionPath.path) -Recurse -Force
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
                        $temp = Import-Csv $TempPath -Delimiter ';'
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

    (
        Start-Job {
            param(
                $RecipientStartID,
                $RecipientEndID,
                $Exportfile,
                $ErrorFile,
                $TempRecipientFile,
                $ExportFromOnPrem,
                $CredentialPasswordFile,
                $CredentialUsernameFile,
                $ExportAccessRights,
                $ExportSendAs,
                $ExportSendOnBehalf,
                $ExportManagedby,
                $ExportFolderPermissions,
                $ExportFullAccessPerTrustee,
                $TranscriptFile,
                $ResolveGroups,
                $ExchangeConnectionUri
            )


            function ConnectExchangeAndKeepAlive {
                $Stoploop = $false
                [int]$Retrycount = 0

                while ($Stoploop -eq $false) {
                    if ($ExportFromOnPrem -eq $false) {
                        try {
                            if (-not (Get-PSSession | Where-Object { ($_.name -like 'ExchangeSession') -and ($_.state -like 'opened') })) {
                                $CloudUser = Get-Content $CredentialUsernameFile
                                $CloudPassword = Get-Content $CredentialPasswordFile | ConvertTo-SecureString
                                $UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CloudUser, $CloudPassword
                                $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $UserCredential -Authentication Basic -AllowRedirection -Name 'ExchangeSession' -ErrorAction Stop
                                Import-PSSession -Session $script:ExchangeSession -DisableNameChecking -AllowClobber | Out-Null
                                Get-Recipient -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction stop > $null
                            }
                            $Stoploop = $true
                        } catch {
                            if ($Retrycount -le 3) {
                                Write-Host 'Could not connect to Exchange Online. Trying again in 70 seconds.'
                                Start-Sleep -Seconds 70
                                $Retrycount = $Retrycount + 1
                            } else {
                                Write-Host 'Could not connect to Exchange Online after three retires. Exiting.'
                                exit
                                $Stoploop = $true
                            }
                        }
                    } else {
                        try {
                            if (-not (Get-PSSession | Where-Object { ($_.name -like 'ExchangeSession') -and ($_.state -like 'opened') })) {
                                $env:tmp = 'c:\alexclude\PowerShell.temp'
                                Get-ChildItem $env:tmp -Directory | Where-Object { $_.LastWriteTime -le (Get-Date).adddays(-2) } | Remove-Item -Force -Recurse
                                Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Authentication Kerberos) -DisableNameChecking -AllowClobber -Name 'ExchangeSession' | Out-Null
                                Set-AdServerSettings -ViewEntireForest $True
                                (Get-Recipient -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction stop)  > $null
                                $Stoploop = $true
                            }
                        } catch {
                            if ($Retrycount -le 3) {
                                Write-Host 'Could not connect to Exchange on-prem. Trying again in 70 seconds.'
                                Start-Sleep -Seconds 70
                                $Retrycount = $Retrycount + 1
                            } else {
                                Write-Host 'Could not connect to Exchange on-prem after three retires. Exiting.'
                                exit
                                $Stoploop = $true
                            }
                        }
                    }
                }
            }


            filter get_member_recurse {
                if ($_) {
                    try {
                        ConnectExchangeAndKeepAlive
                        $tempObject = $null
                        if ($ExportFromOnPrem -eq $false) {
                            $tempObject = Get-recipient -identity $_.tostring() -resultsize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        } else {
                            $tempObject = Get-recipient -identity $_.tostring() -properties organizationalunit, id -resultsize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        }
                        if ($tempObject) {
                            if ($tempObject.RecipientType -ilike '*group') {
                                if ($ExportFromOnPrem -eq $false) {
                                    Get-DistributionGroupMember $tempObject.identity -resultsize unlimited -ErrorAction silentlycontinue | get_member_recurse
                                } else {
                                    Get-DistributionGroupMember $tempObject.identity -resultsize unlimited -ErrorAction silentlycontinue | get_member_recurse
                                }
                            } else {
                                $tempObject
                            }
                        } else {
                            $_.ToString()
                        }
                    } catch {
                        $_.ToString()
                    }
                }
            }


            $TranscriptFile = $TranscriptFile + '_temp' + $RecipientStartID
            $TranscriptFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TranscriptFile)
            (New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile))  > $null
            Start-Transcript -Path $TranscriptFile -Force
            Write-Host "Preparations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            Set-Location $PSScriptRoot
            $Exportfile = $Exportfile + '_temp' + $RecipientStartID
            $ErrorFile = $ErrorFile + '_temp' + $RecipientStartID
            $Exportfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Exportfile)
            $ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
            (New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile))  > $null
            (New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile))  > $null
            if (Test-Path $Exportfile) { (Remove-Item $ExportFile -Force) }
            if (Test-Path $Errorfile) { (Remove-Item $ErrorFile -Force) }
            Write-Host "  RecipientStartID: $RecipientStartID"
            Write-Host "  RecipientEndID: $RecipientEndID"
            Write-Host "  Time: @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

            $script:ExchangeSession = $null

            $Recipients = (Import-Csv $TempRecipientFile)
            $RecipientCount = $Recipients.count
            $Count = $RecipientStartID + 1
            #$BatchID = ($RecipientStartID / ($RecipientEndID - $RecipientStartID + 1)) + 1
            ConnectExchangeAndKeepAlive

            for ($RecipientStartID; $RecipientStartID -le $RecipientEndID; $RecipientStartID++) {
                Write-Host "RecipientID: $RecipientStartID; '$($Recipients[$RecipientStartID].PrimarySmtpAddress)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                if ($RecipientStartID -ge $RecipientCount) { break }
                ConnectExchangeAndKeepAlive
                $Recipient = $Mailbox = $null

                $Recipient = Get-recipient $Recipients[$RecipientStartID].PrimarySmtpAddress -ResultSize 1
                if ($ExportFromOnPrem -eq $true) {
                    if ($Recipient.RecipientTypeDetails -like 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                } else {
                    if ($Recipient.RecipientTypeDetails -like 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                }
                $GrantorDisplayName = $Recipient.DisplayName.tostring()
                $GrantorPrimarySMTP = $Recipient.PrimarySMTPAddress.tostring()
                $GrantorRecipientType = $Recipient.RecipientType.tostring()
                $GrantorRecipientTypeDetails = $Recipient.RecipientTypeDetails.tostring()
                $GrantorOU = $Recipient.OrganizationalUnit.tostring()
                $GrantorDN = $Recipient.DistinguishedName.tostring()
                if (
                    (($ExportAccessRights -eq $true) -and ($GrantorRecipientType -inotlike '*group') -and ($ExportFullAccessPerTrustee -eq $true)) -or
                    (($ExportSendOnBehalf -eq $true) -and (($grantorrecipienttype -inotlike '*group')))
                ) {
                    $Mailbox = $Recipient | Get-Mailbox -WarningAction silentlycontinue
                    $GrantorLegacyExchangeDN = $Mailbox.LegacyExchangeDN
                } else {
                    $Mailbox = $null
                    $GrantorLegacyExchangeDN = $null
                }
                $Text = ('{0:000000}/{1:000000}: ' -f $count, $RecipientCount) + $GrantorPrimarySMTP + ', ' + $GrantorRecipienttype + '/' + $GrantorRecipientTypeDetails + ', ' + $GrantorEnvironment

                if (($Recipient.Recipienttype -eq 'PublicFolder') -or ($Recipient.Recipienttype -eq 'MailContact')) { $Text += (", recipient type $GrantorRecipientType not supported."); continue }


                # Access Rights (full access etc.)
                if ($ExportAccessRights -eq $true) {
                    Write-Host "  Access rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $Text += ', AccessRights'
                    if ($GrantorRecipientType -inotlike '*group') {
                        try {
                            ConnectExchangeAndKeepAlive
                            $TrusteeRights = $null
                            if ($ExportFromOnPrem -eq $false) {
                                $TrusteeRights = $Recipient | Get-RecipientPermission -ResultSize unlimited -WarningAction SilentlyContinue | Where-Object { ($_.IsInherited -eq $false) -and ($_.trustee -inotlike 'NT AUTHORITY\*') }
                            } else {
                                $TrusteeRights = $Recipient | Get-MailboxPermission -ResultSize unlimited -WarningAction SilentlyContinue | Where-Object { ($_.IsInherited -eq $false) -and ($_.user -inotlike 'NT AUTHORITY\*') } | Select-Object *, @{ name = 'trustee'; Expression = { $_.user } }
                            }

                            if ($TrusteeRights) {
                                foreach ($TrusteeRight in $TrusteeRights) {
                                    ConnectExchangeAndKeepAlive
                                    $trustees = @()
                                    if ($ResolveGroups) {
                                        ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                            $trustees += $_
                                        }
                                    } else {
                                        ($TrusteeRight.trustee) | ForEach-Object {
                                            $temp = Get-recipient -identity $_ -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                            if ($temp) {
                                                $trustees += $temp
                                            } else {
                                                $trustees += $_.tostring()
                                            }
                                        }
                                    }
                                    foreach ($Trustee in $Trustees) {
                                        if ($ExportFromOnPrem -eq $true) {
                                            if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                        } else {
                                            if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                        }

                                        foreach ($AccessRight in $TrusteeRight.AccessRights) {
                                            $temp = $TrusteeRight | Select-Object (
                                                @{ name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } },
                                                @{ name = 'Grantor Display Name'; expression = { $GrantorDisplayName } },
                                                @{ name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } },
                                                @{ name = 'Grantor Environment'; expression = { $GrantorEnvironment } },
                                                @{ name = 'Trustee Primary SMTP'; Expression = { $Trustee.PrimarySmtpAddress } },
                                                @{ name = 'Trustee Display Name'; Expression = { $Trustee.DisplayName } },
                                                @{ name = 'Trustee Original Identity'; Expression = { $_.trustee.ToString() } },
                                                @{ name = 'Trustee Recipient Type'; expression = { $Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails } },
                                                @{ name = 'Trustee Environment'; expression = { $TrusteeEnvironment } },
                                                @{ name = 'Permission'; expression = { $AccessRight } },
                                                @{ name = 'Folder Name'; Expression = { '' } },
                                                @{ name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } },
                                                @{ name = 'Grantor OU'; Expression = { $GrantorOU } },
                                                @{ name = 'Trustee OU'; Expression = { $Trustee.OrganizationalUnit } }
                                            )
                                            Write-Host "    $($temp[0])"
                                            $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';' -ErrorAction Stop -WarningAction Stop
                                        }
                                    }
                                }
                            }
                        } catch {
                            Write-Host "$errorfile"
                            $Text += ' #white:red#ERROR#'
                            '==============================' | Out-File $ErrorFile -Append -Force
                            ('{0:000000}/{1:000000}: {2}, AccessRights' -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append -Force
                            $error[0] | Out-File $ErrorFile -Append -Force
                        }
                    }
                }


                # Send As
                if ($ExportSendAs -eq $true) {
                    Write-Host "  Send As @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $Text += ', SendAs'
                    try {
                        ConnectExchangeAndKeepAlive
                        $TrusteeRights = $null
                        if ($ExportFromOnPrem -eq $false) {
                            $TrusteeRights = Get-RecipientPermission -identity $GrantorDN -ResultSize unlimited | Where-Object { ($_.Trustee -inotlike 'NT AUTHORITY\*') -and ($_.AccessRights -like '*SendAs*') }
                        } else {
                            $TrusteeRights = Get-ADPermission -identity $GrantorDN | Where-Object { ($_.user -notlike 'NT AUTHORITY\*') -and ($_.ExtendedRights -like '*Send-As*') } | Select-Object *, @{ name = 'trustee'; Expression = { $_.identity } }
                        }

                        if ($TrusteeRights) {
                            foreach ($TrusteeRight in $TrusteeRights) {
                                ConnectExchangeAndKeepAlive
                                $trustees = @()
                                if ($ResolveGroups) {
                                    ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                        $trustees += $_
                                    }
                                } else {
                                    ($TrusteeRight.trustee) | ForEach-Object {
                                        $temp = Get-recipient -identity $_ -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                        if ($temp) {
                                            $trustees += $temp
                                        } else {
                                            $trustees += $_.tostring()
                                        }
                                    }
                                }
                                foreach ($Trustee in $Trustees) {
                                    if ($ExportFromOnPrem -eq $true) {
                                        if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                    } else {
                                        if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                    }

                                    $temp = $TrusteeRight | Select-Object (
                                        @{ name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } },
                                        @{ name = 'Grantor Display Name'; expression = { $GrantorDisplayName } },
                                        @{ name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } },
                                        @{ name = 'Grantor Environment'; expression = { $GrantorEnvironment } },
                                        @{ name = 'Trustee Primary SMTP'; Expression = { $Trustee.PrimarySmtpAddress } },
                                        @{ name = 'Trustee Display Name'; Expression = { $Trustee.DisplayName } },
                                        @{ name = 'Trustee Original Identity'; Expression = { $_.trustee.ToString() } },
                                        @{ name = 'Trustee Recipient Type'; expression = { $Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails } },
                                        @{ name = 'Trustee Environment'; expression = { $TrusteeEnvironment } },
                                        @{ name = 'Permission'; expression = { 'SendAs' } },
                                        @{ name = 'Folder Name'; Expression = { '' } },
                                        @{ name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } },
                                        @{ name = 'Grantor OU'; Expression = { $GrantorOU } },
                                        @{ name = 'Trustee OU'; Expression = { $Trustee.OrganizationalUnit } }
                                    )
                                    Write-Host "    $($temp[0])"
                                    $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';' -ErrorAction Stop -WarningAction Stop
                                }
                            }
                        }
                    } catch {
                        Write-Host "$errorfile"
                        $Text += ' #white:red#ERROR#'
                        '==============================' | Out-File $ErrorFile -Append -Force
                        ('{0:000000}/{1:000000}: {2}, SendAs' -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append -Force
                        $error[0] | Out-File $ErrorFile -Append -Force
                    }
                }


                # Send On Behalf
                if ($ExportSendOnBehalf -eq $true) {
                    Write-Host "  Send On Behalf @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $Text += ', SendOnBehalf'
                    try {
                        ConnectExchangeAndKeepAlive
                        $TrusteeRights = $null

                        if ($ExportFromOnPrem -eq $false) {
                            if (($GrantorRecipientType -ilike '*group') -and ($GrantorRecipientType -ine 'DynamicDistributionGroup') -and ($GrantorRecipientTypeDetails -ine 'groupmailbox')) {
                                $TrusteeRights = ($Recipient | Get-distributiongroup).GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                            } else {
                                if ($GrantorRecipientTypeDetails -ieq 'groupmailbox') {
                                    $TrusteeRights = ($Recipient | Get-UnifiedGroup).GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                } elseif (($GrantorRecipientType -ieq 'DynamicDistributionGroup')) {
                                    $TrusteeRights = $Recipient.GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                } else {
                                    $TrusteeRights = $Mailbox.GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                }
                            }
                        } else {
                            if (($GrantorRecipientType -ilike '*group') -and ($GrantorRecipientType -ine 'DynamicDistributionGroup')) {
                                $TrusteeRights = ($Recipient | Get-distributiongroup).GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                            } else {
                                if (($GrantorRecipientType -ieq 'DynamicDistributionGroup')) {
                                    $TrusteeRights = $Recipient.GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                } else {
                                    $TrusteeRights = $Mailbox.GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                }
                            }
                        }

                        if ($TrusteeRights) {
                            foreach ($TrusteeRight in $TrusteeRights) {
                                ConnectExchangeAndKeepAlive
                                $trustees = @()
                                if ($ResolveGroups) {
                                    ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                        $trustees += $_
                                    }
                                } else {
                                    ($TrusteeRight.trustee) | ForEach-Object {
                                        $temp = Get-recipient -identity $_ -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                        if ($temp) {
                                            $trustees += $temp
                                        } else {
                                            $trustees += $_.tostring()
                                        }
                                    }
                                }
                                foreach ($Trustee in $Trustees) {
                                    if ($ExportFromOnPrem -eq $true) {
                                        if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                    } else {
                                        if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                    }

                                    $temp = $TrusteeRight | Select-Object (
                                        @{ name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } },
                                        @{ name = 'Grantor Display Name'; expression = { $GrantorDisplayName } },
                                        @{ name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } },
                                        @{ name = 'Grantor Environment'; expression = { $GrantorEnvironment } },
                                        @{ name = 'Trustee Primary SMTP'; Expression = { $Trustee.PrimarySmtpAddress } },
                                        @{ name = 'Trustee Display Name'; Expression = { $Trustee.DisplayName } },
                                        @{ name = 'Trustee Original Identity'; Expression = { $_.trustee.ToString() } },
                                        @{ name = 'Trustee Recipient Type'; expression = { $Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails } },
                                        @{ name = 'Trustee Environment'; expression = { $TrusteeEnvironment } },
                                        @{ name = 'Permission'; expression = { 'SendOnBehalf' } },
                                        @{ name = 'Folder Name'; Expression = { '' } },
                                        @{ name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } },
                                        @{ name = 'Grantor OU'; Expression = { $GrantorOU } },
                                        @{ name = 'Trustee OU'; Expression = { $Trustee.OrganizationalUnit } }
                                    )
                                    Write-Host "    $($temp[0])"
                                    $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';' -ErrorAction Stop -WarningAction Stop
                                }
                            }
                        }
                    } catch {
                        Write-Host "$errorfile"
                        $Text += ' #white:red#ERROR#'
                        '==============================' | Out-File $ErrorFile -Append -Force
                        ('{0:000000}/{1:000000}: {2}, SendOnBehalf' -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append -Force
                        $error[0] | Out-File $ErrorFile -Append -Force
                    }
                }


                # Managed By
                if (($ExportManagedby -eq $true) -and ($GrantorRecipientType -ilike '*group')) {
                    Write-Host "  Managed By @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $Text += ', ManagedBy'
                    try {
                        ConnectExchangeAndKeepAlive
                        $TrusteeRights = $null
                        if ($ExportFromOnPrem -eq $false) {
                            $TrusteeRights = $Recipient.Managedby | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                        } else {
                            $TrusteeRights = $Recipient.Managedby | Select-Object *, @{ name = 'trustee'; Expression = { $_.user } }
                        }

                        if ($TrusteeRights) {
                            foreach ($TrusteeRight in $TrusteeRights) {
                                ConnectExchangeAndKeepAlive
                                $trustees = @()
                                if ($ResolveGroups) {
                                    ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                        $trustees += $_
                                    }
                                } else {
                                    ($TrusteeRight.trustee) | ForEach-Object {
                                        $temp = Get-recipient -identity $_ -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                        if ($temp) {
                                            $trustees += $temp
                                        } else {
                                            $trustees += $_.tostring()
                                        }
                                    }
                                }
                                foreach ($Trustee in $Trustees) {
                                    if ($ExportFromOnPrem -eq $true) {
                                        if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                    } else {
                                        if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                    }

                                    $temp = $TrusteeRight | Select-Object (
                                        @{ name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } },
                                        @{ name = 'Grantor Display Name'; expression = { $GrantorDisplayName } },
                                        @{ name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } },
                                        @{ name = 'Grantor Environment'; expression = { $GrantorEnvironment } },
                                        @{ name = 'Trustee Primary SMTP'; Expression = { $Trustee.PrimarySmtpAddress } },
                                        @{ name = 'Trustee Display Name'; Expression = { $Trustee.DisplayName } },
                                        @{ name = 'Trustee Original Identity'; Expression = { $_.trustee.ToString() } },
                                        @{ name = 'Trustee Recipient Type'; expression = { $Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails } },
                                        @{ name = 'Trustee Environment'; expression = { $TrusteeEnvironment } },
                                        @{ name = 'Permission'; expression = { 'ManagedBy' } },
                                        @{ name = 'Folder Name'; Expression = { '' } },
                                        @{ name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } },
                                        @{ name = 'Grantor OU'; Expression = { $GrantorOU } },
                                        @{ name = 'Trustee OU'; Expression = { $Trustee.OrganizationalUnit } }
                                    )
                                    Write-Host "    $($temp[0])"
                                    $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';' -ErrorAction Stop -WarningAction Stop
                                }
                            }
                        }
                    } catch {
                        Write-Host "$errorfile"
                        $Text += ' #white:red#ERROR#'
                        '==============================' | Out-File $ErrorFile -Append -Force
                        ('{0:000000}/{1:000000}: {2}, ManagedBy' -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append -Force
                        $error[0] | Out-File $ErrorFile -Append -Force
                    }
                }


                # Folder permissions
                if (($ExportFolderPermissions -eq $true) -and ($GrantorRecipientType -NotMatch 'group')) {
                    Write-Host "  Folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $Text += ', Folders'
                    try {
                        ConnectExchangeAndKeepAlive
                        $Folders = Get-MailboxFolderStatistics -identity $GrantorDn | ForEach-Object { $_.folderpath } | ForEach-Object { $_.replace('/', '\') }
                        $Folders = ($Folders += '\') | Sort-Object # '\' is the root folder of the mailbox
                        if ($error.count -eq 0) {
                            ForEach ($Folder in $Folders) {
                                Write-Host "    $Folder @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                $FolderKey = $GrantorPrimarySMTP + ':' + $Folder
                                $TrusteeRights = $null
                                $TrusteeRights = Get-MailboxFolderPermission -identity $FolderKey -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Where-Object { ($_.user.usertype -ine 'Default') -and ($_.user.usertype -ine 'Anonymous') -and ($_.user.displayname -ine $Recipient.DisplayName)}

                                if ($TrusteeRights) {
                                    foreach ($TrusteeRight in $TrusteeRights) {
                                        ConnectExchangeAndKeepAlive
                                        $trustees = @()

                                        if ($ResolveGroups) {
                                            if ($ExportFromOnPrem -eq $true) {
                                                $TrusteeRight.user.adrecipient.alias | Where-Object { $_ } | get_member_recurse | ForEach-Object {
                                                    $trustees += $_
                                                }
                                            } else {
                                                $TrusteeRight.user.recipientprincipal.alias | Where-Object { $_ } | get_member_recurse | ForEach-Object {
                                                    $trustees += $_
                                                }
                                            }
                                        } else {
                                            if ($ExportFromOnPrem -eq $true) {
                                                $TrusteeRight.user.adrecipient.alias | Where-Object { $_ } | ForEach-Object {
                                                    $temp = Get-recipient -identity $_ -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                                    if ($temp) {
                                                        $trustees += $temp
                                                    } else {
                                                        $trustees += $_.tostring()
                                                    }
                                                }
                                            } else {
                                                $TrusteeRight.user.recipientprincipal.alias | Where-Object { $_ } | ForEach-Object {
                                                    $temp = Get-recipient -identity $_ -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                                    if ($temp) {
                                                        $trustees += $temp
                                                    } else {
                                                        $trustees += $_.tostring()
                                                    }
                                                }
                                            }
                                        }

                                        if ($ExportFromOnPrem -eq $true) {
                                            $TrusteeRight.user.adrecipient.originalsmtpaddress | Where-Object { $_ } | ForEach-Object {
                                                $trustees += $_.tostring()
                                            }
                                        } else {
                                            $TrusteeRight.user.recipientprincipal.originalsmtpaddress | Where-Object { $_ } | ForEach-Object {
                                                $trustees += $_.tostring()
                                            }
                                        }

                                        foreach ($Trustee in $Trustees) {
                                            if ($ExportFromOnPrem -eq $true) {
                                                if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -like 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            foreach ($AccessRight in $TrusteeRight.AccessRights) {
                                                $temp = $TrusteeRight | Select-Object (
                                                    @{ name = 'Grantor Primary SMTP'; expression = { $GrantorPrimarySMTP } },
                                                    @{ name = 'Grantor Display Name'; expression = { $GrantorDisplayName } },
                                                    @{ name = 'Grantor Recipient Type'; expression = { $GrantorRecipientType + '/' + $GrantorRecipientTypeDetails } },
                                                    @{ name = 'Grantor Environment'; expression = { $GrantorEnvironment } },
                                                    @{ name = 'Trustee Primary SMTP'; Expression = { $Trustee.PrimarySmtpAddress } },
                                                    @{ name = 'Trustee Display Name'; Expression = { $Trustee.DisplayName } },
                                                    @{ name = 'Trustee Original Identity'; Expression = { $_.user.displayname.ToString() } },
                                                    @{ name = 'Trustee Recipient Type'; expression = { $Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails } },
                                                    @{ name = 'Trustee Environment'; expression = { $TrusteeEnvironment } },
                                                    @{ name = 'Permission'; expression = { $AccessRight } },
                                                    @{ name = 'Folder Name'; Expression = { $Folder } },
                                                    @{ name = 'Grantor LegacyExchangeDN'; Expression = { $GrantorLegacyExchangeDN } },
                                                    @{ name = 'Grantor OU'; Expression = { $GrantorOU } },
                                                    @{ name = 'Trustee OU'; Expression = { $Trustee.OrganizationalUnit } }
                                                )
                                                Write-Host "      $($temp[0])"
                                                $temp | Export-Csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';' -ErrorAction Stop -WarningAction Stop
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Host "$errorfile"
                        $Text += ' #white:red#ERROR#'
                        '==============================' | Out-File $ErrorFile -Append -Force
                        ('{0:000000}/{1:000000}: {2}, Folders' -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append -Force
                        $error[0] | Out-File $ErrorFile -Append -Force
                    }
                }

                $count++
                $text | Out-File ($Exportfile + '_Job') -Append -Force
                [System.GC]::Collect() # garbage collection
            }

            Remove-PSSession $script:ExchangeSession

            Write-Host "Done @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        } -Name ("$RecipientStartID" + '_Job') -ArgumentList $RecipientStartID, $RecipientEndID, $Exportfile, $ErrorFile, $TempRecipientFile, $ExportFromOnPrem, $CredentialPasswordFile, $CredentialUsernameFile, $ExportAccessRights, $ExportSendAs, $ExportSendOnBehalf, $ExportManagedby, $ExportFolderPermissions, $ExportFullAccessPerTrustee, $TranscriptFile, $ResolveGroups, $ExchangeConnectionUri
    )  > $null

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
            $temp = Import-Csv $TempPath -Delimiter ';'
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

if (($ExportAccessRights -eq $true) -and ($ExportFullAccessPerTrustee -eq $true) -and (Test-Path $exportfile)) {
    Write-Host 'Creating full access permission files per trustee.'
    $AllowedChars = @('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')
    $PrimarySMTPAddressesToIgnore = @('xxx@domain.com', 'yyy@domain.com') #List of primary SMTP addresses to ignore (service account, for example). Wildcards are not allowed.
    $RecipientPermissions = Import-Csv $ExportFile -Delimiter ';'
    for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
        if (($RecipientPermissions[$x].'Permission' -like '*FullAccess*') -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne $RecipientPermissions[$x].'Grantor Primary SMTP') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -notin $PrimarySMTPAddressesToIgnore) -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -notin $PrimarySMTPAddressesToIgnore)) {
            if ($AllowedChars.Contains($RecipientPermissions[$x].'Trustee Primary SMTP'.substring(0, 1).tolower()) -eq $true) {
                $FileName = 'prefix_' + $RecipientPermissions[$x].'Trustee Primary SMTP'.substring(0, 1).tolower() + '.csv'
            } else {
                $FileName = 'prefix__.csv'
            }

            $RecipientPermissions[$x].'Trustee Primary SMTP' = $RecipientPermissions[$x].'Trustee Primary SMTP'.ToLower()
            $RecipientPermissions[$x].'Grantor Primary SMTP' = $RecipientPermissions[$x].'Grantor Primary SMTP'.ToLower()
            $RecipientPermissions[$x].'Grantor LegacyExchangeDN' = $RecipientPermissions[$x].'Grantor LegacyExchangeDN'
            $RecipientPermissions[$x].'Grantor Display Name' = $RecipientPermissions[$x].'Grantor Display Name'
            $RecipientPermissions[$x] | Select-Object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name' | Export-Csv ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee\' + $FileName) -Append -Force -NoTypeInformation -Delimiter ';'
        }
    }

    if ($TargetFolder -ne '') {
        if (Test-Path $TargetFolder) {
            Get-ChildItem ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee\') -Filter 'prefix_*.csv' -File | ForEach-Object {
                $temp = Import-Csv $_.fullname -Delimiter ';' | Select-Object * -Unique
                $temp | Export-Csv $_.fullname -NoTypeInformation -Force -Delimiter ';'
                if (Test-Path ($TargetFolder + '\' + $_.Name)) {
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


if (Test-Path $ExportFile) {
    Write-Host 'Cleaning output file.'
    $RecipientPermissions = Import-Csv $ExportFile -Delimiter ';' | Select-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission', 'Folder Name', 'Grantor OU', 'Trustee OU' | Sort-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission', 'Folder Name', 'Grantor OU', 'Trustee OU'
    $RecipientPermissions | Export-Csv $ExportFile -NoTypeInformation -Force -Delimiter ';'
}

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