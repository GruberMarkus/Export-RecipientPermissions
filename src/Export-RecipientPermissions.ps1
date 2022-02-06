[CmdletBinding(PositionalBinding = $false)]


Param(
    # Environments to consider: Microsoft 365 (Exchange Online) or Exchange on premises
    [boolean]$ExportFromOnPrem = $true, # $true exports from on-prem, $false from Exchange Online
    [uri[]]$ExchangeConnectionUriList = (
        'http://W01-EX01.sv-services.at/powershell',
        'http://W02-EX02.sv-services.at/powershell',
        'http://W01-EX03.sv-services.at/powershell',
        'http://W02-EX04.sv-services.at/powershell',
        'http://W01-EX05.sv-services.at/powershell',
        'http://W02-EX06.sv-services.at/powershell',
        'http://W01-EX07.sv-services.at/powershell',
        'http://W02-EX08.sv-services.at/powershell',
        'http://W01-EX09.sv-services.at/powershell',
        'http://W02-EX10.sv-services.at/powershell'
    ), #'https://outlook.office365.com/powershell-liveid/'
    [string]$PowershellTempDir = 'c:\alexclude\PowerShell.temp', # Directory for temporary PowerShell files from Import-PSSession, usually $env:tmp
    [string]$GetRecipientParameters = @'
-filter "emailaddresses -like '*@pv.at'" -ResultSize unlimited -WarningAction SilentlyContinue
'@,

    # Permission types to export
    [boolean]$ExportAccessRights = $true, # Rights like "FullAccess" and "ReadAccess" to the entire mailbox
    [boolean]$ExportFullAccessPerTrustee = $false, # Additionally export a list which user has full access to which mailbox (legacyExchangeDN) for use with tools such as OutlookRedemption
    [boolean]$ExportSendAs = $false, # Send As
    [boolean]$ExportSendOnBehalf = $false, # Send On Behalf
    [boolean]$ExportManagedby = $false, # Only valid for groups
    [boolean]$ExportFolderPermissions = $false, # Export permissions set on specific mailbox folders. This will take very long.
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
    [int]$NumberOfJobsParallel = $ExchangeConnectionUriList.count,
    [int]$RecipientsPerJob = 100,

    # User name and password are stored in secure string format
    [string]$CredentialPasswordFile = 'c:\temp\Export-RecipientPermissions_CredentialPassword.txt',
    [string]$CredentialUsernameFile = 'c:\temp\Export-RecipientPermissions_CredentialUsername.txt'
)


#
# Do not change anything from here on.
#

$script:ExchangeSession = $null


$FunctionScriptBlock_ConnectExchangeAndKeepAlive = {
    $env:tmp = $PowershellTempDir
    $env:temp = $PowershellTempDir

    $Stoploop = $false
    [int]$Retrycount = 0
    while ($Stoploop -eq $false) {
        try {
            if (Get-PSSession | Where-Object { ($_ -eq $script:ExchangeSession) -and ($_.state -ieq 'opened') }) {
                $null = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-SecurityPrincipal -resultsize 1 -ErrorAction stop -WarningAction silentlycontinue }
                $Stoploop = $true
            } else {
                if ($script:ExchangeSession) {
                    Remove-PSSession $script:ExchangeSession
                }

                if ($ExportFromOnPrem -eq $false) {
                    $CloudUser = Get-Content $CredentialUsernameFile
                    $CloudPassword = Get-Content $CredentialPasswordFile | ConvertTo-SecureString
                    $script:UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CloudUser, $CloudPassword
                    $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUriList[0] -Credential $script:UserCredential -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
                } else {
                    $CloudUser = Get-Content $CredentialUsernameFile
                    $CloudPassword = Get-Content $CredentialPasswordFile | ConvertTo-SecureString
                    $script:UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CloudUser, $CloudPassword
                    $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUriList[0]  -Credential $script:UserCredential -ErrorAction Stop -WarningAction SilentlyContinue
                }
                #$null = Import-PSSession -Session $script:ExchangeSession -DisableNameChecking -AllowClobber -CommandName * -ErrorAction stop

                if ($ExportFromOnPrem -eq $true) {
                    Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True -ErrorAction stop }
                }
            }
        } catch {

            if ($script:ExchangeSession) {
                Remove-PSSession $script:ExchangeSession
            }

            if ($Retrycount -lt 3) {
                Write-Host $error[0]
                Write-Host 'Could not connect to Exchange. Trying again in 70 seconds.'
                Start-Sleep -Seconds 70
                $Retrycount = $Retrycount + 1
            } else {
                ('==============================', 'Could not connect to Exchange after three retries. Exiting.', "RecipientStartID: $RecipientStartID", "RecipientEndID: $RecipientEndID", "Time: @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@", $Error[0]) | ForEach-Object {
                    Write-Host $_
                    $_ | Out-File $ErrorFile -Append -Force

                }
                exit 1
                $Stoploop = $true
            }
        }
    }
}


function ConnectExchangeAndKeepAlive {
    & $FunctionScriptBlock_ConnectExchangeAndKeepAlive
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

try {
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
    $null = (New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile))
    $null = (New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile))
    $null = (New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile))
    $null = (New-Item -ItemType Directory -Force -Path (Split-Path -Path $TempRecipientFile))
    if (Test-Path $Exportfile) { (Remove-Item $ExportFile -Force) }
    if (Test-Path $Errorfile) { (Remove-Item $ErrorFile -Force) }
    if (Test-Path $TranscriptFile) { (Remove-Item $TranscriptFile -Force) }
    if (Test-Path $TempRecipientFile) { (Remove-Item $TempRecipientFile -Force) }
    if (($ExportFullAccessPerTrustee -eq $true) -and ($ExportAccessRights -eq $true)) {
        if (Test-Path ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee')) {
            Remove-Item ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee') -Force -Recurse
        }
        $null = (New-Item -ItemType Directory -Force -Path ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee'))
    }

    $tempTranscriptFile = [io.path]::ChangeExtension($TranscriptFile, 'TEMP.main.txt')
    Start-Transcript -Path $tempTranscriptFile -Force

    $tempConnectionUriQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    while ($tempConnectionUriQueue.count -lt 10000) {
        foreach ($x in $ExchangeConnectionUriList) {
            $tempConnectionUriQueue.Enqueue($x.AbsoluteUri)
            if ($tempConnectionUriQueue.count -ge $NumberOfJobsParallel) {
                break
            }
        }
    }


    #if ($ExportFromOnPrem -eq $false) {
    if ((Test-Path $CredentialUsernameFile) -and (Test-Path $CredentialPasswordFile)) { } else {
        Write-Host 'Please enter cloud user name for later use.'
        Read-Host | Out-File $CredentialUsernameFile
        Write-Host 'Please enter cloud admin password for later use.'
        Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File $CredentialPasswordFile
    }
    #}

    # Test connection
    ConnectExchangeAndKeepAlive

    # Export list of objects
    Write-Host 'Importing objects. This may take a long time.'
    #(invoke-command -Session $script:ExchangeSession -ScriptBlock $([Scriptblock]::Create('Get-Recipient ' + $GetRecipientParameters + " | Select-Object PrimarySmtpAddress"))) | Select-Object PrimarySmtpAddress | Export-Csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'

    Write-Host "  Recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "    Single-thread operation using connection $($ExchangeConnectionUriList[0])"
    $AllRecipients = [system.collections.arraylist]::Synchronized([system.collections.arraylist](Invoke-Command -Session $script:ExchangeSession -HideComputerName -ScriptBlock { get-recipient -resultsize unlimited -ErrorAction stop -WarningAction silentlycontinue | Select-Object -Property identity, recipienttype, recipienttypedetails, displayname, primarysmtpaddress, managedby, distinguishedname, 'UserFriendlyName' }))

    Write-Host "  Mailboxes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ((($ExportAccessRights -eq $true) -and ($ExportFullAccessPerTrustee -eq $true)) -or ($ExportSendOnBehalf -eq $true)) {
        Write-Host "    Single-thread operation using connection $($ExchangeConnectionUriList[0])"
        $AllMailboxes = [system.collections.arraylist]::Synchronized([system.collections.arraylist](Invoke-Command -Session $script:ExchangeSession -HideComputerName -ScriptBlock { get-mailbox -resultsize unlimited -ErrorAction stop -WarningAction silentlycontinue | Select-Object -Property identity, legacyexchangedn, grantsendonbehalfto }))
    } else {
        Write-Host '    Not required with current export settings.'
    }


    Write-Host "  Access rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($ExportAccessRights) {
        Write-Host "    Multi-thread operation using up to $($NumberOfJobsParallel) connections concurrently"
        Write-Host "    $(($AllRecipients | Where-Object {$_.RecipientType -ilike '*mailbox'}).count) maiboxes to check. Done (in steps of 100): 0, " -NoNewline

        $AllAccessrights = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new())
        $tempRecipientsDone = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())

        $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $NumberOfJobsParallel)
        $RunspacePool.Open()

        $Runspaces = (0..($NumberOfJobsParallel - 1)) | ForEach-Object {
            $Runspace = [powershell]::Create().AddScript(
                {
                    param(
                        $ThreadId,
                        $PowershellTempDir,
                        $AllRecipients,
                        $AllAccessRights,
                        $tempConnectionUriQueue,
                        $tempRecipientsDone
                    )

                    $env:tmp = $PowershellTempDir
                    $env:temp = $PowershellTempDir

                    $connectionUri = $tempConnectionUriQueue.dequeue()

                    $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication Kerberos -Credential $using:UserCredential
                    Invoke-Command -Session $script:ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True }

                    for ($x = $ThreadId; $x -lt $AllRecipients.count; $x = $x + $using:NumberOfJobsParallel) {
                        if ($AllRecipients[$x].RecipientType -ilike '*mailbox') {
                            $AllAccessRights.addrange((Invoke-Command -Session $script:ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $($args[0]) -resultsize unlimited -ErrorAction stop -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, isinherited, deny } -ArgumentList $AllRecipients[$x].PrimarySMTPAddress.Address | Where-Object { ($_.IsInherited -eq $false) -and ($_.trustee -inotlike 'NT AUTHORITY\*') }))
                            $tempRecipientsDone.enqueue($x)
                        }

                    }

                    [System.GC]::Collect() # garbage collection
                    Remove-PSSession $script:ExchangeSession
                }
            ).AddParameters(
                @{
                    ThreadID               = $_
                    PowershellTempDir      = $PowershellTempDir
                    AllRecipients          = $AllRecipients
                    AllAccessRights        = $AllAccessRights
                    tempConnectionUriQueue = $tempConnectionUriQueue
                    tempRecipientsDone     = $tempRecipientsDone
                }
            )

            $Runspace.RunspacePool = $RunspacePool

            [PSCustomObject]@{
                Instance = $Runspace
                State    = $Runspace.BeginInvoke()
            }
        }

        while ( $Runspaces.State.IsCompleted -contains $False) {
            $count = $tempRecipientsDone.count
            if (($count -gt 0) -and ($count % 100 -eq 0)) {
                Write-Host "$($count), " -NoNewline
            }
            Start-Sleep -Seconds 1
        }

        Write-Host
        [System.GC]::Collect() # garbage collection


        Write-Host "  UserfriendlyNames of AccessRights to DistinguishedNmes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        Write-Host "    Threadable operation using up to $($NumberOfJobsParallel) connections concurrently"

        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        (($AllAccessrights | Where-Object { ($_.user.rawidentity -like '*\*') -and ($_.user.rawidentity -inotlike 'NT AUTHORITY\*') }).user.securityidentifier | Select-Object -Unique) | ForEach-Object {
            $tempQueue.enqueue($_)
        }
        $tempQueueCount = $tempQueue.count

        Write-Host "    $($tempQueueCount) unique identities to check. Done (in steps of 100): 0, " -NoNewline


        $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $NumberOfJobsParallel)
        $RunspacePool.Open()

        $Runspaces = (0..($NumberOfJobsParallel - 1)) | ForEach-Object {
            $Runspace = [powershell]::Create().AddScript(
                {
                    param (
                        $ThreadId,
                        $PowershellTempDir,
                        $AllRecipients,
                        $AllAccessRights,
                        $tempConnectionUriQueue,
                        $tempQueue
                    )

                    $env:tmp = $PowershellTempDir
                    $env:temp = $PowershellTempDir

                    $connectionUri = $tempConnectionUriQueue.dequeue()

                    $script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Authentication Kerberos -Credential $using:UserCredential
                    Invoke-Command -Session $script:ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True }

                    while ($tempQueue.count -gt 0) {
                        $x = $tempQueue.dequeue()
                        $temp = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { get-securityprincipal -filter "sid -eq '$($args[0])'" -resultsize 1 | Select-Object userfriendlyname, distinguishedname } -ArgumentList $x
                        if ($temp) {
                            foreach ($Recipient in $AllRecipients) {
                                if ($Recipient.distinguishedname -eq $temp.distinguishedname) {
                                    $Recipient.UserFriendlyName = $temp.UserFriendlyName
                                    break
                                }
                            }
                        }
                    }

                    [System.GC]::Collect() # garbage collection
                    Remove-PSSession $script:ExchangeSession
                }
            ).AddParameters(
                @{
                    ThreadID               = $_
                    PowershellTempDir      = $PowershellTempDir
                    AllRecipients          = $AllRecipients
                    AllAccessRights        = $AllAccessRights
                    tempConnectionUriQueue = $tempConnectionUriQueue
                    tempQueue              = $tempQueue
                }
            )

            $Runspace.RunspacePool = $RunspacePool

            [PSCustomObject]@{
                Instance = $Runspace
                State    = $Runspace.BeginInvoke()
            }
        }

        while ( $Runspaces.State.IsCompleted -contains $False) {
            $count = $tempQueue.count
            if (($count -gt 0) -and ($count % 100 -eq 0)) {
                Write-Host "$($count), " -NoNewline
            }
            Start-Sleep -Seconds 1
        }

        Write-Host
        [System.GC]::Collect() # garbage collection

        Write-Host
        ($AllRecipients | Where-Object { $_.userfriendlyname }).count
        return
    } else {
        Write-Host '    Not required with current export settings.'
    }

    [System.GC]::Collect() # garbage collection

    Write-Host '  Group membership'
    if ($ResolveGroups) {
        #$AllRecipients | where {$_.RecipientType -ilike "*group"}
    } else {
        Write-Host '    Not required with current export settings.'
    }


    Write-Host "  Imports completed @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"


    Write-Host 'Disconnecting from Exchange'
    Remove-PSSession $script:ExchangeSession

    $RecipientCount = $AllRecipients.count
    $Batches = [Math]::Ceiling($RecipientCount / $RecipientsPerJob)

    Write-Host "  $RecipientCount recipients found."
    Write-Host "Reading permissions in $Batches batches of $RecipientsPerJob recipients each."
    Write-Host "  Up to $NumberOfJobsParallel batches will run in parallel."
    Write-Host '  Screen output is updated every time a single recipient is completed.'


    $script:jobs = New-Object System.Collections.ArrayList
    $sessionstate = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $NumberOfJobsParallel)
    $RunspacePool.Open()


    $ResultsForExportfile = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
    $RecipientsDoneForDisplay = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))


    for ($RecipientStartID = 0; $RecipientStartID -lt $RecipientCount; $RecipientStartID += $RecipientsPerJob) {

        $RecipientEndID = [math]::min(($RecipientStartID + $RecipientsPerJob - 1), ($RecipientCount - 1 ))

        $JobErrorFile = [io.path]::ChangeExtension(($ErrorFile), ('TEMP.{0:000000}-{1:000000}.txt' -f ($RecipientStartID + 1), ($RecipientEndID + 1)))
        $JobTranscriptFile = [io.path]::ChangeExtension(($TranscriptFile), ('TEMP.{0:000000}-{1:000000}.txt' -f ($RecipientStartID + 1), ($RecipientEndID + 1)))

        $PowerShell = [powershell]::Create()
        $PowerShell.RunspacePool = $RunspacePool

        [void]$PowerShell.AddScript( {
                param(
                    $RecipientStartID,
                    $RecipientEndID,
                    $ErrorFile,
                    $AllRecipients,
                    $ExportFromOnPrem,
                    $ExportAccessRights,
                    $ExportSendAs,
                    $ExportSendOnBehalf,
                    $ExportManagedby,
                    $ExportFolderPermissions,
                    $ExportFullAccessPerTrustee,
                    $TranscriptFile,
                    $ResolveGroups,
                    $PowershellTempDir,
                    $ResultsForExportfile,
                    $RecipientsDoneForDisplay,
                    $AllAccessRights,
                    $All
                )



                filter get_member_recurse {
                    if ($_) {
                        try {
                            $tempObject = $null
                            $tempObject = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue } -ArgumentList $_.tostring()
                            if ($tempObject) {
                                if ($tempObject.RecipientType -ilike '*group') {
                                    Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-DistributionGroupMember -identity $($args[0]) -resultsize unlimited -ErrorAction silentlycontinue } -ArgumentList $tempobject.primarysmtpaddress | get_member_recurse
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

                function main {
                    $RecipientCount = $AllRecipients.count
                    $Count = $RecipientStartID + 1


                    for ($RecipientStartID; $RecipientStartID -le $RecipientEndID; $RecipientStartID++) {
                        Write-Host ("RecipientID: {0:00000}; '$($AllRecipients[$RecipientStartID].PrimarySmtpAddress)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@" -f $count)
                        $Recipient = $Mailbox = $null

                        $Recipient = $AllRecipients[$RecipientStartID]
                        if ($ExportFromOnPrem -eq $true) {
                            if ($Recipient.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                        } else {
                            if ($Recipient.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
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
                            $Mailbox = $AllMailboxes | Where-Object { $_.identity.distinguishedname -eq $Recipient.identity.distinguishedname }
                            $GrantorLegacyExchangeDN = $Mailbox.LegacyExchangeDN
                        } else {
                            $Mailbox = $null
                            $GrantorLegacyExchangeDN = $null
                        }
                        $Text = ("{0:000000}/{1:000000}, $($Recipient.PrimarySmtpAddress)" -f $count, $RecipientCount)

                        if (($Recipient.Recipienttype -eq 'PublicFolder') -or ($Recipient.Recipienttype -eq 'MailContact')) { $Text += (", #white:red#recipient type $GrantorRecipientType not supported#"); continue }


                        # Access Rights (full access etc.)
                        if ($ExportAccessRights -eq $true) {
                            Write-Host "  Access rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                            $Text += ', AccessRights'
                            if ($GrantorRecipientType -inotlike '*group') {
                                try {
                                    $TrusteeRights = $null
                                    if ($ExportFromOnPrem -eq $false) {
                                        #    $TrusteeRights = invoke-command -Session $script:ExchangeSession -ScriptBlock {Get-RecipientPermission $($args[0]) -ResultSize unlimited -erroraction stop -WarningAction SilentlyContinue} -ArgumentList $Recipient.primarysmtpaddress | Where-Object { ($_.IsInherited -eq $false) -and ($_.trustee -inotlike 'NT AUTHORITY\*') }
                                    } else {
                                        $TrusteeRights = $AllAccessrights | Where-Object { ($_.identity.distinguishedname -eq $recipient.distinguishedname) -and ($_.IsInherited -eq $false) -and ($_.user.rawidentity -inotlike 'NT AUTHORITY\*') } | Select-Object *, @{ name = 'trustee'; Expression = { $_.user.rawidentity } }

                                        #$TrusteeRights = invoke-command -Session $script:ExchangeSession -ScriptBlock {Get-MailboxPermission $($args[0]) -ResultSize unlimited -erroraction stop -WarningAction SilentlyContinue} -ArgumentList $recipient.primarysmtpaddress | Where-Object { ($_.IsInherited -eq $false) -and ($_.user -inotlike 'NT AUTHORITY\*') } | Select-Object *, @{ name = 'trustee'; Expression = { $_.user } }
                                    }

                                    if ($TrusteeRights) {
                                        foreach ($TrusteeRight in $TrusteeRights) {
                                            $trustees = @()
                                            if ($ResolveGroups) {
                                                ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                                    $trustees += $_
                                                }
                                            } else {
                                                ($TrusteeRight.trustee) | ForEach-Object {
                                                    $temp = $null
                                                    foreach ($x in $AllRecipients) {
                                                        if ($x.UserfriendlyName -ieq $TrusteeRight.trustee) {
                                                            $temp = $x
                                                            break
                                                        }
                                                    }
                                                    if ($temp) {
                                                        $trustees += $temp
                                                    } else {
                                                        $trustees += $_.tostring()
                                                    }
                                                }
                                            }
                                            foreach ($Trustee in $Trustees) {
                                                if ($ExportFromOnPrem -eq $true) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                foreach ($AccessRight in ($TrusteeRight.AccessRights | ForEach-Object { $_ -split ', ' })) {
                                                    $temp = '"' + (
                                                        (
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            ($GrantorRecipientType + '/' + $GrantorRecipientTypeDetails),
                                                            $GrantorEnvironment,
                                                            $Trustee.PrimarySmtpAddress,
                                                            $Trustee.DisplayName,
                                                            $TrusteeRight.trustee.ToString(),
                                                            ($Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails),
                                                            $TrusteeEnvironment,
                                                            $AccessRight,
                                                            '',
                                                            $GrantorLegacyExchangeDN,
                                                            $GrantorOU,
                                                            $Trustee.OrganizationalUnit
                                                        ) -join '";"'
                                                    ) + '"'
                                                    Write-Host "    $($temp)"
                                                    [System.Threading.Monitor]::Enter($ResultsForExportfile.SyncRoot)
                                                    if ($temp -inotin $ResultsForExportfile) { $ResultsForExportfile.add($temp) }
                                                    [System.Threading.Monitor]::Exit($ResultsForExportfile.SyncRoot)
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
                                $TrusteeRights = $null
                                if ($ExportFromOnPrem -eq $false) {
                                    $TrusteeRights = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-RecipientPermission -identity $($args[0]) -ResultSize unlimited -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $Recipient.primarysmtpaddress | Where-Object { ($_.Trustee -inotlike 'NT AUTHORITY\*') -and ($_.AccessRights -contains 'SendAs') }
                                } else {
                                    $TrusteeRights = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-ADPermission -identity $($args[0]) -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $GrantorDN | Where-Object { ($_.user -inotlike 'NT AUTHORITY\*') -and ($_.ExtendedRights -contains 'Send-As') } | Select-Object *, @{ name = 'trustee'; Expression = { $_.identity } }
                                }

                                if ($TrusteeRights) {
                                    foreach ($TrusteeRight in $TrusteeRights) {
                                        $trustees = @()
                                        if ($ResolveGroups) {
                                            ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                                $trustees += $_
                                            }
                                        } else {
                                            ($TrusteeRight.trustee) | ForEach-Object {
                                                $temp = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } -ArgumentList $_
                                                if ($temp) {
                                                    $trustees += $temp
                                                } else {
                                                    $trustees += $_.tostring()
                                                }
                                            }
                                        }
                                        foreach ($Trustee in $Trustees) {
                                            if ($ExportFromOnPrem -eq $true) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            $temp = '"' + (
                                                (
                                                    $GrantorPrimarySMTP,
                                                    $GrantorDisplayName,
                                                    ($GrantorRecipientType + '/' + $GrantorRecipientTypeDetails),
                                                    $GrantorEnvironment,
                                                    $Trustee.PrimarySmtpAddress,
                                                    $Trustee.DisplayName,
                                                    $TrusteeRight.trustee.ToString(),
                                                    ($Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails),
                                                    $TrusteeEnvironment,
                                                    'SendAs',
                                                    '',
                                                    $GrantorLegacyExchangeDN,
                                                    $GrantorOU,
                                                    $Trustee.OrganizationalUnit
                                                ) -join '";"'
                                            ) + '"'
                                            Write-Host "    $($temp)"
                                            [System.Threading.Monitor]::Enter($ResultsForExportfile.SyncRoot)
                                            if ($temp -inotin $ResultsForExportfile) { $ResultsForExportfile.add($temp) }
                                            [System.Threading.Monitor]::Exit($ResultsForExportfile.SyncRoot)
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
                                $TrusteeRights = $null

                                if ($ExportFromOnPrem -eq $false) {
                                    if (($GrantorRecipientType -ilike '*group') -and ($GrantorRecipientType -ine 'DynamicDistributionGroup') -and ($GrantorRecipientTypeDetails -ine 'groupmailbox')) {
                                        $TrusteeRights = (Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-distributiongroup $($args[0]) -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $recipient.primarysmtpaddress).GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                    } else {
                                        if ($GrantorRecipientTypeDetails -ieq 'groupmailbox') {
                                            $TrusteeRights = (Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-UnifiedGroup $($args[0]) -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $recipient.primarysmtpaddress).GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                        } elseif (($GrantorRecipientType -ieq 'DynamicDistributionGroup')) {
                                            $TrusteeRights = $Recipient.GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                        } else {
                                            $TrusteeRights = $Mailbox.GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
                                        }
                                    }
                                } else {
                                    if (($GrantorRecipientType -ilike '*group') -and ($GrantorRecipientType -ine 'DynamicDistributionGroup')) {
                                        $TrusteeRights = (Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-distributiongroup $($args[0]) -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $recipient.primarysmtpaddress).GrantSendOnBehalfto | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }
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
                                        $trustees = @()
                                        if ($ResolveGroups) {
                                            ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                                $trustees += $_
                                            }
                                        } else {
                                            ($TrusteeRight.trustee) | ForEach-Object {
                                                $temp = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } -ArgumentList $_
                                                if ($temp) {
                                                    $trustees += $temp
                                                } else {
                                                    $trustees += $_.tostring()
                                                }
                                            }
                                        }
                                        foreach ($Trustee in $Trustees) {
                                            if ($ExportFromOnPrem -eq $true) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            $temp = '"' + (
                                                (
                                                    $GrantorPrimarySMTP,
                                                    $GrantorDisplayName,
                                                    ($GrantorRecipientType + '/' + $GrantorRecipientTypeDetails),
                                                    $GrantorEnvironment,
                                                    $Trustee.PrimarySmtpAddress,
                                                    $Trustee.DisplayName,
                                                    $TrusteeRight.trustee.ToString(),
                                                    ($Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails),
                                                    $TrusteeEnvironment,
                                                    'SendOnBehalf',
                                                    '',
                                                    $GrantorLegacyExchangeDN,
                                                    $GrantorOU,
                                                    $Trustee.OrganizationalUnit
                                                ) -join '";"'
                                            ) + '"'
                                            Write-Host "    $($temp)"
                                            [System.Threading.Monitor]::Enter($ResultsForExportfile.SyncRoot)
                                            if ($temp -inotin $ResultsForExportfile) { $ResultsForExportfile.add($temp) }
                                            [System.Threading.Monitor]::Exit($ResultsForExportfile.SyncRoot)
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
                                $TrusteeRights = $null
                                $TrusteeRights = $Recipient.Managedby | Select-Object *, @{ name = 'trustee'; Expression = { $_ } }

                                if ($TrusteeRights) {
                                    foreach ($TrusteeRight in $TrusteeRights) {
                                        $trustees = @()
                                        if ($ResolveGroups) {
                                            ($TrusteeRight.trustee | get_member_recurse) | ForEach-Object {
                                                $trustees += $_
                                            }
                                        } else {
                                            ($TrusteeRight.trustee) | ForEach-Object {
                                                $temp = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } -ArgumentList $_
                                                if ($temp) {
                                                    $trustees += $temp
                                                } else {
                                                    $trustees += $_.tostring()
                                                }
                                            }
                                        }
                                        foreach ($Trustee in $Trustees) {
                                            if ($ExportFromOnPrem -eq $true) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            $temp = '"' + (
                                                (
                                                    $GrantorPrimarySMTP,
                                                    $GrantorDisplayName,
                                                    ($GrantorRecipientType + '/' + $GrantorRecipientTypeDetails),
                                                    $GrantorEnvironment,
                                                    $Trustee.PrimarySmtpAddress,
                                                    $Trustee.DisplayName,
                                                    $TrusteeRight.trustee.ToString(),
                                                    ($Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails),
                                                    $TrusteeEnvironment,
                                                    'ManagedBy',
                                                    '',
                                                    $GrantorLegacyExchangeDN,
                                                    $GrantorOU,
                                                    $Trustee.OrganizationalUnit
                                                ) -join '";"'
                                            ) + '"'
                                            Write-Host "    $($temp)"
                                            [System.Threading.Monitor]::Enter($ResultsForExportfile.SyncRoot)
                                            if ($temp -inotin $ResultsForExportfile) { $ResultsForExportfile.add($temp) }
                                            [System.Threading.Monitor]::Exit($ResultsForExportfile.SyncRoot)
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
                        if (($ExportFolderPermissions -eq $true) -and ($GrantorRecipientType -iNotMatch 'group')) {
                            Write-Host "  Folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                            $Text += ', Folders'
                            try {
                                $Folders = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-MailboxFolderStatistics -identity $($args[0]) -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $recipient.primarysmtpaddress | ForEach-Object { $_.folderpath } | ForEach-Object { $_.replace('/', '\') }
                                $Folders = ($Folders += '\') | Sort-Object # '\' is the root folder of the mailbox
                                if ($error.count -eq 0) {
                                    ForEach ($Folder in $Folders) {
                                        Write-Host "    $Folder @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                        $FolderKey = $GrantorPrimarySMTP + ':' + $Folder
                                        $TrusteeRights = $null
                                        $TrusteeRights = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-MailboxFolderPermission -identity $($args[0]) -ErrorAction stop -WarningAction silentlycontinue } -ArgumentList $FolderKey | Where-Object { ($_.user.usertype -ine 'Default') -and ($_.user.usertype -ine 'Anonymous') -and ($_.user.displayname -ine $Recipient.DisplayName) }

                                        if ($TrusteeRights) {
                                            foreach ($TrusteeRight in $TrusteeRights) {
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
                                                            $temp = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } -ArgumentList $_
                                                            if ($temp) {
                                                                $trustees += $temp
                                                            } else {
                                                                $trustees += $_.tostring()
                                                            }
                                                        }
                                                    } else {
                                                        $TrusteeRight.user.recipientprincipal.alias | Where-Object { $_ } | ForEach-Object {
                                                            $temp = Invoke-Command -Session $script:ExchangeSession -ScriptBlock { Get-recipient -identity $($args[0]) -resultsize 1 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue } -ArgumentList $_
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
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                    } else {
                                                        if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                    }

                                                    foreach ($AccessRight in $TrusteeRight.AccessRights) {
                                                        $temp = '"' + (
                                                            (
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                ($GrantorRecipientType + '/' + $GrantorRecipientTypeDetails),
                                                                $GrantorEnvironment,
                                                                $Trustee.PrimarySmtpAddress,
                                                                $Trustee.DisplayName,
                                                                $TrusteeRight.user.displayname.ToString(),
                                                                ($Trustee.RecipientType + '/' + $Trustee.RecipientTypeDetails),
                                                                $TrusteeEnvironment,
                                                                $AccessRight,
                                                                $Folder,
                                                                $GrantorLegacyExchangeDN,
                                                                $GrantorOU,
                                                                $Trustee.OrganizationalUnit
                                                            ) -join '":"'
                                                        ) + '"'
                                                        Write-Host "      $($temp)"
                                                        [System.Threading.Monitor]::Enter($ResultsForExportfile.SyncRoot)
                                                        if ($temp -inotin $ResultsForExportfile) { $ResultsForExportfile.add($temp) }
                                                        [System.Threading.Monitor]::Exit($ResultsForExportfile.SyncRoot)
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            } catch {
                                $Text += ' #white:red#ERROR#'
                                '==============================' | Out-File $ErrorFile -Append -Force
                                ('{0:000000}/{1:000000}: {2}, Folders' -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | Out-File $ErrorFile -Append -Force
                                $error[0] | Out-File $ErrorFile -Append -Force
                            }
                        }

                        $count++
                        [System.Threading.Monitor]::Enter($RecipientsDoneForDisplay.SyncRoot)
                        $RecipientsDoneForDisplay.Add($text)
                        [System.Threading.Monitor]::Exit($RecipientsDoneForDisplay.SyncRoot)

                        [System.GC]::Collect() # garbage collection
                    }
                }


                try {
                    $env:tmp = $PowershellTempDir
                    $env:temp = $PowershellTempDir
                    $ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
                    $TranscriptFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TranscriptFile)
                    $null = (New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile))
                    $null = (New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile))
                    if (Test-Path $Errorfile) { (Remove-Item $ErrorFile -Force) }

                    Start-Transcript -Path $TranscriptFile -Force

                    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                    $DebugPreferenceOrig = $DebugPreference
                    $DebugPreference = 'Continue'
                    Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"
                    $DebugPreference = $DebugPreferenceOrig

                    Write-Host "Preparations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    Set-Location $PSScriptRoot

                    Write-Host "  Recipient start ID: $($RecipientStartID + 1)"
                    Write-Host "  Recipient end ID: $($RecipientEndID + 1)"
                    Write-Host "  Time: @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                    main
                } catch {
                    '==============================' | Out-File $ErrorFile -Append -Force
                    'Unexpected error. Exiting.' | Out-File $ErrorFile -Append -Force
                    ("RecipientStartID: $RecipientStartID", "RecipientEndID: $RecipientEndID", "Time: @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@") | ForEach-Object {
                        $_ | Out-File $ErrorFile -Append -Force

                    }
                    $error[0] | Out-File $ErrorFile -Append -Force
                    exit 1
                } finally {
                    Remove-PSSession $script:ExchangeSession

                    Write-Host "Done @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                    Stop-Transcript
                }
            }).AddArgument($RecipientStartID).AddArgument($RecipientEndID).AddArgument($JobErrorfile).AddArgument($AllRecipients).AddArgument($ExportFromOnPrem).AddArgument($ExportAccessRights).AddArgument($ExportSendAs).AddArgument($ExportSendOnBehalf).AddArgument($ExportManagedby).AddArgument($ExportFolderPermissions).AddArgument($ExportFullAccessPerTrustee).AddArgument($JobTranscriptFile).AddArgument($ResolveGroups).AddArgument($PowershellTempDir).AddArgument($ConnectionUriList).AddArgument($ResultsForExportfile).AddArgument($RecipientsDoneForDisplay).AddArgument($AllAccessRights)


        $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
        $Handle = $PowerShell.BeginInvoke($Object, $Object)
        $temp = '' | Select-Object PowerShell, Handle, Object, StartTime, Done, RecipientStartID, Name, ErrorFile, TranscriptFile
        $temp.PowerShell = $PowerShell
        $temp.Handle = $Handle
        $temp.Object = $Object
        $temp.StartTime = $null
        $temp.Done = $false
        $temp.ErrorFile = $JobErrorFile
        $temp.TranscriptFile = $JobTranscriptFile

        [void]$script:jobs.Add($Temp)
    }


    $ResultsForExportfileLastindex = -1
    $RecipientsDoneForDisplayLastindex = -1

    $ExportFileFilestream = New-Object IO.FileStream($ExportFile, 'OpenOrCreate', 'ReadWrite', 'none')
    $ExportFileFilestreamWriter = New-Object IO.StreamWriter($ExportFileFilestream)
    $ExportFileFilestreamWriter.WriteLine('"Grantor Primary SMTP";"Grantor Display Name";"Grantor Recipient Type";"Grantor Environment";"Trustee Primary SMTP";"Trustee Display Name";"Trustee Original Identity";"Trustee Recipient Type";"Trustee Environment";"Permission";"Folder";"Grantor LegacyExchangeDN";"Grantor OU";"Trustee OU"')


    while (($script:jobs | Where-Object { $_.done -eq $false }).count -ne 0) {
        $script:jobs | ForEach-Object {
            # Add starttime attribute to each job
            if (($null -eq $_.StartTime) -and ($_.Powershell.Streams.Debug[0].Message -ilike 'Start(Ticks) = *')) {
                $StartTicks = $_.powershell.Streams.Debug[0].Message -replace '[^0-9]'
                $_.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }


            # Clean up as soon as job is completed
            if ((($_.handle.IsCompleted -eq $true) -and ($_.Done -eq $false))) {
                # append temp error file and delete temp file
                if (Test-Path $_.ErrorFile) {
                    Get-Content $_.ErrorFile | Out-File $Errorfile -Append -Force
                    Remove-Item $_.ErrorFile -Force
                }

                # append temp transcript file and delete temp file
                if (Test-Path $_.TranscriptFile) {
                    Get-Content $_.TranscriptFile | Out-File $Transcriptfile -Append -Force
                    Remove-Item $_.TranscriptFile -Force
                }


                $_.Done = $true
                $_.PowerShell.EndInvoke($_.handle)
                $_.PowerShell.Dispose()
            }


            # Show new completed recipients and total progress
            [System.Threading.Monitor]::Enter($RecipientsDoneForDisplay.SyncRoot)
            if (($RecipientsDoneForDisplay.count - 1) -gt $RecipientsDoneForDisplayLastindex) {
                $RecipientsDoneForDisplay[($RecipientsDoneForDisplayLastindex + 1)..$($RecipientsDoneForDisplay.count - 1)] | ForEach-Object {
                    write-hostcolored $_
                    $RecipientsDoneForDisplayLastindex++
                    if (((($RecipientsDoneForDisplayLastindex + 1) % 10 -eq 0) -and ($RecipientsDoneForDisplayLastindex -gt 0)) -or ($RecipientsDoneForDisplayLastindex -ge ($RecipientCount - 1))) {
                        Write-Host ("  {0:000000}/{1:000000} recipients done, {2:000000}/{3:000000} batches done @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@" -f ($RecipientsDoneForDisplayLastindex + 1), $RecipientCount, ($script:jobs | Where-Object { $_.done -eq $true }).count, $Batches) -ForegroundColor Yellow
                    }
                }
            }
            [System.Threading.Monitor]::Exit($RecipientsDoneForDisplay.SyncRoot)


            # update export file
            [System.Threading.Monitor]::Enter($ResultsForExportfile.SyncRoot)
            if (($ResultsForExportfile.count - 1) -gt $ResultsForExportfileLastindex) {
                $ResultsForExportfile[$($ResultsForExportfileLastindex + 1)..$($ResultsForExportfile.count - 1)] | ForEach-Object {
                    $ExportFileFilestreamWriter.WriteLine($_)
                    $ResultsForExportfileLastindex++
                }
            }
            [System.Threading.Monitor]::Exit($ResultsForExportfile.SyncRoot)
        }
    }

    Write-Host ("  {0:000000}/{1:000000} recipients done, {2:000000}/{3:000000} batches done @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@" -f ($RecipientsDoneForDisplayLastindex + 1), $RecipientCount, $(($script:jobs | Where-Object { $_.done -eq $true }).count), $Batches) -ForegroundColor Yellow

    $ExportFileFilestreamWriter.dispose()
    $ExportFileFilestream.dispose()

    if (($ExportAccessRights -eq $true) -and ($ExportFullAccessPerTrustee -eq $true) -and (Test-Path $exportfile)) {
        Write-Host 'Creating full access permission files per trustee.'
        $AllowedChars = @('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')
        $PrimarySMTPAddressesToIgnore = @('xxx@domain.com', 'yyy@domain.com') #List of primary SMTP addresses to ignore (service account, for example). Wildcards are not allowed.
        $RecipientPermissions = Import-Csv $ExportFile -Delimiter ';'
        for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
            if (($RecipientPermissions[$x].'Permission' -ilike 'FullAccess') -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne $RecipientPermissions[$x].'Grantor Primary SMTP') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -inotin $PrimarySMTPAddressesToIgnore) -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -inotin $PrimarySMTPAddressesToIgnore)) {
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
        $RecipientPermissions = Import-Csv $ExportFile -Delimiter ';' | Select-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission', 'Folder', 'Grantor OU', 'Trustee OU' | Sort-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission', 'Folder', 'Grantor OU', 'Trustee OU'
        if ($RecipientPermissions) {
            $RecipientPermissions | Export-Csv $ExportFile -NoTypeInformation -Force -Delimiter ';'
        }
    }

    if (Test-Path $TempRecipientFile) { (Remove-Item $TempRecipientFile -Force) }

} catch {
    Write-Host 'Unexpected error. Exiting.'
    $error[0]
} finally {
    Write-Host

    if ($script:ExchangeSession) { Remove-PSSession $script:ExchangeSession }

    if ($ExportFileFilestreamWriter) { $ExportFileFilestreamWriter.dispose() }
    if ($ExportFileFilestream) { $ExportFileFilestream.dispose() }



    Write-Host 'Closing runspaces and runspacepool, please wait'
    if ($RunspacePool) {
        $Runspacepool.close()
        $Runspacepool.dispose()
    }
    #while (($script:jobs.Done | Where-Object { $_ -eq $false }).count -ne 0) {
    #    $script:jobs | where { $_.Done -ne $true } | foreach {
    #        $_.PowerShell.Stop()
    #        $_.Done = $true
    #    }
    #}
    Write-Host 'Done.'
    Stop-Transcript
    if (Test-Path $tempTranscriptFile) {
        Get-Content $tempTranscriptFile | Out-File $TranscriptFile -Append -Force
        Remove-Item $tempTranscriptFile -Force
    }

}
