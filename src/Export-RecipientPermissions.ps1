[CmdletBinding(PositionalBinding = $false)]


Param(
    # Export from On-Prem or from Exchange Online
    # $true for export from on-prem
    # $false for export from Exchange Online
    [boolean]$ExportFromOnPrem = $false,


    # Server URIs to connect to
    # For on-prem installations, list all Exchange Server Remote PowerShell URIs the script can use
    # For Exchange Online use 'https://outlook.office365.com/powershell-liveid/', or the URI specific to your cloud environment
    [uri[]]$ExchangeConnectionUriList = ('https://outlook.office365.com/powershell-liveid/'),


    # Credentials for Exchange connection
    # Username and password are stored as encrypted secure strings
    [string]$ExchangeCredentialUsernameFile = '.\Export-RecipientPermissions_CredentialUsername.txt',
    [string]$ExchangeCredentialPasswordFile = '.\Export-RecipientPermissions_CredentialPassword.txt',


    # Maximum Exchange, AD and local sessions/jobs running in parallel
    # Watch CPU and RAM usage, and your Exchange throttling policy
    [int]$ParallelJobsExchange = $ExchangeConnectionUriList.count * 3,
    [int]$ParallelJobsAD = 50,
    [int]$ParallelJobsLocal = 100,


    # Grantors to consider
    # Only checks recipients that match the filter criteria. Only reduces the number of grantors, not the number of trustees.
    # Attributes that can filtered:
    #   .DistinguishedName
    #   .RecipientType, .RecipientTypeDetails
    #   .DisplayName
    #   .PrimarySmtpAddress: .Local, .Domain, .Address
    #   .EmailAddresses: .PrefixString, .IsPrimaryAddress, .SmtpAddress, .ProxyAddressString
    #   On-prem only: .Identity: .tostring() (CN), .DomainId, .Parent (parent CN)
    # Set to $null or '' to define all recipients as grantors to consider
    [string]$GrantorFilter = $null, #" `$Recipient.primarysmtpaddress.domain -ieq 'example.com'" },


    # Permissions to report
    #
    # Mailbox Access Rights
    # Rights set on the mailbox itself, such as "FullAccess" and "ReadAccess"
    [boolean]$ExportMailboxAccessRights = $true,
    [boolean]$ExportMailboxAccessRightsSelf = $false, # Report mailbox access rights granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)
    [boolean]$ExportMailboxAccessRightsInherited = $false, # Report inherited mailbox access rights (only works on-prem)
    #
    # Mailbox Folder Permissions
    # This part of the report can take very long
    [boolean]$ExportMailboxFolderPermissions = $true,
    [boolean]$ExportMailboxFolderPermissionsAnonymous = $false, # Report mailbox folder permissions granted to the special "Anonymous" user ("Anonymous" in English, "Anonym" in German, etc.)
    [boolean]$ExportMailboxFolderPermissionsDefault = $false, # Report mailbox folder permissions granted to the special "Default" user ("Default" in English, "Standard" in German, etc.)
    [boolean]$ExportMailboxFolderPermissionsOwnerAtLocal = $false, # Exchange Online only. For group mailboxes, export permissions granted to the special "Owner@Local" user.
    [boolean]$ExportMailboxFolderPermissionsMemberAtLocal = $false, # Exchange Online only. For group mailboxes, export permissions granted to the special "Member@Local" user.
    #
    # Send As
    [boolean]$ExportSendAs = $true,
    [boolean]$ExportSendAsSelf = $false, # Report Send As right granted to the SID "S-1-5-10" ("NT AUTHORITY\SELF" in English, "NT-AUTORITÄT\SELBST in German, etc.)
    #
    # Send On Behalf
    [boolean]$ExportSendOnBehalf = $true,
    #
    # Managed By
    # Only for distribution groups, and not to be confused with the "Manager" attribute
    [boolean]$ExportManagedBy = $true,


    # Name (and path) of the permission report file
    [string]$ExportFile = '.\export\Export-RecipientPermissions_Result.csv',


    # Name (and path) of the debug log file
    # Set to $null or '' to disable debugging
    [string]$DebugFile = '.\export\Export-RecipientPermissions_Debug.txt',


    # Interval to update the job progress
    # Updates are based von recipients done, not on duration
    # Number must be 1 or higher, low numbers mean bigger debug files
    [int][ValidateRange(1, [int]::MaxValue)]$UpdateInterval = 100
)


#
# Do not change anything from here on.
#


$ConnectExchangeOnline = {
    [int]$Retrycount = 1
    while ($Stoploop -ne $true) {
        try {
            if (-not (Get-PSSession | Where-Object { ($_.name -like 'ExchangeSession') -and ($_.state -like 'opened') })) {
                $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $ExchangeCredential -Authentication Basic -AllowRedirection -Name 'ExchangeSession' -ErrorAction Stop
                $null = Invoke-Command -Session $ExchangeSession -ScriptBlock { Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop }
            }
            $Stoploop = $true
        } catch {
            if ($Retrycount -le 3) {
                Write-Host 'Could not connect to Exchange Online. Trying again in 70 seconds.'
                Start-Sleep -Seconds 70
                $Retrycount++
            } else {
                Write-Host 'Could not connect to Exchange Online after three retries. Exiting.'
                exit 1
                $Stoploop = $true
            }
        }
    }
}


try {
    Set-Location $PSScriptRoot

    $ExportFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportFile)
    $DebugFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DebugFile)
    $ExchangeCredentialUsernameFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExchangeCredentialUsernameFile)
    $ExchangeCredentialPasswordFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExchangeCredentialPasswordFile)

    if ($DebugFile) {
        try {
            Stop-Transcript
        } catch {
        }
        $null = Start-Transcript -Path $DebugFile -Force
    }

    Clear-Host
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    $tempConnectionUriQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new(10000))
    while ($tempConnectionUriQueue.count -le 10000) {
        foreach ($ExchangeConnectionUri in $ExchangeConnectionUriList) {
            $tempConnectionUriQueue.Enqueue($ExchangeConnectionUri.AbsoluteUri)
        }
    }

    if (Test-Path $ExportFile) {
        Remove-Item $ExportFile -Force
    }
    foreach ($RecipientFile in (Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))) {
        Remove-Item $Recipientfile -Force
    }

    if ($DebugFile) {
        foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
            Remove-Item $JobDebugFile -Force
        }
    }

    if (Test-Path $ExportFile) {
        Remove-Item $ExportFile -Force
    }

    # Credentials
    Write-Host
    Write-Host "Exchange credentials @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (-not ((Test-Path $ExchangeCredentialUsernameFile) -and (Test-Path $ExchangeCredentialPasswordFile))) {
        Write-Host '  No stored credential found'
        Write-Host '    Username and password are stored as encrypted secure strings'
        Read-Host -Prompt '    Please enter username for later use (characters are masked)' -AsSecureString | ConvertFrom-SecureString | Out-File $ExchangeCredentialUsernameFile -Force -Encoding utf8
        Read-Host -Prompt '    Please enter password for later use (characters are masked)' -AsSecureString | ConvertFrom-SecureString | Out-File $ExchangeCredentialPasswordFile -Force -Encoding utf8
    }

    Write-Host '  Loading credentials encrypted as secure strings'
    Write-Host "    Username file: '$ExchangeCredentialUsernameFile'"
    Write-Host "    Password file: '$ExchangeCredentialPasswordFile'"
    Write-Host '  To change username and/or password, delete one or all of the files mentioned above and run the script again'
    $ExchangeCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ([PSCredential]::new('X', (Get-Content $ExchangeCredentialUsernameFile -Encoding UTF8 | ConvertTo-SecureString)).GetNetworkCredential().Password), (Get-Content $ExchangeCredentialPasswordFile -Encoding UTF8 | ConvertTo-SecureString)


    # Connect to Exchange
    Write-Host
    Write-Host "Connect to Exchange for single-thread operations @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    $connectionUri = $tempConnectionUriQueue.dequeue()
    Write-Host "  Single-thread operation, use connection '$($connectionUri)'"

    if ($ExportFromOnPrem -eq $true) {
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $ExchangeCredential -Authentication Kerberos
        Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True }
    } else {
        . ([scriptblock]::Create($ConnectExchangeOnline))
    }


    # Import recipients
    Write-Host
    Write-Host "Import recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "  Single-thread operation, use connection '$($connectionUri)'"

    $AllRecipients = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new(1000000))
    $AllRecipients.AddRange(@((Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Get-Recipient -resultsize unlimited -WarningAction silentlycontinue | Select-Object -Property identity, distinguishedname, recipienttype, recipienttypedetails, displayname, primarysmtpaddress, emailaddresses, managedby, 'UserFriendlyName' }) | Sort-Object { $_.primarysmtpaddress.address }))
    $AllRecipients.TrimToSize()
    Write-Host ('  {0:0000000} recipients found' -f $($AllRecipients.count))


    # Import recipient permissions (SendAs)
    Write-Host
    Write-Host "Import Send As permissions from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (($ExportFromOnPrem -eq $false) -and ($ExportSendAs -eq $true)) {
        Write-Host "  Single-thread operation, use connection '$($connectionUri)'"
        $AllRecipientsSendas = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count * 2))
        $AllRecipientsSendas.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-recipientpermission -resultsize unlimited -WarningAction silentlycontinue | Select-Object identity, trustee, trusteesidstring, accessrights, accesscontroltype, isinherited, inheritancetype }))
        $AllRecipientsSendas.TrimToSize()
        Write-Host ('  {0:0000000} Send As permissions found' -f $($AllRecipientsSendas.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Import Send On Behalf from cloud
    Write-Host
    Write-Host "Import Send On Behalf permissions from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (($ExportFromOnPrem -eq $false) -and ($ExportSendOnBehalf -eq $true)) {
        Write-Host "  Single-thread operation, use connection '$($connectionUri)'"
        $AllRecipientsSendonbehalf = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count * 2))
        Write-Host "  Mailboxes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailbox -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto }))
        Write-Host "  Distribution groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-distributiongroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto }))
        Write-Host "  Dynamic Distribution Groups @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-dynamicdistributiongroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto }))
        Write-Host "  Unified Groups (Microsoft 365 Groups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $AllRecipientsSendonbehalf.AddRange(@(Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-unifiedgroup -filter 'GrantSendOnBehalfTo -ne $null' -resultsize unlimited -WarningAction silentlycontinue | Select-Object identity, grantsendonbehalfto }))
        $AllRecipientsSendonbehalf.TrimToSize()
        Write-Host ('  {0:0000000} recipients with Send On Behalf permissions found' -f $($AllRecipientsSendonbehalf.count))
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Disconnect from Exchange
    Write-Host
    Write-Host "Single-thread operations completed, remove connection to Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Remove-PSSession $ExchangeSession


    # Create lookup hashtables for GUID, DistinguishedName and PrimarySmtpAddress
    Write-Host
    Write-Host "Create lookup hashtables for GUID, DistinguishedName and PrimarySmtpAddress @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  DistinguishedName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsDnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $AllRecipientsDnToIndex[$(($AllRecipients[$x]).distinguishedname)] = $x
    }

    Write-Host "  GUID to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsGuidToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $AllRecipientsGuidToIndex[$(($AllRecipients[$x]).identity.objectguid.guid)] = $x
    }

    Write-Host "  PrimarySmtpAddress to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsSmtpToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        if (($AllRecipients[$x]).primarysmtpaddress.address) {
            $AllRecipientsSmtpToIndex[$(($AllRecipients[$x]).primarysmtpaddress.address)] = $x
        }
    }

    # UserFriendlyNames
    Write-Host
    Write-Host "Find each recipient's UserFriendlyNames @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxAccessRights -or $ExportSendAs) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        for ($RecipientID = 0; $RecipientID -lt $AllRecipients.count; $RecipientID++) {
            $tempqueue.enqueue($RecipientID)
        }

        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min([math]::ceiling($tempQueueCount / 100), $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param (
                            $AllRecipients,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $AllRecipientsGuidToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchangeOnline,
                            $ExchangeCredential,
                            $ScriptPath
                        )

                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Create connection between UserFriendlyNames and recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $connectionUri = $tempConnectionUriQueue.dequeue()
                            if ($ExportFromOnPrem) {
                                $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $ExchangeCredential -Authentication Kerberos
                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True }
                            } else {
                                . ([scriptblock]::Create($ConnectExchangeOnline))
                            }

                            while ($tempQueue.count -gt 0) {
                                Write-Host "Filter string @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                $dequeued = 0
                                $filterstring = ''

                                while (($dequeued -lt 100) -and ($tempQueue.count -gt 0)) {
                                    $x = $tempQueue.dequeue()
                                    $filterstring += "(guid -eq '$($AllRecipients[$x].identity.objectguid.guid)') -or "
                                    $dequeued++
                                }
                                $filterstring = $filterstring.trimend(' -or ')

                                Write-Host "  $filterstring"

                                foreach ($securityprincipal in @(Invoke-Command -Session $ExchangeSession -ScriptBlock { get-securityprincipal -filter "$($args[0])" -resultsize unlimited -WarningAction silentlycontinue | Select-Object userfriendlyname, guid } -ArgumentList $filterstring)) {
                                    Write-Host "  '$($securityprincipal.guid.guid)' = '$($securityprincipal.UserFriendlyName)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                    ($AllRecipients[$($AllRecipientsGuidToIndex[$($securityprincipal.guid.guid)])]).UserFriendlyName = $securityprincipal.UserFriendlyName
                                }
                            }
                        } catch {
                            $_ | Out-String | Write-Host
                        } finally {
                            Remove-PSSession $ExchangeSession
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients            = $AllRecipients
                        tempConnectionUriQueue   = $tempConnectionUriQueue
                        tempQueue                = $tempQueue
                        AllRecipientsGuidToIndex = $AllRecipientsGuidToIndex
                        DebugFile                = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem         = $ExportFromOnPrem
                        ExchangeCredential       = $ExchangeCredential
                        ScriptPath               = $PSScriptRoot
                        ConnectExchangeOnline    = $ConnectExchangeOnline
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host "  $($tempQueueCount) recipients to check. Done (in steps of $($UpdateInterval)):"

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - (($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count * 100))
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }
            Write-Host

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all recipients have been checked. Enable debugging option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Create lookup hashtable for UserFriendlyName
    Write-Host
    Write-Host "Create lookup hashtable for UserFriendlyName @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  UserFriendlyName to recipients array index @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllRecipientsUfnToIndex = [system.collections.hashtable]::Synchronized([system.collections.hashtable]::new($AllRecipients.count, [StringComparer]::OrdinalIgnoreCase))
    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $Recipient = $AllRecipients[$x]
        if ($Recipient.userfriendlyname) {
            $AllRecipientsUfnToIndex[$Recipient.userfriendlyname] = $x
        }
    }


    # Grantors
    Write-Host
    Write-Host "Define grantors by filtering recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "  Filter: { $($GrantorFilter) }"
    $GrantorsToConsider = [system.collections.arraylist]::Synchronized([system.collections.arraylist]::new($AllRecipients.count))

    for ($x = 0; $x -lt $AllRecipients.count; $x++) {
        $Recipient = $AllRecipients[$x]

        if (-not $GrantorFilter) {
            $null = $GrantorsToConsider.add($x)
        } else {
            if ((. ([scriptblock]::Create($GrantorFilter)))) {
                $null = $GrantorsToConsider.add($x)
            }
        }
    }
    $GrantorsToConsider.TrimToSize()
    Write-Host ('  {0:0000000}/{1:0000000} recipients are considered as grantors' -f $($GrantorsToConsider.count), $($AllRecipients.count))


    # Mailbox Access Rights
    Write-Host
    Write-Host "Mailbox Access Rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxAccessRights) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.recipienttypedetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $ExportMailboxAccessRightsSelf,
                            $ExportMailboxAccessRightsInherited,
                            $ExportFile,
                            $AllRecipientsUfnToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchangeOnline,
                            $ExchangeCredential,
                            $ScriptPath
                        )

                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Mailbox access rights @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $connectionUri = $tempConnectionUriQueue.dequeue()
                            if ($ExportFromOnPrem) {
                                $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $ExchangeCredential -Authentication Kerberos
                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True }
                            } else {
                                . ([scriptblock]::Create($ConnectExchangeOnline))
                            }

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileResult.clear()
                                $RecipientID = $tempQueue.dequeue()

                                $Grantor = $AllRecipients[$RecipientID]
                                $Trustee = $null

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                if (($GrantorRecipienttype -ieq 'PublicFolder') -or ($GrantorRecipienttype -ieq 'MailContact')) {
                                    continue
                                }

                                foreach ($MailboxPermission in
                                    @($(
                                            if ($ExportFromOnPrem) {
                                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $args[0] -resultsize unlimited -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                            } else {
                                                if ($GrantorRecipientTypeDetails -ine 'groupmailbox') {
                                                    Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxpermission -identity $args[0] -resultsize unlimited -WarningAction silentlycontinue | Select-Object -Property identity, user, accessrights, deny, isinherited, inheritanceType } -ArgumentList $GrantorPrimarySMTP | Select-Object identity, user, accessrights, deny, isinherited, inheritanceType
                                                }
                                            }
                                        ))
                                ) {
                                    foreach ($TrusteeRight in @($MailboxPermission | Where-Object { if ($ExportMailboxAccessRightsSelf) { $true } else { $_.user.SecurityIdentifier -ine 'S-1-5-10' } } | Where-Object { if ($ExportMailboxAccessRightsInherited) { $true } else { $_.IsInherited -ne $true } } | Select-Object *, @{ name = 'trustee'; Expression = { $_.user.rawidentity } })) {
                                        $trustees = [system.collections.arraylist]::new(1000)
                                        $index = $AllRecipientsUfnToIndex[$($TrusteeRight.trustee)]
                                        if ($index -ge 0) {
                                            $trustees.add($AllRecipients[$index])
                                        } else {
                                            $trustees.add($TrusteeRight.trustee)
                                        }
                                        foreach ($Trustee in $Trustees) {
                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }
                                            foreach ($Accessright in ($TrusteeRight.Accessrights -split ', ')) {
                                                $ExportFileresult.add((('"' + (
                                                                (
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $Accessright,
                                                                    $(if ($Trusteeright.deny) { 'Deny' } else { 'Allow' }),
                                                                    $Trusteeright.IsInherited,
                                                                    $Trusteeright.InheritanceType,
                                                                    $TrusteeRight.trustee,
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    ("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) -join '";"'
                                                            ) + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                            }
                                        }
                                    }
                                }
                                $ExportFileResult | Sort-Object -Unique | Out-File ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Append -Encoding utf8 -Force
                            }
                        } catch {
                            $_ | Out-String | Write-Host
                        } finally {
                            Remove-PSSession $ExchangeSession

                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                      = $AllRecipients
                        tempConnectionUriQueue             = $tempConnectionUriQueue
                        tempQueue                          = $tempQueue
                        ExportMailboxAccessRightsSelf      = $ExportMailboxAccessRightsSelf
                        ExportMailboxAccessRightsInherited = $ExportMailboxAccessRightsInherited
                        ExportFile                         = $ExportFile
                        AllRecipientsUfnToIndex            = $AllRecipientsUfnToIndex
                        DebugFile                          = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                   = $ExportFromOnPrem
                        ExchangeCredential                 = $ExchangeCredential
                        ScriptPath                         = $PSScriptRoot
                        ConnectExchangeOnline              = $ConnectExchangeOnline
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host "  $($tempQueueCount) grantor mailboxes to check. Done (in steps of $($UpdateInterval)):"

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }
            Write-Host

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all grantor mailboxes have been checked. Enable debugging option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Mailbox Folder Permissions
    Write-Host
    Write-Host "Mailbox Folder Permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportMailboxFolderPermissions) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))
        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            $Recipient = $AllRecipients[$x]

            if (($Recipient.recipienttypedetails -ilike '*mailbox') -and ($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsExchange)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel Exchange jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempConnectionUriQueue,
                            $tempQueue,
                            $ExportMailboxFolderPermissions,
                            $ExportMailboxFolderPermissionsAnonymous,
                            $ExportMailboxFolderPermissionsDefault,
                            $ExportMailboxFolderPermissionsOwnerAtLocal,
                            $ExportMailboxFolderPermissionsMemberAtLocal,
                            $ExportFile,
                            $AllRecipientsGuidToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ConnectExchangeOnline,
                            $ExchangeCredential,
                            $ScriptPath
                        )
                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Mailbox folder permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $connectionUri = $tempConnectionUriQueue.dequeue()
                            if ($ExportFromOnPrem) {
                                $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $ExchangeCredential -Authentication Kerberos
                                Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { Set-AdServerSettings -ViewEntireForest $True }
                            } else {
                                . ([scriptblock]::Create($ConnectExchangeOnline))
                            }

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileResult.Clear()
                                $RecipientID = $tempQueue.dequeue()

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                if ($ExportFromOnPrem) {
                                    $Folders = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderstatistics -identity $args[0] | Select-Object folderid, folderpath } -ArgumentList $GrantorPrimarySMTP
                                } else {
                                    $Folders = Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderstatistics -identity $args[0] | Select-Object folderid, folderpath } -ArgumentList $GrantorPrimarySMTP
                                }
                                $Folders[0].folderpath = '/'

                                foreach ($Folder in $Folders) {
                                    Write-Host "  Folder '$($folder.folderpath)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                    foreach ($FolderPermissions in
                                        @($(
                                                if ($ExportFromOnPrem) {
                                                    (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)")
                                                } else {
                                                    if ($GrantorRecipientTypeDetails -ieq 'groupmailbox') {
                                                        (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -groupmailbox -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)")
                                                    } else {
                                                        (Invoke-Command -Session $ExchangeSession -HideComputerName -ScriptBlock { get-mailboxfolderpermission -identity $args[0] -WarningAction silentlycontinue | Select-Object identity, user, accessrights } -ArgumentList "$($GrantorPrimarySMTP):$($Folder.folderid)")
                                                    }
                                                }
                                            ))
                                    ) {
                                        foreach ($FolderPermission in $FolderPermissions) {
                                            foreach ($AccessRight in ($FolderPermission.AccessRights)) {
                                                if ($ExportMailboxFolderPermissionsDefault -eq $false) {
                                                    if ($FolderPermission.user.usertype.value -ieq 'default') { continue }
                                                }

                                                if ($ExportMailboxFolderPermissionsAnonymous -eq $false) {
                                                    if ($FolderPermission.user.usertype.value -ieq 'anonymous') { continue }
                                                }

                                                if ($ExportFromOnPrem) {
                                                    $ExportFileResult.Add((('"' + ((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        $($Folder.Folderpath),
                                                                        $($Accessright),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $($FolderPermission.user.displayname),
                                                                        $($FolderPermission.user.adrecipient.primarysmtpaddress),
                                                                        $($FolderPermission.user.adrecipient.displayname),
                                                                        ("$($FolderPermission.user.adrecipient.recipienttype)/$($FolderPermission.user.adrecipient.recipienttypedetails)" -replace '^/$', ''),
                                                                        $(if ($FolderPermission.user.adrecipient.recipienttypedetails -ilike 'Remote*') { 'Cloud' } else { 'On-Prem' })
                                                                    ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                } else {
                                                    if ($ExportMailboxFolderPermissionsOwnerAtLocal -eq $false) {
                                                        if ($FolderPermission.user.recipientprincipal.primarysmtpaddress -ieq 'owner@local') { continue }
                                                    }

                                                    if ($ExportMailboxFolderPermissionsMemberAtLocal -eq $false) {
                                                        if ($FolderPermission.user.recipientprincipal.primarysmtpaddress -ieq 'member@local') { continue }
                                                    }
                                                    $ExportFileResult.Add((('"' + ((
                                                                        $GrantorPrimarySMTP,
                                                                        $GrantorDisplayName,
                                                                        ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                        $GrantorEnvironment,
                                                                        $($Folder.Folderpath),
                                                                        $($Accessright),
                                                                        'Allow',
                                                                        'False',
                                                                        'None',
                                                                        $($FolderPermission.user.displayname),
                                                                        $($FolderPermission.user.recipientprincipal.primarysmtpaddress),
                                                                        $($FolderPermission.user.recipientprincipal.displayname),
                                                                        ("$($FolderPermission.user.recipientprincipal.recipienttype)/$($FolderPermission.user.recipientprincipal.recipienttypedetails)" -replace '^/$', ''),
                                                                        $(if ($FolderPermission.user.recipientprincipal.recipienttypedetails -ilike 'Remote*') { 'On-prem' } else { 'Cloud' })
                                                                    ) -join '";"') + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                                }
                                            }
                                        }
                                    }
                                }
                                $ExportFileResult | Sort-Object -Unique | Out-File ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Append -Encoding utf8 -Force
                            }
                        } catch {
                            $_ | Out-String | Write-Host

                        } finally {
                            Remove-PSSession $ExchangeSession
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients                               = $AllRecipients
                        tempConnectionUriQueue                      = $tempConnectionUriQueue
                        tempQueue                                   = $tempQueue
                        ExportMailboxFolderPermissions              = $ExportMailboxFolderPermissions
                        ExportMailboxFolderPermissionsAnonymous     = $ExportMailboxFolderPermissionsAnonymous
                        ExportMailboxFolderPermissionsDefault       = $ExportMailboxFolderPermissionsDefault
                        ExportMailboxFolderPermissionsOwnerAtLocal  = $ExportMailboxFolderPermissionsOwnerAtLocal
                        ExportMailboxFolderPermissionsMemberAtLocal = $ExportMailboxFolderPermissionsMemberAtLocal
                        ExportFile                                  = $ExportFile
                        AllRecipientsGuidToIndex                    = $AllRecipientsGuidToIndex
                        DebugFile                                   = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem                            = $ExportFromOnPrem
                        ExchangeCredential                          = $ExchangeCredential
                        ScriptPath                                  = $PSScriptRoot
                        ConnectExchangeOnline                       = $ConnectExchangeOnline
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host "  $($tempQueueCount) grantor mailboxes to check. Done (in steps of $($UpdateInterval)):"

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }
            Write-Host

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all grantor mailboxes have been checked. Enable debugging option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            [GC]::Collect(); Start-Sleep 1
        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Send As
    Write-Host
    Write-Host "Send As @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportSendAs) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            if ($x -in $GrantorsToConsider) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        if ($ExportFromOnPrem) {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsAD)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel AD jobs"
        } else {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"
        }
        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $AllRecipientsUfnToIndex,
                            $AllRecipientsSmtpToIndex,
                            $ExportSendAsSelf,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $AllRecipientsSendas
                        )
                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Send As @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $result = [system.collections.arraylist]::new(1000)
                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $result.clear()
                                $ExportFileresult.clear()
                                $RecipientID = $tempQueue.dequeue()

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                if ($ExportFromOnPrem) {
                                    foreach ($entry in (([adsi]"LDAP://<GUID=$($Grantor.identity.objectguid.guid)>").ObjectSecurity.Access)) {
                                        $trustee = $null
                                        if ($entry.ObjectType -eq 'ab721a54-1e2f-11d0-9819-00aa0040529b') {
                                            if (($entry.identityreference -ilike '*\*') -and ($ExportSendAsSelf -eq $false)) {
                                                if ((([System.Security.Principal.NTAccount]::new($entry.identityreference)).Translate([System.Security.Principal.SecurityIdentifier])).value -ieq 'S-1-5-10') {
                                                    continue
                                                }
                                            }

                                            $index = $AllRecipientsUfnToIndex[$($entry.identityreference.tostring())]
                                            if ($index -ge 0) {
                                                $trustee = $AllRecipients[$index]
                                            } else {
                                                $trustee = $entry.identityreference
                                            }

                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            $ExportFileresult.add((('"' + (
                                                            (
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                '',
                                                                'SendAs',
                                                                $entry.AccessControlType,
                                                                $entry.IsInherited,
                                                                $entry.InheritanceType,
                                                                $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                $Trustee.PrimarySmtpAddress.address,
                                                                $Trustee.DisplayName,
                                                                ("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) -join '";"'
                                                        ) + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                        }
                                    }
                                } else {
                                    foreach ($entry in $AllRecipientsSendas) {
                                        if ($entry.identity.objectguid.guid -eq $Grantor.identity.objectguid.guid) {
                                            if (($ExportSendAsSelf -eq $false) -and ($entry.trusteesidstring -ieq 'S-1-5-10')) {
                                                continue
                                            }
                                            $trustee = $null
                                            if ($entry.trustee -ilike '*\*') {
                                                $index = $AllRecipientsUfnToIndex[$($entry.trustee)]
                                            } elseif ($entry.trustee -ilike '*@*') {
                                                $index = $AllRecipientsSmtpToIndex[$($entry.trustee)]
                                            }
                                            if ($index -ge 0) {
                                                $trustee = $AllRecipients[$index]
                                            } else {
                                                $trustee = $entry.trustee
                                            }

                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            foreach ($AccessRight in $entry.AccessRights) {
                                                $ExportFileresult.add((('"' + (
                                                                (
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    $AccessRight,
                                                                    $entry.AccessControlType,
                                                                    $entry.IsInherited,
                                                                    $entry.InheritanceType,
                                                                    $(($Trustee.displayname, $entry.trustee) | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    ("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) -join '";"'
                                                            ) + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                            }
                                        }
                                    }
                                }
                            }
                            $ExportFileResult | Sort-Object -Unique | Out-File ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Append -Encoding utf8 -Force
                        } catch {
                            $_ | Out-String | Write-Host
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients            = $AllRecipients
                        tempQueue                = $tempQueue
                        ExportFile               = $ExportFile
                        AllRecipientsUfnToIndex  = $AllRecipientsUfnToIndex
                        AllRecipientsSmtpToIndex = $AllRecipientsSmtpToIndex
                        ExportSendAsSelf         = $ExportSendAsSelf
                        DebugFile                = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem         = $ExportFromOnPrem
                        ScriptPath               = $PSScriptRoot
                        AllRecipientsSendas      = $AllRecipientsSendas
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host "  $($tempQueueCount) grantors to check. Done (in steps of $($UpdateInterval)):"

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }
            Write-Host

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all grantors have been checked. Enable debugging option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Send On Behalf
    Write-Host
    Write-Host "Send On Behalf @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportSendOnBehalf) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        for ($x = 0; $x -lt $AllRecipients.count; $x++) {
            if (($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        if ($ExportFromOnPrem) {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsAD)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel AD jobs"
        } else {
            $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)
            Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"
        }

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $AllRecipientsDnToIndex,
                            $AllRecipientsGuidToIndex,
                            $DebugFile,
                            $ExportFromOnPrem,
                            $ScriptPath,
                            $AllRecipientsSendonbehalf
                        )

                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }

                            Write-Host "Send On Behalf @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)
                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()
                                $RecipientID = $tempQueue.dequeue()

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                if ($ExportFromOnPrem) {
                                    $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher("(objectguid=$([System.String]::Join('', (([guid]$($Grantor.identity.objectguid.guid)).ToByteArray() | ForEach-Object { '\' + $_.ToString('x2') })).ToUpper()))")
                                    $directorySearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($Grantor.identity.domainid)")
                                    $directorySearcher.PropertiesToLoad.Add('publicDelegates')
                                    $directorySearcherResults = $directorySearcher.FindOne()

                                    foreach ($directorySearcherResult in $directorySearcherResults) {
                                        foreach ($delegateBindDN in $directorySearcherResult.properties.publicdelegates) {
                                            $index = $AllRecipientsDnToIndex[$delegateBindDN]
                                            if ($index -ge 0) {
                                                $trustee = $AllRecipients[$index]
                                            } else {
                                                $trustee = $delegateBindDN
                                            }

                                            if ($ExportFromOnPrem) {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                            } else {
                                                if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                            }

                                            $ExportFileresult.add((('"' + (
                                                            (
                                                                $GrantorPrimarySMTP,
                                                                $GrantorDisplayName,
                                                                ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                $GrantorEnvironment,
                                                                '',
                                                                'SendOnBehalf',
                                                                'Allow',
                                                                'False',
                                                                'None',
                                                                $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                $Trustee.PrimarySmtpAddress.address,
                                                                $Trustee.DisplayName,
                                                                ("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                $TrusteeEnvironment
                                                            ) -join '";"'
                                                        ) + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                        }
                                    }
                                } else {
                                    foreach ($entry in $AllRecipientsSendonbehalf) {
                                        if ($entry.identity.objectguid.guid -eq $Grantor.identity.objectguid.guid) {
                                            $trustee = $null
                                            foreach ($AccessRight in $entry.GrantSendOnBehalfTo) {
                                                $index = $AllRecipientsGuidToIndex[$($AccessRight.objectguid.guid)]
                                                if ($index -ge 0) {
                                                    $trustee = $AllRecipients[$index]
                                                } else {
                                                    $trustee = $AccessRight.tostring()
                                                }

                                                if ($ExportFromOnPrem) {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                                } else {
                                                    if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                                }

                                                $ExportFileresult.add((('"' + (
                                                                (
                                                                    $GrantorPrimarySMTP,
                                                                    $GrantorDisplayName,
                                                                    ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                                    $GrantorEnvironment,
                                                                    '',
                                                                    'SendOnBehalf',
                                                                    'Allow',
                                                                    'False',
                                                                    'None',
                                                                    $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                                    $Trustee.PrimarySmtpAddress.address,
                                                                    $Trustee.DisplayName,
                                                                    ("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                                    $TrusteeEnvironment
                                                                ) -join '";"'
                                                            ) + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                            }
                                        }
                                    }
                                }
                                $ExportFileResult | Sort-Object -Unique | Out-File ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Append -Encoding utf8 -Force
                            }
                        } catch {
                            $_ | Out-String | Write-Host
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients             = $AllRecipients
                        tempQueue                 = $tempQueue
                        ExportFile                = $ExportFile
                        AllRecipientsDnToIndex    = $AllRecipientsDnToIndex
                        AllRecipientsGuidToIndex  = $AllRecipientsGuidToIndex
                        DebugFile                 = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ExportFromOnPrem          = $ExportFromOnPrem
                        ScriptPath                = $PSScriptRoot
                        AllRecipientsSendonbehalf = $AllRecipientsSendonbehalf
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host "  $($tempQueueCount) grantors to check. Done (in steps of $($UpdateInterval)):"

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }
            Write-Host

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all grantors have been checked. Enable debugging option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Managed By
    Write-Host
    Write-Host "Managed By @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($ExportManagedBy) {
        $tempQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new($AllRecipients.count))

        foreach ($x in (0..($AllRecipients.count - 1))) {
            if (($x -in $GrantorsToConsider)) {
                $tempQueue.enqueue($x)
            }
        }
        $tempQueueCount = $tempQueue.count

        $ParallelJobsNeeded = [math]::min($tempQueueCount, $ParallelJobsLocal)

        Write-Host "  Multi-thread operation, create $($ParallelJobsNeeded) parallel local jobs"

        if ($ParallelJobsNeeded -ge 1) {
            $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $ParallelJobsNeeded)
            $RunspacePool.Open()

            $runspaces = [system.collections.arraylist]::new($ParallelJobsNeeded)

            1..$ParallelJobsNeeded | ForEach-Object {
                $Powershell = [powershell]::Create()
                $Powershell.RunspacePool = $RunspacePool

                [void]$Powershell.AddScript(
                    {
                        param(
                            $AllRecipients,
                            $tempQueue,
                            $ExportFile,
                            $AllRecipientsGuidToIndex,
                            $DebugFile,
                            $ScriptPath
                        )
                        try {
                            Set-Location $ScriptPath

                            if ($DebugFile) {
                                $null = Start-Transcript -Path $DebugFile -Force
                            }
                            Write-Host "Managed By @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                            $ExportFileResult = [system.collections.arraylist]::new(1000)

                            while ($tempQueue.count -gt 0) {
                                $ExportFileresult.clear()
                                $RecipientID = $tempQueue.dequeue()

                                $Grantor = $AllRecipients[$RecipientID]

                                $GrantorDisplayName = $Grantor.DisplayName
                                $GrantorPrimarySMTP = $Grantor.PrimarySMTPAddress.address
                                $GrantorRecipientType = $Grantor.RecipientType.value
                                $GrantorRecipientTypeDetails = $Grantor.RecipientTypeDetails.value
                                if ($ExportFromOnPrem) {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'Cloud' } else { $GrantorEnvironment = 'On-Prem' }
                                } else {
                                    if ($Grantor.RecipientTypeDetails -ilike 'Remote*') { $GrantorEnvironment = 'On-Prem' } else { $GrantorEnvironment = 'Cloud' }
                                }

                                Write-Host "$($GrantorPrimarySMTP), $($GrantorRecipientType)/$($GrantorRecipientTypeDetails) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

                                foreach ($TrusteeRight in $Grantor.ManagedBy) {
                                    $trustees = [system.collections.arraylist]::new(1000)
                                    $index = $AllRecipientsGuidToIndex[$($TrusteeRight.objectguid.guid)]
                                    if ($index -ge 0) {
                                        $trustees.add($AllRecipients[$index])
                                    } else {
                                        $trustees.add((($TrusteeRight.distinguishedname, $TrusteeRight.identity.objectguid.guid) | Select-Object -First 1))
                                    }

                                    foreach ($Trustee in $Trustees) {
                                        if ($ExportFromOnPrem) {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'Cloud' } else { $TrusteeEnvironment = 'On-Prem' }
                                        } else {
                                            if ($Trustee.RecipientTypeDetails -ilike 'Remote*') { $TrusteeEnvironment = 'On-Prem' } else { $TrusteeEnvironment = 'Cloud' }
                                        }
                                        $ExportFileresult.add((('"' + (
                                                        (
                                                            $GrantorPrimarySMTP,
                                                            $GrantorDisplayName,
                                                            ("$GrantorRecipientType/$GrantorRecipientTypeDetails" -replace '^/$', ''),
                                                            $GrantorEnvironment,
                                                            '',
                                                            'ManagedBy',
                                                            'Allow',
                                                            'False',
                                                            'None',
                                                            $(($Trustee.displayname, $Trustee) | Select-Object -First 1),
                                                            $Trustee.PrimarySmtpAddress.address,
                                                            $Trustee.DisplayName,
                                                            ("$($Trustee.RecipientType.value)/$($Trustee.RecipientTypeDetails.value)" -replace '^/$', ''),
                                                            $TrusteeEnvironment
                                                        ) -join '";"'
                                                    ) + '"') -replace '(?<!;|^)"(?!;|$)', '""'))
                                    }
                                }
                                $ExportFileResult | Sort-Object -Unique | Out-File ([io.path]::ChangeExtension(($ExportFile), ('TEMP.{0:0000000}.txt' -f $RecipientID))) -Append -Encoding utf8 -Force
                            }
                        } catch {
                            $_ | Out-String | Write-Host
                        } finally {
                            if ($DebugFile) {
                                $null = Stop-Transcript
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                ).AddParameters(
                    @{
                        AllRecipients            = $AllRecipients
                        tempQueue                = $tempQueue
                        ExportFile               = $ExportFile
                        AllRecipientsGuidToIndex = $AllRecipientsGuidToIndex
                        DebugFile                = ([io.path]::ChangeExtension(($DebugFile), ('TEMP.{0:0000000}.txt' -f $_)))
                        ScriptPath               = $PSScriptRoot
                    }
                )

                $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
                $Handle = $Powershell.BeginInvoke($Object, $Object)

                $temp = '' | Select-Object PowerShell, Handle, Object
                $temp.PowerShell = $PowerShell
                $temp.Handle = $Handle
                $temp.Object = $Object
                [void]$runspaces.Add($Temp)
            }

            Write-Host "  $($tempQueueCount) grantors to check. Done (in steps of $($UpdateInterval)):"

            $lastCount = -1
            while (($runspaces.Handle.IsCompleted -contains $False)) {
                Start-Sleep -Seconds 1
                $done = ($tempQueueCount - $tempQueue.count - ($runspaces.Handle.IsCompleted | Where-Object { $_ -eq $false }).count)
                for ($x = $lastCount; $x -le $done; $x++) {
                    if (($x -gt $lastCount) -and (($x % $UpdateInterval -eq 0) -or ($x -eq $tempQueueCount))) {
                        Write-Host (("`b" * 100) + ('    {0:0000000} @{1}@' -f $x, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'))) -NoNewline
                        if ($x -eq 0) { Write-Host }
                        $lastCount = $x
                    }
                }
            }
            Write-Host

            if ($tempQueue.count -ne 0) {
                Write-Host '  Not all grantors have been checked. Enable debugging option and check log file.' -ForegroundColor red
            }

            foreach ($runspace in $runspaces) {
                $runspace.PowerShell.Dispose()
            }

            $RunspacePool.dispose()
            'temp', 'powershell', 'handle', 'object', 'runspaces', 'runspacepool' | ForEach-Object { Remove-Variable -Name $_ }

            if ($DebugFile) {
                $null = Stop-Transcript
                Start-Sleep -Seconds 1
                foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
                    Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
                    Remove-Item $JobDebugFile -Force
                }
                $null = Start-Transcript -Path $DebugFile -Append -Force
            }

            [GC]::Collect(); Start-Sleep 1

        }
    } else {
        Write-Host '  Not required with current export settings.'
    }


    # Sort and combine temporary files
    Write-Host
    Write-Host "Create sorted export file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host "  '$ExportFile'"

    '"Grantor Primary SMTP";"Grantor Display Name";"Grantor Recipient Type";"Grantor Environment";"Folder";"Permission";"Allow/Deny";"Inherited";"InheritanceType";"Trustee Original Identity";"Trustee Primary SMTP";"Trustee Display Name";"Trustee Recipient Type";"Trustee Environment"' | Out-File $ExportFile -Encoding utf8 -Force

    foreach ($RecipientFile in (Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))) {
        Get-Content $RecipientFile | Sort-Object -Unique | Out-File $ExportFile -Encoding utf8 -Append -Force
        Remove-Item $Recipientfile -Force
    }

} catch {
    Write-Host 'Unexpected error. Exiting.'
    $_ | Out-String | Write-Host
} finally {
    Write-Host

    if ($ExchangeSession) { Remove-PSSession $ExchangeSession }

    Write-Host 'Closing runspaces and runspacepool, please wait'
    if ($runspaces) {
        foreach ($runspace in $runspaces) {
            $runspace.PowerShell.Dispose()
        }
    }
    if ($RunspacePool) {
        $RunspacePool.dispose()
    }

    foreach ($RecipientFile in (Get-ChildItem ([io.path]::ChangeExtension(($ExportFile), ('TEMP.*.txt'))))) {
        Get-Content $RecipientFile | Sort-Object -Unique | Out-File $ExportFile -Encoding utf8 -Append -Force
        Remove-Item $Recipientfile -Force
    }

    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host

    if ($DebugFile) {
        $null = Stop-Transcript
        Start-Sleep -Seconds 1
        foreach ($JobDebugFile in (Get-ChildItem ([io.path]::ChangeExtension(($DebugFile), ('TEMP.*.txt'))))) {
            Get-Content $JobDebugFile | Out-File $DebugFile -Append -Encoding utf8
            Remove-Item $JobDebugFile -Force
        }
    }

    Remove-Variable * -ErrorAction SilentlyContinue
    [System.GC]::Collect() # garbage collection
    Start-Sleep 1
}
