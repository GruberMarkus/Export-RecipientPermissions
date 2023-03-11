[CmdletBinding(PositionalBinding = $false)]


Param(
    # If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.
    # If a string starts with a minus or dash ('-domain-a.local'), the domain after the dash or minus is removed from the list (no wildcards allowed).
    # All domains belonging to the Active Directory forest of the currently logged in user are always considered, but specific domains can be removed (`'*', '-childA1.childA.user.forest'`).
    # When a cross-forest trust is detected by the '*' option, all domains belonging to the trusted forest are considered but specific domains can be removed (`'*', '-childX.trusted.forest'`).
    # Default value: '*'
    [string[]]$TrustsToCheckForGroups = @('*'),


    [string[]]$AdObjectsToCheck = @(
        # Accepted string formats (examples are in this order):
        #   Distinguished Name
        #   Canonical name
        #   Domain\SamAccountName (pre Windows 2000 logon name, NT4 logon name)
        #   User Principal Name (UPN)
        #   AD Object GUID (in curly braces)
        #   SID or SIDHistory
        'CN=Jeff Smith,CN=users,DC=example,DC=com',
        'example.com/Users/Hank Morgan',
        'EXAMPLE\GruberMa',
        'John.Carpenter@example.com',
        '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}',
        'S-1-5-21-1180699209-877415012-3182924384-1004'
    )
)

function CheckADConnectivity {
    param (
        [array]$CheckDomains,
        [string]$CheckProtocolText,
        [string]$Indent
    )
    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 25)
    $RunspacePool.Open()

    for ($DomainNumber = 0; $DomainNumber -lt $CheckDomains.count; $DomainNumber++) {
        if ($($CheckDomains[$DomainNumber]) -eq '') {
            continue
        }

        $PowerShell = [powershell]::Create()
        $PowerShell.RunspacePool = $RunspacePool

        [void]$PowerShell.AddScript( {
                Param (
                    [string]$CheckDomain,
                    [string]$CheckProtocolText
                )
                $DebugPreference = 'Continue'
                Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"
                Write-Output "$CheckDomain"
                $Search = New-Object DirectoryServices.DirectorySearcher
                $Search.PageSize = 1000
                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($CheckProtocolText)://$CheckDomain")
                $Search.filter = '(objectclass=user)'
                try {
                    $null = ([ADSI]"$(($Search.FindOne()).path)")
                    Write-Output 'QueryPassed'
                } catch {
                    Write-Output 'QueryFailed'
                }
            }).AddArgument($($CheckDomains[$DomainNumber])).AddArgument($CheckProtocolText)
        $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
        $Handle = $PowerShell.BeginInvoke($Object, $Object)
        $temp = '' | Select-Object PowerShell, Handle, Object, StartTime, Done
        $temp.PowerShell = $PowerShell
        $temp.Handle = $Handle
        $temp.Object = $Object
        $temp.StartTime = $null
        $temp.Done = $false
        [void]$script:jobs.Add($Temp)
    }
    while (($script:jobs.Done | Where-Object { $_ -eq $false }).count -ne 0) {
        foreach ($job in $script:jobs) {
            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -match 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -replace '[^0-9]'
                $job.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $job.StartTime) {
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $job.Object[0..$(($job.object).count - 1)]
                    Write-Host "$Indent$($data[0])"
                    if ($data -icontains 'QueryPassed') {
                        Write-Host "$Indent  $CheckProtocolText query successful"
                        $returnvalue = $true
                    } else {
                        Write-Host "$Indent  $CheckProtocolText query failed, remove domain from list." -ForegroundColor Red
                        Write-Host "$Indent  If this error is permanent, check firewalls, DNS and AD trust. Consider parameter TrustsToCheckForGroups." -ForegroundColor Red

                        if ($TrustsToCheckForGroups -icontains $data[0]) {
                            $TrustsToCheckForGroups.remove($data[0])
                        }

                        $LookupDomainsToTrusts.remove($data[0])

                        $returnvalue = $false
                    }
                    $job.Done = $true
                }
            }
        }
    }
    return $returnvalue
}


function ConvertSidToGuidAndFillResult {
    param (
        $sid,
        $AdObjectToCheckDn,
        $indent
    )

    try {
        $objTrans = New-Object -ComObject 'NameTranslate'
        $objNT = $objTrans.GetType()
        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($sid)"))
        $GroupGuid = $($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}'))
        Write-Verbose "$indent$($GroupGuid)"

        $script:MemberOfRecurse += New-Object PSObject -Property (
            [ordered]@{
                'Original object'                      = $ADObjectToCheck.ToString()
                'MemberOf recurse group objectGUID'    = $GroupGuid.ToString()
                'MemberOf recurse group canonicalName' = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 2).ToString()
            }
        )

        $script:SIDsToCheckInTrusts += $sid
    } catch {
        try {
            # Non-domain-specific well-known SID with a domain specific ObjectGUID?
            $SidHex = @()
            $ot = New-Object System.Security.Principal.SecurityIdentifier($Sid)
            $c = New-Object 'byte[]' $ot.BinaryLength
            $ot.GetBinaryForm($c, 0)
            foreach ($char in $c) {
                $SidHex += $('\{0:x2}' -f $char)
            }

            $local:Search = New-Object DirectoryServices.DirectorySearcher
            $local:Search.PageSize = 1000
            $local:Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')")
            $local:Search.filter = "(&(objectclass=group)(objectsid=$($SidHex -join '')))"

            @('canonicalname', 'objectguid') | ForEach-Object {
                if (-not $local:search.PropertiesToLoad.Contains($_)) {
                    $null = $local:search.PropertiesToLoad.add($_)
                }
            }
            $Group = $local:search.FindOne()
            $GroupGuid = [guid]::new($Group.Properties.objectguid[0]).guid
            Write-Verbose "$indent$($GroupGuid)"
            $script:MemberOfRecurse += New-Object PSObject -Property (
                [ordered]@{
                    'Original object'                      = $ADObjectToCheck.ToString()
                    'MemberOf recurse group objectGUID'    = $GroupGuid.ToString()
                    'MemberOf recurse group canonicalName' = $Group.Properties.canonicalname[0].ToString()
                }
            )
        } catch {
            Write-Verbose "$indent$($_ | Out-String)"
        }
    }
}


# Setup
$script:jobs = New-Object System.Collections.ArrayList
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$Search = New-Object DirectoryServices.DirectorySearcher
$Search.PageSize = 1000
$script:MemberOfRecurse = @()


Write-Host "Enumerate domains @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$x = $TrustsToCheckForGroups
[System.Collections.ArrayList]$TrustsToCheckForGroups = @()
$LookupDomainsToTrusts = @{}
# Users own domain/forest is always included
try {
    $objTrans = New-Object -ComObject 'NameTranslate'
    $objNT = $objTrans.GetType()
    $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
    $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value))) # 12 = ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME
    $UserForest = (([ADSI]"LDAP://$(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1) -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')).tolower()

    if ($UserForest -ne '') {
        Write-Host "  User forest: $UserForest"
        $TrustsToCheckForGroups += $UserForest.tolower()
        $LookupDomainsToTrusts.add($UserForest, $UserForest)

        $Search.SearchRoot = "GC://$($UserForest)"
        $Search.Filter = '(ObjectClass=trustedDomain)'

        $TrustedDomains = @(
            @($Search.FindAll()) | Sort-Object @{Expression = {
                    $TemporaryArray = @($_.properties.name.Split('.'))
                    [Array]::Reverse($TemporaryArray)
                    $TemporaryArray
                }
            }
        )

        # Internal trusts
        foreach ($TrustedDomain in $TrustedDomains) {
            if (($TrustedDomain.properties.trustattributes -eq 32) -and ($TrustedDomain.properties.name -ine $UserForest) -and (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower()))) {
                Write-Host "    Child domain: $($TrustedDomain.properties.name.tolower())"
                $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $UserForest)
            }
        }

        # Other trusts
        if ($x[0] -eq '*') {
            foreach ($TrustedDomain in $TrustedDomains) {
                # No intra-forest trusts, only bidirectional trusts and outbound trusts
                if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3)) ) {
                    if ($TrustedDomain.properties.trustattributes -eq 8) {
                        # Cross-forest trust
                        Write-Host "  Trusted forest: $($TrustedDomain.properties.name.tolower())"
                        if ("-$($TrustedDomain.properties.name)" -iin $x) {
                            Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name.tolower())'"
                        } else {
                            $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                            $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                        }

                        $temp = @(
                            @(@(Resolve-DnsName -Name "_gc._tcp.$($TrustedDomain.properties.name)" -Type srv).nametarget) | ForEach-Object { ($_ -split '\.')[1..999] -join '.' } | Where-Object { $_ -ine $TrustedDomain.properties.name } | Select-Object -Unique | Sort-Object @{Expression = {
                                    $TemporaryArray = @($_.Split('.'))
                                    [Array]::Reverse($TemporaryArray)
                                    $TemporaryArray
                                }
                            }
                        )

                        $temp | ForEach-Object {
                            Write-Host "    Child domain: $($_.tolower())"
                            $LookupDomainsToTrusts.add($_.tolower(), $TrustedDomain.properties.name.tolower())
                        }
                    } else {
                        # No cross-forest trust
                        Write-Host "  Trusted domain: $($TrustedDomain.properties.name)"
                        if ("-$($TrustedDomain.properties.name)" -iin $x) {
                            Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                        } else {
                            $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                            $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                        }
                    }
                }
            }
        }

        for ($a = 0; $a -lt $x.Count; $a++) {
            if (($a -eq 0) -and ($x[$a] -ieq '*')) {
                continue
            }

            $y = ($x[$a] -replace ('DC=', '') -replace (',', '.')).tolower()

            if ($y -eq $x[$a]) {
                Write-Host "  User provided trusted domain/forest: $y"
            } else {
                Write-Host "  User provided trusted domain/forest: $($x[$a]) -> $y"
            }

            if (($a -ne 0) -and ($x[$a] -ieq '*')) {
                Write-Host '    Entry * is only allowed at first position in list. Skip entry.' -ForegroundColor Red
                continue
            }

            if ($y -match '[^a-zA-Z0-9.-]') {
                Write-Host '    Allowed characters are a-z, A-Z, ., -. Skip entry.' -ForegroundColor Red
                continue
            }

            if (-not ($y.StartsWith('-'))) {
                if ($TrustsToCheckForGroups -icontains $y) {
                    Write-Host '    Trusted domain/forest already in list.' -ForegroundColor Yellow
                } else {
                    if ($TrustedDomains.properties.name -icontains $y) {
                        foreach ($TrustedDomain in @($TrustedDomains | Where-Object { $_.properties.name -ieq $y })) {
                            # No intra-forest trusts, only bidirectional trusts and outbound trusts
                            if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3)) ) {
                                if ($TrustedDomain.properties.trustattributes -eq 8) {
                                    # Cross-forest trust
                                    Write-Host "    Trusted forest: $($TrustedDomain.properties.name)"
                                    if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                        Write-Host "      Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                    } else {
                                        $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                        $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                    }

                                    $temp = @(
                                        @(@(Resolve-DnsName -Name "_gc._tcp.$($TrustedDomain.properties.name)" -Type srv).nametarget) | ForEach-Object { ($_ -split '\.')[1..999] -join '.' } | Where-Object { $_ -ine $TrustedDomain.properties.name } | Select-Object -Unique | Sort-Object @{Expression = {
                                                $TemporaryArray = @($_.Split('.'))
                                                [Array]::Reverse($TemporaryArray)
                                                $TemporaryArray
                                            }
                                        }
                                    )

                                    $temp | ForEach-Object {
                                        Write-Host "      Child domain: $($_.tolower())"
                                        $LookupDomainsToTrusts.add($_.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                } else {
                                    # No cross-forest trust
                                    Write-Host "    Trusted domain: $($TrustedDomain.properties.name)"
                                    if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                        Write-Host "      Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                    } else {
                                        $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                        $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }
                            }
                        }
                    } else {
                        Write-Host '    No trust to this domain/forest found.' -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host '    Remove trusted domain/forest.'
                for ($z = 0; $z -lt $TrustsToCheckForGroups.Count; $z++) {
                    if ($TrustsToCheckForGroups[$z] -ieq $y.substring(1)) {
                        $TrustsToCheckForGroups.RemoveAt($z)
                        $LookupDomainsToTrusts = $LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $y.substring(1) }
                    }
                }
            }
        }

        $TrustsToCheckForGroups = @($TrustsToCheckForGroups | Where-Object { $_ })


        Write-Host
        Write-Host "Check trusts for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        CheckADConnectivity @(@(@($TrustsToCheckForGroups) + @($LookupDomainsToTrusts.GetEnumerator() | ForEach-Object { $_.Name })) | Select-Object -Unique) 'LDAP' '  ' | Out-Null


        Write-Host
        Write-Host "Check trusts for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        CheckADConnectivity $TrustsToCheckForGroups 'GC' '  ' | Out-Null
    }
} catch {
    $y = ''
    Write-Verbose $error[0]
    Write-Host '  Problem connecting to logged in user''s Active Directory (see verbose stream for error message).' -ForegroundColor Yellow
}



Write-Host
Write-Host "Enumerate group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
foreach ($AdObjectToCheck in $AdObjectsToCheck) {
    Write-Host "  '$($AdObjectToCheck)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    try {
        $AdObjectToCheckDn = $null
        $AdObjectToCheckGuid = $null

        # Get DN of AD object
        $objTrans = New-Object -ComObject 'NameTranslate'
        $objNT = $objTrans.GetType()
        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AdObjectToCheck)"))
        $AdObjectToCheckDn = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1)
        $AdObjectToCheckGuid = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
        $script:MemberOfRecurse += New-Object PSObject -Property (
            [ordered]@{
                'Original object'                      = $ADObjectToCheck.ToString()
                'MemberOf recurse group objectGUID'    = $ADObjectToCheckGuid.ToString()
                'MemberOf recurse group canonicalName' = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 2).ToString()
            }
        )

        $script:SIDsToCheckInTrusts = @()

        Write-Verbose "    $($LookupDomainsToTrusts[$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        # Security groups, no matter if enabled for mail or not
        Write-Verbose "      Security groups via LDAP query of tokengroups attribute @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $UserAccount = [ADSI]"LDAP://$($AdObjectToCheckDn)"
        $UserAccount.GetInfoEx(@('tokengroups'), 0)
        foreach ($sidBytes in $UserAccount.Properties.tokengroups) {
            $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)).value
            Write-Verbose "        $sid"
            ConvertSidToGuidAndFillResult $sid $AdObjectToCheckDn '          '
        }

        # Distribution groups (static only)
        Write-Verbose "      Distribution groups (static only) via GC query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')")
        $Search.filter = "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648))(member:1.2.840.113556.1.4.1941:=$($AdObjectToCheckDn)))"
        foreach ($DistributionGroup in $search.findall()) {
            if ($DistributionGroup.properties.objectsid) {
                $sid = (New-Object System.Security.Principal.SecurityIdentifier $($DistributionGroup.properties.objectsid), 0).value
                Write-Verbose "        $sid"
                ConvertSidToGuidAndFillResult $sid $AdObjectToCheckDn '          '
            }

            foreach ($SidHistorySid in @($DistributionGroup.properties.sidhistory | Where-Object { $_ })) {
                $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                Write-Verbose "        SidHistory: $sid"
                $script:SIDsToCheckInTrusts += $sid
            }
        }

        # Domain local groups
        Write-Verbose "      Domain local groups via LDAP query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        foreach ($DomainToCheckForDomainLocalGroups in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ieq $LookupDomainsToTrusts[$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')] }).name)) {
            Write-Verbose "        $($DomainToCheckForDomainLocalGroups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DomainToCheckForDomainLocalGroups)")
            $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($AdObjectToCheckDn)))"
            foreach ($LocalGroup in $search.findall()) {
                if ($LocalGroup.properties.objectsid) {
                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $($LocalGroup.properties.objectsid), 0).value
                    Write-Verbose "          $sid"
                    ConvertSidToGuidAndFillResult $sid $AdObjectToCheckDn '            '
                }

                foreach ($SidHistorySid in @($LocalGroup.properties.sidhistory | Where-Object { $_ })) {
                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                    Write-Verbose "          SidHistory: $sid"
                    $script:SIDsToCheckInTrusts += $sid
                }
            }
        }

        # Loop through all domains to check if the mailbox account has a group membership there
        # Across a trust, a user can only be added to a domain local group.
        # Domain local groups can not be used outside their own domain, so we don't need to query recursively
        # But when it's a cross-forest trust, we need to query every every domain on that other side of the trust
        #   This is handled before by adding every single domain of a cross-forest trusted forest to $TrustsToCheckForGroups
        if ($script:SIDsToCheckInTrusts.count -gt 0) {
            $script:SIDsToCheckInTrusts = @($script:SIDsToCheckInTrusts | Select-Object -Unique)
            $LdapFilterSIDs = '(|'

            foreach ($SidToCheckInTrusts in $script:SIDsToCheckInTrusts) {
                try {
                    $SidHex = @()
                    $ot = New-Object System.Security.Principal.SecurityIdentifier($SidToCheckInTrusts)
                    $c = New-Object 'byte[]' $ot.BinaryLength
                    $ot.GetBinaryForm($c, 0)
                    foreach ($char in $c) {
                        $SidHex += $('\{0:x2}' -f $char)
                    }
                    # Foreign Security Principals have an objectSID, but no sIDHistory
                    # The sIDHistory of the current object is part of $script:SIDsToCheckInTrusts and therefore also considered in $LdapFilterSIDs
                    $LdapFilterSIDs += ('(objectsid=' + $($SidHex -join '') + ')')
                } catch {
                    Write-Host '      Error creating LDAP filter for search across trusts.' -ForegroundColor Red
                    $error[0]
                }
            }
            $LdapFilterSIDs += ')'
        } else {
            $LdapFilterSIDs = ''
        }

        if ($LdapFilterSids -ilike '*(objectsid=*') {
            # Across each trust, search for all Foreign Security Principals matching a SID from our list
            foreach ($TrustToCheckForFSPs in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')] }).value | Select-Object -Unique)) {
                Write-Host "    $($TrustToCheckForFSPs) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustToCheckForFSPs)")
                $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                foreach ($fsp in $Search.FindAll()) {
                    if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                        # A Foreign Security Principal (FSP) is created in each (sub)domain in which it is granted permissions
                        # A FSP it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal
                        # FSPs have on tokengroups attribute, which would not contain domain local groups anyhow
                        # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member, incl. nesting
                        Write-Verbose "      Found ForeignSecurityPrincipal $($fsp.properties.cn) in $((($fsp.path -split ',DC=')[1..999] -join '.'))"

                        if ($((($fsp.path -split ',DC=')[1..999] -join '.')) -iin @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')] }).name)) {
                            try {
                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                foreach ($group in $Search.findall()) {
                                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $($group.properties.objectsid), 0).value
                                    Write-Verbose "        $sid"
                                    ConvertSidToGuidAndFillResult $sid $AdObjectToCheckDn '          '

                                    foreach ($SidHistorySid in @($group.properties.sidhistory | Where-Object { $_ })) {
                                        $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                        Write-Verbose "        SidHistory: $sid"
                                    }
                                }
                            } catch {
                                Write-Host "        Error: $($error[0].exception)" -ForegroundColor red
                            }
                        } else {
                            Write-Verbose "        Ignoring, because '$($fsp.path)' is not part of a trust in TrustsToCheckForGroups."
                        }
                    }
                }
            }
        }
    } catch {
        @(
            'ERROR',
            "$($_ | Out-String)",
            "AdObjectToCheck: $($AdObjectToCheck)",
            "AdObjectToCheckDn: $($AdObjectToCheckDn)",
            "AdObjectToCheckGuid $($AdObjectToCheckGuid)"
        ) | ForEach-Object {
            Write-Host "    $_" -ForegroundColor Red
        }
    }
}


Write-Host
Write-Host "Final MemberOfRecurse result @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$script:MemberOfRecurse = $script:MemberOfRecurse | Select-Object -Property * -Unique | Sort-Object -Property 'Original object', 'MemberOf recurse group canonicalName', 'MemberOf recurse group objectGUID'
$script:MemberOfRecurse | Format-Table


Write-Host
Write-Host "Configure and start Export-RecipientPermissions (demo) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$params = @{
    ExportFromOnPrem                            = $true
    UseDefaultCredential                        = $true

    ExportMailboxAccessRights                   = $true
    ExportMailboxAccessRightsSelf               = $false
    ExportMailboxAccessRightsInherited          = $true
    ExportMailboxFolderPermissions              = $true
    ExportMailboxFolderPermissionsAnonymous     = $false
    ExportMailboxFolderPermissionsDefault       = $false
    ExportMailboxFolderPermissionsOwnerAtLocal  = $false
    ExportMailboxFolderPermissionsMemberAtLocal = $false
    ExportSendAs                                = $true
    ExportSendAsSelf                            = $false
    ExportSendOnBehalf                          = $true
    ExportManagedBy                             = $true
    ExportLinkedMasterAccount                   = $true
    ExportPublicFolderPermissions               = $true
    ExportPublicFolderPermissionsAnonymous      = $false
    ExportPublicFolderPermissionsDefault        = $false
    ExportForwarders                            = $true
    ExportManagementRoleGroupMembers            = $true
    ExportDistributionGroupMembers              = 'None'
    ExportGroupMembersRecurse                   = $false
    ExpandGroups                                = $false
    ExportGuids                                 = $true
    ExportGrantorsWithNoPermissions             = $false
    ExportTrustees                              = 'All'

    RecipientProperties                         = @()
    GrantorFilter                               = $null
    TrusteeFilter                               = $null
    ExportFileFilter                            = "if (`$ExportFileLine.'Trustee AD ObjectGUID' -iin $('@(''' + (@($script:MemberOfRecurse.'MemberOf recurse group objectGUID') -join ''', ''') + ''')')) { `$true } else { `$false }"

    ExportFile                                  = '.\export\Export-RecipientPermissions_Result_MemberOfRecurse.csv'
    ErrorFile                                   = '.\export\Export-RecipientPermissions_Error_MemberOfRecurse.csv'
    DebugFile                                   = $null

    verbose                                     = $false
}

Write-Host '  Parameters ($params hashtable)'
$params | Format-Table

Write-Host '  Run Export-RecipientPermissions with parameters from $params (demo)'
Write-Host "    '& ..\..\Export-RecipientPermissions.ps1 @params'"


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"