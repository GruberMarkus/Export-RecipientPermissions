# List of domains to check for group membership.
# If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.
# If a string starts with a minus or dash ('-domain-a.local'), the domain after the dash or minus is removed from the list (no wildcards allowed).
# All domains belonging to the Active Directory forest of the currently logged in user are always considered, but specific domains can be removed (`'*', '-childA1.childA.user.forest'`).
# When a cross-forest trust is detected by the '*' option, all domains belonging to the trusted forest are considered but specific domains can be removed (`'*', '-childX.trusted.forest'`).
# Default value: '*'
$TrustsToCheckForGroups = @('*')


$AdObjectsToCheck = @(
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
                    $UserAccount = ([ADSI]"$(($Search.FindOne()).path)")
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

                        if ($InternalTrustsToCheckForDomainLocalGroups -icontains $data[0]) {
                            $InternalTrustsToCheckForDomainLocalGroups.remove($data[0])
                        }

                        if ($ExternalTrustsToCheckForDomainLocalGroups -icontains $data[0]) {
                            $ExternalTrustsToCheckForDomainLocalGroups.remove($data[0])
                        }

                        $returnvalue = $false
                    }
                    $job.Done = $true
                }
            }
        }
    }
    return $returnvalue
}


# Setup
$script:jobs = New-Object System.Collections.ArrayList
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$Search = New-Object DirectoryServices.DirectorySearcher
$Search.PageSize = 1000
$MemberOfRecurse = @()


Write-Host "Enumerate domains @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$x = $TrustsToCheckForGroups
[System.Collections.ArrayList]$TrustsToCheckForGroups = @()
[System.Collections.ArrayList]$InternalTrustsToCheckForDomainLocalGroups = @()
if ($GraphOnly -eq $false) {
    # Users own domain/forest is always included
    try {
        $objTrans = New-Object -ComObject 'NameTranslate'
        $objNT = $objTrans.GetType()
        $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
        $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value))) # 12 = ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME
        $y = (([ADSI]"LDAP://$(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1) -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')).tolower()

        if ($y -ne '') {
            Write-Host "  User forest: $y"
            $TrustsToCheckForGroups += $y.tolower()

            # Internal trusts
            $Search.SearchRoot = "GC://$($TrustsToCheckForGroups[0])"
            $Search.Filter = '(ObjectClass=trustedDomain)'

            foreach ($TrustedDomain in $Search.FindAll()) {
                # Only intra-forest trusts
                if ($TrustedDomain.properties.trustattributes -eq 32) {
                    $InternalTrustsToCheckForDomainLocalGroups += $TrustedDomain.properties.name.tolower()
                }
            }

            $InternalTrustsToCheckForDomainLocalGroups = @(
                $InternalTrustsToCheckForDomainLocalGroups | Select-Object -Unique | Sort-Object @{Expression = {
                        $TemporaryArray = @($_.Split('.'))
                        [Array]::Reverse($TemporaryArray)
                        $TemporaryArray
                    }
                }
            )

            foreach ($InternalTrustToCheckForDomainLocalGroups in $InternalTrustsToCheckForDomainLocalGroups) {
                if ($InternalTrustToCheckForDomainLocalGroups -ine $y) {
                    Write-Host "    Child domain: $($InternalTrustToCheckForDomainLocalGroups)"
                }
            }

            # Other domains - either the list provided, or all outgoing and bidirectional trusts
            if ($x[0] -eq '*') {
                $Search.SearchRoot = "GC://$($TrustsToCheckForGroups[0])"
                $Search.Filter = '(ObjectClass=trustedDomain)'

                $TrustedDomains = @(
                    @($Search.FindAll()) | Sort-Object @{Expression = {
                            $TemporaryArray = @($_.properties.name.Split('.'))
                            [Array]::Reverse($TemporaryArray)
                            $TemporaryArray
                        }
                    }
                )

                foreach ($TrustedDomain in $TrustedDomains) {
                    # DNS name of the other side of the trust (could be the root domain or any subdomain)
                    # $TrustName = $TrustedDomain.properties.name

                    # Domain SID of the other side of the trust
                    # $TrustNameSID = (New-Object system.security.principal.securityidentifier($($TrustedDomain.properties.securityidentifier), 0)).value

                    # Trust direction
                    # https://docs.microsoft.com/en-us/dotnet/api/system.directoryservices.activedirectory.trustdirection?view=net-5.0
                    # $TrustDirectionNumber = $TrustedDomain.properties.trustdirection

                    # Trust type
                    # https://docs.microsoft.com/en-us/dotnet/api/system.directoryservices.activedirectory.trusttype?view=net-5.0
                    # $TrustTypeNumber = $TrustedDomain.properties.trusttype

                    # Trust attributes
                    # https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-adts/e9a2d23c-c31e-4a6f-88a0-6646fdb51a3c
                    # $TrustAttributesNumber = $TrustedDomain.properties.trustattributes

                    # Which domains does the current user have access to?
                    # No intra-forest trusts, only bidirectional trusts and outbound trusts

                    if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3)) ) {
                        if ($TrustedDomain.properties.trustattributes -eq 8) {
                            # Cross-forest trust
                            Write-Host "  Trusted forest: $($TrustedDomain.properties.name)"
                            if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                            } else {
                                $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
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
                            }
                        } else {
                            # No cross-forest trust
                            Write-Host "  Trusted domain: $($TrustedDomain.properties.name)"
                            if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                            } else {
                                $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
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
                        $TrustsToCheckForGroups += $y.tolower()
                    }
                } else {
                    Write-Host '    Remove trusted domain/forest.'
                    for ($z = 0; $z -lt $TrustsToCheckForGroups.Count; $z++) {
                        if ($TrustsToCheckForGroups[$z] -ieq $y.substring(1)) {
                            $TrustsToCheckForGroups[$z] = ''
                        }
                    }
                }
            }

            $TrustsToCheckForGroups = @($TrustsToCheckForGroups | Where-Object { $_ })


            Write-Host
            Write-Host "Check trusts for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            CheckADConnectivity @(@(@($TrustsToCheckForGroups) + @($InternalTrustsToCheckForDomainLocalGroups)) | Select-Object -Unique) 'LDAP' '  ' | Out-Null


            Write-Host
            Write-Host "Check trusts for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            CheckADConnectivity $TrustsToCheckForGroups 'GC' '  ' | Out-Null
            # $InternalTrustsToCheckForDomainLocalGroups does not need to be checked for GC connectivity, as local groups can only be queried via LDAP
        } else {
            Write-Host '  Problem connecting to logged in user''s Active Directory (no error message, but forest root domain name is empty), assuming Graph/Azure AD from now on.' -ForegroundColor Yellow
            $GraphOnly = $true
        }
    } catch {
        $y = ''
        Write-Verbose $error[0]
        Write-Host '  Problem connecting to logged in user''s Active Directory (see verbose stream for error message), assuming Graph/Azure AD from now on.' -ForegroundColor Yellow
        $GraphOnly = $true
    }
}


Write-Host
Write-Host "Enumerate group membership @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
foreach ($AdObjectToCheck in $AdObjectsToCheck) {
    Write-Host "  '$($AdObjectToCheck)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    try {
        $AdObjectToCheckDn = $null
        $AdObjectToCheckGuid = $null
        $objResult = $null

        # Get DN of AD object
        $objTrans = New-Object -ComObject 'NameTranslate'
        $objNT = $objTrans.GetType()
        $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
        $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($AdObjectToCheck)"))
        $AdObjectToCheckDn = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1)
        $AdObjectToCheckGuid = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}')
        $MemberOfRecurse += "$($ADObjectToCheck);;$($AdObjectToCheckGuid)"

        # Setup
        $GroupsSids = @()
        $SIDsToCheckInTrusts = @()

        # Security groups, no matter if enabled for mail or not
        Write-Verbose "      Security groups via LDAP query of tokengroups attribute @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $UserAccount = [ADSI]"LDAP://$($AdObjectToCheckDn)"
        $UserAccount.GetInfoEx(@('tokengroups'), 0)
        foreach ($sidBytes in $UserAccount.Properties.tokengroups) {
            $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)).value
            Write-Verbose "        $sid"
            $GroupsSIDs += $sid
            $SIDsToCheckInTrusts += $sid
        }

        # Distribution groups (static only)
        Write-Verbose "      Distribution groups (static only) via GC query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$(($($AdObjectToCheckDn) -split ',DC=')[1..999] -join '.')")
        $Search.filter = "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648))(member:1.2.840.113556.1.4.1941:=$($AdObjectToCheckDn)))"
        foreach ($DistributionGroup in $search.findall()) {
            if ($DistributionGroup.properties.objectsid) {
                $sid = (New-Object System.Security.Principal.SecurityIdentifier $($DistributionGroup.properties.objectsid), 0).value
                Write-Verbose "        $sid"
                $GroupsSIDs += $sid
                $SIDsToCheckInTrusts += $sid
            }

            foreach ($SidHistorySid in @($DistributionGroup.properties.sidhistory | Where-Object { $_ })) {
                $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                Write-Verbose "        $sid"
                $GroupsSIDs += $sid
                $SIDsToCheckInTrusts += $sid
            }
        }

        # Domain local groups in the current user's forest
        Write-Verbose "      Domain local groups in the current user's forest via LDAP query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        foreach ($InternalTrustToCheckForDomainLocalGroups in $InternalTrustsToCheckForDomainLocalGroups) {
            Write-Verbose "        $($InternalTrustToCheckForDomainLocalGroups) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($InternalTrustToCheckForDomainLocalGroups)")
            $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($AdObjectToCheckDn)))"
            foreach ($LocalGroup in $search.findall()) {
                if ($LocalGroup.properties.objectsid) {
                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $($LocalGroup.properties.objectsid), 0).value
                    Write-Verbose "          $sid"
                    $GroupsSIDs += $sid
                    $SIDsToCheckInTrusts += $sid
                }

                foreach ($SidHistorySid in @($LocalGroup.properties.sidhistory | Where-Object { $_ })) {
                    $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                    Write-Verbose "          $sid"
                    $GroupsSIDs += $sid
                    $SIDsToCheckInTrusts += $sid
                }
            }
        }

        $GroupsSIDs = @($GroupsSIDs | Select-Object -Unique)
        $SIDsToCheckInTrusts = @($SIDsToCheckInTrusts | Select-Object -Unique)

        # Loop through all domains to check if the mailbox account has a group membership there
        # Across a trust, a user can only be added to a domain local group.
        # Domain local groups can not be used outside their own domain, so we don't need to query recursively
        # But when it's a cross-forest trust, we need to query every every domain on that other side of the trust
        #   This is handled before by adding every single domain of a cross-forest trusted forest to $TrustsToCheckForGroups
        if ($SIDsToCheckInTrusts.count -gt 0) {
            $LdapFilterSIDs = '(|'
            foreach ($SidToCheckInTrusts in $SIDsToCheckInTrusts) {
                try {
                    $SidHex = @()
                    $ot = New-Object System.Security.Principal.SecurityIdentifier($SidToCheckInTrusts)
                    $c = New-Object 'byte[]' $ot.BinaryLength
                    $ot.GetBinaryForm($c, 0)
                    foreach ($char in $c) {
                        $SidHex += $('\{0:x2}' -f $char)
                    }
                    # Foreign Security Principals have an objectSID, but no sIDHistory
                    # The sIDHistory of the current object is part of $SIDsToCheckInTrusts and therefore also considered in $LdapFilterSIDs
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
            for ($DomainNumber = 0; $DomainNumber -lt $TrustsToCheckForGroups.count; $DomainNumber++) {
                if (($TrustsToCheckForGroups[$DomainNumber] -ne '') -and ($TrustsToCheckForGroups[$DomainNumber] -ine $UserDomain) -and ($UserDomain -ne '')) {
                    Write-Host "    $($TrustsToCheckForGroups[$DomainNumber]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustsToCheckForGroups[$DomainNumber])")
                    $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                    foreach ($fsp in $Search.FindAll()) {
                        if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                            # A Foreign Security Principal (FSP) is created in each (sub)domain in which it is granted permissions
                            # A FSP it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal
                            # FSPs have on tokengroups attribute, which would not contain domain local groups anyhow
                            # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member, incl. nesting
                            Write-Verbose "      Found ForeignSecurityPrincipal $($fsp.properties.cn) in $((($fsp.path -split ',DC=')[1..999] -join '.'))"
                            
                            try {
                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                foreach ($group in $Search.findall()) {
                                    $sid = New-Object System.Security.Principal.SecurityIdentifier($group.properties.objectsid[0], 0).value
                                    Write-Verbose "        $sid"
                                    $GroupsSIDs += $sid

                                    foreach ($SidHistorySid in @($group.properties.sidhistory | Where-Object { $_ })) {
                                        $sid = (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value
                                        Write-Verbose "        $sid"
                                        $GroupsSIDs += $sid
                                    }

                                }
                            } catch {
                                Write-Host "        Error: $($error[0].exception)" -ForegroundColor red
                            }
                        }
                    }
                }
            }
        }


        # Translate SIDs to GUIDs
        Write-Verbose "      SIDs to GUIDs @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        foreach ($GroupSid in $GroupsSids) {
            $objTrans = New-Object -ComObject 'NameTranslate'
            $objNT = $objTrans.GetType()
            $null = $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
            $null = $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, "$($GroupSid)"))
            $GroupGuid = $($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 7).trimstart('{').trimend('}'))
            Write-Verbose "      $($GroupSid): $($GroupGuid)"
            $MemberOfRecurse += "$($ADObjectToCheck);$($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3));$($GroupGuid)"
        }
    } catch {
        @(
            'ERROR',
            "$($_ | Out-String)",
            "AdObjectToCheck: $($AdObjectToCheck)",
            "AdObjectToCheckDn: $($AdObjectToCheckDn)",
            "AdObjectToCheckGuid $($AdObjectToCheckGuid)",
            "objResult.properties.'msds-principalname': $($objResult.properties.'msds-principalname')"
        ) | ForEach-Object {
            Write-Host "    $_" -ForegroundColor Red
        }
    }
}


Write-Host
Write-Host "Final MemberOfRecurse result @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$MemberOfRecurse = $MemberOfRecurse | Select-Object -Unique
$MemberOfRecurse = $MemberOfRecurse | Select-Object -Unique | ConvertFrom-Csv -Delimiter ';' -Header @('Original object', 'MemberOf recurse NT4 name', 'MemberOf recurse AD GUID')
$MemberOfRecurse | Out-String -Stream | ForEach-Object { Write-Output "  $_" }


Write-Host
Write-Host "Configure and start Export-RecipientPermissions (demo) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
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
    ExportFileFilter                            = "if (`$ExportFileLine.'Trustee AD ObjectGUID' -iin $('@(''' + (@($MemberOfRecurse.'MemberOf recurse AD GUID') -join ''', ''') + ''')')) { `$true } else { `$false }"

    ExportFile                                  = '.\export\Export-RecipientPermissions_Result_MemberOfRecurse.csv'
    ErrorFile                                   = '.\export\Export-RecipientPermissions_Error_MemberOfRecurse.csv'
    DebugFile                                   = $null

    verbose                                     = $false
}

Write-Host '  Parameters ($params hashtable)'
$params | Out-String -Stream | ForEach-Object { Write-Host "  $($_.trim())" }

Write-Host '  Run Export-RecipientPermissions with parameters from $params (demo)'
Write-Host "    '& ..\..\Export-RecipientPermissions.ps1 @params'"


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"