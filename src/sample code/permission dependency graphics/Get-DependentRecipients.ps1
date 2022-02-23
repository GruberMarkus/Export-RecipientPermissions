param (
    # Input: Primary SMTP addresses of users you want to move. Every line in this file containing the @-sign is imported.
    $InitialRecipientsFile = '.\input.csv',

    # Input: Export file of Export-RecipientPermissions.ps1
    $RecipientPermissionsFile = '.\Export-RecipientPermissions_Output.csv', 

    # Input: List of primary SMTP addresses to ignore (service accounts, for example). Wildcards are not allowed.
    $PrimarySMTPAddressesToIgnore = @('xxx@domain.com', 'yyy@domain.com'),

    # Output: List of all recipients, initial and additional (a.k.a. dependent)
    $ExportAllRecipientsFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AllRecipients.csv',

    # Output: List of additional (a.k.a. dependent) recipients
    $ExportAdditionalRecipientsFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AdditionalRecipients.csv',

    # Output: File containing initial and additional recipients, their permission and information why they are dependent recipients
    $RecipientPermissionsFileNew = '.\Get-DependentRecipients_Output\Export-RecipientPermissions_Output_Modified.csv',

    # Output: Graphical representation of the file above in Graph Modeling Language (GML). Can be viewed with yWorks yEd, for example.
    $AllRecipientsGMLFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AllRecipients.gml',

    # Output: Summary file (number of initial recipients, additional recipients, total recipients, and root cause permission count)
    $SummaryFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_Summary.txt'
)

Clear-Host

Set-Location $PSScriptRoot

$InitialRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InitialRecipientsFile)
$RecipientPermissionsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecipientPermissionsFile)
$ExportAllRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAllRecipientsFile)
$ExportAdditionalRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAdditionalRecipientsFile)
$AllRecipientsGMLFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AllRecipientsGMLFile)
$RecipientPermissionsFileNew = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecipientPermissionsFileNew)

New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAllRecipientsFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAdditionalRecipientsFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $AllRecipientsGMLFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $RecipientPermissionsFileNew) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $SummaryFile) | Out-Null

if (-not (Test-Path $RecipientPermissionsFile)) {
    Write-Host "Recipient permissions file '$RecipientPermissionsFile' not found, exiting."
    break
}

if (-not (Test-Path $InitialRecipientsFile)) {
    Write-Host "Initial recipients file '$InitialRecipientsFile' not found, exiting."
    break
}

if (Test-Path $ExportAllRecipientsFile) { (Remove-Item $ExportAllRecipientsFile) }
if (Test-Path $ExportAdditionalRecipientsFile) { (Remove-Item $ExportAdditionalRecipientsFile) }
if (Test-Path $AllRecipientsGMLFile) { (Remove-Item $AllRecipientsGMLFile) }
if (Test-Path $RecipientPermissionsFileNew) { (Remove-Item $RecipientPermissionsFileNew) }
if (Test-Path $SummaryFile) { (Remove-Item $SummaryFile) }


Write-Host "Import recipient permissions from '$RecipientPermissionsFile'"
$RecipientPermissions = [system.collections.arraylist]::new()
$RecipientPermissions.AddRange(@(Import-Csv $RecipientPermissionsFile -Delimiter ';' | Select-Object *, 'Grantor InitialOrAdditional', 'Trustee InitialOrAdditional'))


Write-Host
Write-Host ("Filter $($RecipientPermissions.count) recipient permissions")
for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
    # Remove non-resolvable trustees
    if (-not $RecipientPermissions[$x].'Trustee Primary SMTP') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove folder permissions (folder permissions do work cross-premises)
    if ($RecipientPermissions[$x].'Folder') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove lines where grantor or trustee is already in the cloud
    if (($RecipientPermissions[$x].'Grantor Environment' -ieq 'Cloud') -or ($RecipientPermissions[$x].'Trustee Environment' -ieq 'Cloud')) {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove lines containing "FullAccess" permission ("FullAccess" works cross-premises)
    if ($RecipientPermissions[$x].'Permission' -ieq 'FullAccess') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove lines containing "SendOnBehalf" permission ("SendOnBehalf" works cross-premises)
    if ($RecipientPermissions[$x].'Permission' -ieq 'SendOnBehalf') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove lines containing "SendAs" permission ("SendAs" does work cross-premises)
    if ($RecipientPermissions[$x].'Permission' -ieq 'SendAs') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove lines with "ManagedBy" permission ("ManagedBy" does not work cross premises)
    #if ($RecipientPermissions[$x].'Permission' -ieq 'ManagedBy') {
    #    $RecipientPermissions.remove($x)
    #    $x--
    #    continue
    #}


    # Remove lines with primary SMTP addresses to ignore
    if (($RecipientPermissions[$x].'Grantor Primary SMTP' -iin $PrimarySMTPAddressesToIgnore) -or ($RecipientPermissions[$x].'Trustee Primary SMTP' -iin $PrimarySMTPAddressesToIgnore)) {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }
}

$RecipientPermissions.TrimToSize()


# Select all columns and create additional columns
Write-Host "Number of permissions to consider: $($RecipientPermissions.count)"


Write-Host "Import initial recipients from '$InitialRecipientsFile'"
$InitialRecipients = @(Import-Csv -Header PrimarySMTPAddress $InitialRecipientsFile | Select-Object PrimarySMTPAddress | Where-Object { $_.PrimarySMTPAddress.contains('@') } | Sort-Object PrimarySMTPAddress -Unique)
$InitialRecipientsCount = $InitialRecipients.count
$AdditionalRecipients = $InitialRecipients



Write-Host
$AllRecipientsGMLFileString = 'graph' + [Environment]::NewLine + '[' + [Environment]::NewLine + '    hierarchic 1' + [Environment]::NewLine + '    directed 1'
$AllRecipientsGMLFileEdgeString = $null
$AdditionalRecipientsString = $null
$InitialRecipientsString = 'Primary SMTP Address;Recipient Type;Environment' + [Environment]::NewLine

# $AdditionalRecipients will be expanded step by step until no new entries are added
for ($i = 0; $i -lt $AdditionalRecipients.count; $i++) {
    Write-Host ('{0:000000}/{1:000000}, {2}' -f ($i + 1), $AdditionalRecipients.count, $AdditionalRecipients[$i].PrimarySMTPAddress)
    for ($j = 0; $j -lt $RecipientPermissions.count; $j++) {
        # Where is the current recipient a grantor and who are the trustees?
        if (($RecipientPermissions[$j].'Grantor Primary SMTP' -eq $AdditionalRecipients[$i].PrimarySMTPAddress) -and (-not $RecipientPermissions[$j].'Grantor InitialOrAdditional')) {
            if (($i -lt $InitialRecipientsCount) -and ($InitialRecipientsstring -cnotlike ('*' + [Environment]::NewLine + $AdditionalRecipients[$i].PrimarySMTPAddress + ';*'))) { $InitialRecipientsString += (($AdditionalRecipients[$i].PrimarySMTPAddress) + ';' + ($RecipientPermissions[$j].'Grantor Recipient Type') + [Environment]::NewLine) }
            if ($AdditionalRecipients.PrimarySMTPAddress -cnotcontains $RecipientPermissions[$j].'Trustee Primary SMTP') {
                $AdditionalRecipients += New-Object PsObject -Property @{ PrimarySMTPAddress = $RecipientPermissions[$j].'Trustee Primary SMTP' }
                $AdditionalRecipientsString += (($RecipientPermissions[$j].'Trustee Primary SMTP') + ';' + ($RecipientPermissions[$j].'Trustee Recipient Type') + ';' + ($RecipientPermissions[$j].'Trustee Environment') + [Environment]::NewLine)
            }
            if (-not $RecipientPermissions[$j].'Trustee InitialOrAdditional') {
                $AllRecipientsGMLFileEdgeString += [Environment]::NewLine + '    edge' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        source ' + $i + [Environment]::NewLine + '        target ' + $AdditionalRecipients.primarysmtpaddress.IndexOf($RecipientPermissions[$j].'Trustee Primary SMTP') + [Environment]::NewLine + '        label "grants ' + $RecipientPermissions[$j].'Permission' + '"' + [Environment]::NewLine + '    ]'
            }
            if ($i -lt $InitialRecipientsCount) {
                $RecipientPermissions[$j].'Grantor InitialOrAdditional' = 'Initial'
            } else {
                $RecipientPermissions[$j].'Grantor InitialOrAdditional' = 'Additional'
            }
        }

        # Where is the current recipient a trustee and who is the grantor?
        if ($RecipientPermissions[$j].'Trustee Primary SMTP' -ceq $RecipientPermissions[$j].'Grantor Primary SMTP') {
            $RecipientPermissions[$j].'Trustee InitialOrAdditional' = $RecipientPermissions[$j].'Grantor InitialOrAdditional'
        }
        if (($RecipientPermissions[$j].'Trustee Primary SMTP' -ceq $AdditionalRecipients[$i].PrimarySMTPAddress) -and (-not $RecipientPermissions[$j].'Trustee InitialOrAdditional')) {
            if (($i -lt $InitialRecipientsCount) -and ($InitialRecipientsstring -cnotlike '*' + [Environment]::NewLine + $AdditionalRecipients[$i].PrimarySMTPAddress + ';*')) { $InitialRecipientsString += (($AdditionalRecipients[$i].PrimarySMTPAddress) + ';' + ($RecipientPermissions[$j].'Trustee Recipient Type') + [Environment]::NewLine) }
            if (($RecipientPermissions[$j].'Grantor Recipient Type' -inotlike '*Group') -and ($AdditionalRecipients.PrimarySMTPAddress -cnotcontains $RecipientPermissions[$j].'Grantor Primary SMTP')) {
                $AdditionalRecipients += New-Object PsObject -Property @{ PrimarySMTPAddress = $RecipientPermissions[$j].'Grantor Primary SMTP' }
                $AdditionalRecipientsString += (($RecipientPermissions[$j].'Grantor Primary SMTP') + ';' + ($RecipientPermissions[$j].'Grantor Recipient Type') + ';' + ($RecipientPermissions[$j].'Grantor Environment') + ';' + [Environment]::NewLine)
            }
            if (-not $RecipientPermissions[$j].'Grantor InitialOrAdditional') {
                $AllRecipientsGMLFileEdgeString += [Environment]::NewLine + '    edge' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        source ' + $AdditionalRecipients.primarysmtpaddress.IndexOf($RecipientPermissions[$j].'Grantor Primary SMTP') + [Environment]::NewLine + '        target ' + $i + [Environment]::NewLine + '        label "grants ' + $RecipientPermissions[$j].'Permission' + '"' + [Environment]::NewLine + '    ]'
            }
            if ($i -lt $InitialRecipientsCount) {
                $RecipientPermissions[$j].'Trustee InitialOrAdditional' = 'Initial'
            } else {
                $RecipientPermissions[$j].'Trustee InitialOrAdditional' = 'Additional'
            }
        }
    }
    if ($i -lt $InitialRecipientsCount) {
        $AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id ' + $i + [Environment]::NewLine + '        label "' + $AdditionalRecipients[$i].PrimarySMTPAddress + '"' + [Environment]::NewLine + '        gid -32768' + [Environment]::NewLine + '    ]'
    } else {
        $AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id ' + $i + [Environment]::NewLine + '        label "' + $AdditionalRecipients[$i].PrimarySMTPAddress + '"' + [Environment]::NewLine + '        gid -32767' + [Environment]::NewLine + '    ]'
    }
    if (($i -lt $InitialRecipientsCount) -and ($InitialRecipientsstring -cnotlike ('*' + [Environment]::NewLine + $AdditionalRecipients[$i].PrimarySMTPAddress + ';*'))) { $InitialRecipientsString += (($AdditionalRecipients[$i].PrimarySMTPAddress) + ';unknown (not in permissions export file);unknown (not in permissions export file)' + [Environment]::NewLine) }
}


Write-Host
Write-Host ("Export recipient dependencies to graph file '$AllRecipientsGMLFile'")
$AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id -32768' + [Environment]::NewLine + '        label "Recipients from input file"' + [Environment]::NewLine + '        isGroup 1' + [Environment]::NewLine + '    ]'
$AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id -32767' + [Environment]::NewLine + '        label "Additional recipients"' + [Environment]::NewLine + '        isGroup 1' + [Environment]::NewLine + '    ]'


$AllRecipientsGMLFileString += [Environment]::NewLine + $AllRecipientsGMLFileEdgeString + [Environment]::NewLine + "]"
$AllRecipientsGMLFileString = ($AllRecipientsGMLFileString -replace '&', '&amp;')
$AllRecipientsGMLFileString | Out-File $AllRecipientsGMLFile -force -Encoding utf8


Write-Host
Write-Host ("Export full list of recipients to '$ExportAllRecipientsFile'")
($InitialRecipientsString + $AdditionalRecipientsString) -replace ([Environment]::NewLine + [Environment]::NewLine), ([Environment]::NewLine) | Out-File $ExportAllRecipientsFile -Append -Encoding utf8


Write-Host
Write-Host ("Export only additional recipients to '$ExportAdditionalRecipientsFile'")
'Primary SMTP Address;Recipient Type;Environment' | Out-File $ExportAdditionalRecipientsFile -Encoding utf8
$AdditionalRecipientsString | Sort-Object | Out-File $ExportAdditionalRecipientsFile -Append -Encoding utf8

write-host
Write-Host ("Export modified recipient permissions file to '$RecipientPermissionsFileNew'")
$RecipientPermissions = ($RecipientPermissions | Select-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Permission', 'Folder', 'Grantor InitialOrAdditional', 'Trustee InitialOrAdditional', 'Root cause for additional mailboxes')
foreach ($x in $RecipientPermissions) {
    If (($x.'Grantor InitialOrAdditional' -eq 'Initial') -and ($x.'Trustee InitialOrAdditional' -eq 'Additional')) {
        $x.'Root cause for additional mailboxes' = 'yes'
    } else {
        If (($x.'Grantor InitialOrAdditional' -eq 'Additional') -and ($x.'Trustee InitialOrAdditional' -eq 'Initial')) {
            $x.'Root cause for additional mailboxes' = 'yes'
        } else {
            $x.'Root cause for additional mailboxes' = 'no'
        }
    }
}
$RecipientPermissions | Where-Object { (($_.'Grantor InitialOrAdditional' -ne $null) -or ($_.'Trustee InitialOrAdditional' -ne $null)) } | Export-Csv $RecipientPermissionsFileNew -Delimiter ';' -NoTypeInformation -Force


write-host
Write-Host ("Create summary file '$SummaryFile'")
('{0:000000}' -f $InitialRecipientsCount) + ' initial recipients to migrate.' >> $SummaryFile
('{0:000000}' -f ($AdditionalRecipients.count - $InitialRecipientsCount)) + ' additional recipients to migrate.' >> $SummaryFile
('{0:000000}' -f $AdditionalRecipients.count) + ' total recipients to migrate.' >> $SummaryFile
('{0:000000}' -f ($RecipientPermissions | Where-Object { $_.'Root cause for additional mailboxes' -eq 'yes' } | Measure-Object).count) + ' root cause recipient permissions.' >> $SummaryFile


write-host
Write-Host ('End of script.')
