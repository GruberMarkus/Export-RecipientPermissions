param (
    $InitialRecipientsFile = ".\input.csv", # Primary SMTP addresses of users you want to move. Every line in this file containing the @-sign is imported.
    $RecipientPermissionsFile = ".\Export-RecipientPermissions_Output.csv", # Default output file of Export-RecipientPermissions.ps1.
    $PrimarySMTPAddressesToIgnore = @("xxx@domain.com", "yyy@domain.com"), #List of primary SMTP addresses to ignore (service account, for example). Wildcards are not allowed.
    $ExportAllRecipientsFile = ".\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AllRecipients.csv",
    $ExportAdditionalRecipientsFile = ".\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AdditionalRecipients.csv",
    $AllRecipientsGMLFile = ".\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AllRecipients.gml",
    $RecipientPermissionsFileNew = ".\Get-DependentRecipients_Output\Export-RecipientPermissions_Output_Modified.csv",
    $InitialRecipientsFileNew = ".\Get-DependentRecipients_Output\Get-DependentRecipients_Output_OriginalInput.csv",
    $SummaryFile = ".\Get-DependentRecipients_Output\Get-DependentRecipients_Output_Summary.txt"
)

clear-host

Set-Location $PSScriptRoot

$InitialRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InitialRecipientsFile)
$InitialRecipientsFileNew = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InitialRecipientsFileNew)
$RecipientPermissionsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecipientPermissionsFile)
$ExportAllRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAllRecipientsFile)
$ExportAdditionalRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAdditionalRecipientsFile)
$AllRecipientsGMLFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AllRecipientsGMLFile)
$RecipientPermissionsFileNew = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecipientPermissionsFileNew)

New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAllRecipientsFile) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAdditionalRecipientsFile) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $AllRecipientsGMLFile) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $RecipientPermissionsFileNew) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $InitialRecipientsFileNew) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $SummaryFile) | out-null

if (-not (Test-Path $RecipientPermissionsFile)) {
    write-host "Recipient permissions file '$RecipientPermissionsFile' not found, exiting."
    break
}

if (Test-Path $ExportAllRecipientsFile) {(Remove-Item $ExportAllRecipientsFile)}
if (Test-Path $ExportAdditionalRecipientsFile) {(Remove-Item $ExportAdditionalRecipientsFile)}
if (Test-Path $AllRecipientsGMLFile) {(Remove-Item $AllRecipientsGMLFile)}
if (Test-Path $RecipientPermissionsFileNew) {(Remove-Item $RecipientPermissionsFileNew)}
if (Test-Path $SummaryFile) {(Remove-Item $SummaryFile)}
if (Test-Path $InitialRecipientsFileNew) {(Remove-Item $InitialRecipientsFileNew)}

# Import and remove
write-host "Importing recipient permissions from '$RecipientPermissionsFile'" -nonewline
$RecipientPermissions = Import-Csv $RecipientPermissionsFile -delimiter ';'
write-host "."

write-host ("Filtering " + $RecipientPermissions.count +" recipient permissions") -nonewline
for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
    # Lines to filter are marked with an empty 'Grantor Primary SMTP'
    # Remove empty trustees (empty grantors are filtered later via select)
    if ($RecipientPermissions[$x].'Trustee Primary SMTP' -eq '') {
        $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
        continue
    }

    # Remove lines where grantor or trustee is already in the cloud
    #if (($RecipientPermissions[$x].'Grantor Environment' -eq 'Cloud') -or ($RecipientPermissions[$x].'Trustee Environment' -eq 'Cloud')) {
    #    $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
    #    continue
    #}

    # Remove lines containing "FullAccess" permission (they work cross premises)
    #if ($RecipientPermissions[$x].'Permission' -Match 'FullAccess') {
    #    $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
    #    continue
    #}

    # Remove lines containing "SendOnBehalf" permission (they work cross premises)
    # Does not work yet, but is on the O365 Roadmap
    #if ($RecipientPermissions[$x].'Permission' -eq 'SendOnBehalf') {
    #    $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
    #    continue
    #}

    # Remove lines containing "SendAs" permission (they work cross premises)
    # Does not work yet, but may come with a future O365 update
    #if ($RecipientPermissions[$x].'Permission' -eq 'SendAs') {
    #    $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
    #    continue
    #}

    # Remove lines with ""ManagedBy"" permission
    #if ($RecipientPermissions[$x].'Permission' -eq 'ManagedBy') {
    #    $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
    #    continue
    #}


    # Remove lines with primary SMTP addresses to ignore
    if (($RecipientPermissions[$x].'Grantor Primary SMTP' -In $PrimarySMTPAddressesToIgnore) -or ($RecipientPermissions[$x].'Trustee Primary SMTP' -In $PrimarySMTPAddressesToIgnore)) {
        $RecipientPermissions[$x].'Grantor Primary SMTP' = ''
        continue
    }
}
write-host "."

# Select all columns and create additional columns
write-host "Number of permissions to consider: " -NoNewline
$RecipientPermissions = ($RecipientPermissions | where-object {$_.'Grantor Primary SMTP' -ne ''})
$RecipientPermissions = ($RecipientPermissions | select-object *, 'Grantor InitialOrAdditional', 'Trustee InitialOrAdditional')
write-host $RecipientPermissions.count


write-host "Importing initial recipients from '$InitialRecipientsFile'" -NoNewline
$InitialRecipients = @(import-csv -header PrimarySMTPAddress $InitialRecipientsFile | select-object PrimarySMTPAddress | Where-Object {$_.PrimarySMTPAddress.contains("@")} | Sort-Object PrimarySMTPAddress -unique)
$InitialRecipientsCount = $InitialRecipients.count
$AdditionalRecipients = $InitialRecipients | select-object *, OU
write-host "."

$AllRecipientsGMLFileString = 'graph' + [Environment]::NewLine + '[' + [Environment]::NewLine + '    hierarchic 1' + [Environment]::NewLine + '    directed 1'
$AllRecipientsGMLFileEdgeString = $null
$AdditionalRecipientsString = $null
$InitialRecipientsString = "Primary SMTP Address;Recipient Type;Environment;OU" + [Environment]::NewLine

# $AdditionalRecipients will be expanded step by step until no new entries are added
$OUNumber = -32767
$OUs = @{}
write-host
for ($i = 0; $i -lt $AdditionalRecipients.count; $i++) {
    write-host ("{0:000000}/{1:000000}, {2}" -f ($i+1), $AdditionalRecipients.count, $AdditionalRecipients[$i].PrimarySMTPAddress)
    for ($j=0; $j -lt $RecipientPermissions.count; $j++) {
        # Where is the current recipient a grantor and who are the trustees?
        if ($RecipientPermissions[$j].'Grantor InitialOrAdditional' -eq $null) {
            if (($RecipientPermissions[$j].'Grantor Primary SMTP' -eq $AdditionalRecipients[$i].PrimarySMTPAddress)) {
                if (($i -lt $InitialRecipientsCount) -and ($InitialRecipientsstring -notlike '*' + [Environment]::NewLine + $AdditionalRecipients[$i].PrimarySMTPAddress + ';*')) {$InitialRecipientsString += (($AdditionalRecipients[$i].PrimarySMTPAddress) + ';' + ($RecipientPermissions[$j].'Grantor Recipient Type') + ';' + ($RecipientPermissions[$j].'Grantor OU') + [Environment]::NewLine)}
                if ($AdditionalRecipients.PrimarySMTPAddress -notcontains $RecipientPermissions[$j].'Trustee Primary SMTP') {
                    $AdditionalRecipients += New-Object PsObject -Property @{ PrimarySMTPAddress = $RecipientPermissions[$j].'Trustee Primary SMTP'; OU = $RecipientPermissions[$j].'Trustee OU' }
                    $AdditionalRecipientsString += (($RecipientPermissions[$j].'Trustee Primary SMTP') + ';' + ($RecipientPermissions[$j].'Trustee Recipient Type') + ';' + ($RecipientPermissions[$j].'Trustee Environment') + ';' + ($RecipientPermissions[$j].'Trustee OU') + [Environment]::NewLine)
                }
                if ($RecipientPermissions[$j].'Trustee InitialOrAdditional' -eq $null) {
                    $AllRecipientsGMLFileEdgeString += [Environment]::NewLine + '    edge' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        source ' + $i + [Environment]::NewLine + '        target ' + $AdditionalRecipients.primarysmtpaddress.IndexOf($RecipientPermissions[$j].'Trustee Primary SMTP') + [Environment]::NewLine + '        label "grants ' + $RecipientPermissions[$j].'Permission' + '"' + [Environment]::NewLine +  '    ]'
                }
                if ($i -lt $InitialRecipientsCount) {
                    $RecipientPermissions[$j].'Grantor InitialOrAdditional' = 'Initial'
                } else {
                    $RecipientPermissions[$j].'Grantor InitialOrAdditional' = 'Additional'
                    if (-not $OUs.ContainsKey($RecipientPermissions[$j].'Grantor OU')) {
                        $OUs.add($RecipientPermissions[$j].'Grantor OU', $OUNumber)
                        $OUNumber++
                    }
                }
            }
        }

        # Where is the current recipient a trustee and who is the grantor?
        if ($RecipientPermissions[$j].'Trustee Primary SMTP' -eq $RecipientPermissions[$j].'Grantor Primary SMTP') {
            $RecipientPermissions[$j].'Trustee InitialOrAdditional' = $RecipientPermissions[$j].'Grantor InitialOrAdditional'
        }
        if ($RecipientPermissions[$j].'Trustee InitialOrAdditional' -eq $null) {
            if (($RecipientPermissions[$j].'Trustee Primary SMTP' -eq $AdditionalRecipients[$i].PrimarySMTPAddress)) {
                      if (($i -lt $InitialRecipientsCount) -and ($InitialRecipientsstring -notlike '*' + [Environment]::NewLine + $AdditionalRecipients[$i].PrimarySMTPAddress + ';*')) {$InitialRecipientsString += (($AdditionalRecipients[$i].PrimarySMTPAddress) + ';' + ($RecipientPermissions[$j].'Trustee Recipient Type') + ';' + ($RecipientPermissions[$j].'Trustee OU') + [Environment]::NewLine)}
                if (($RecipientPermissions[$j].'Grantor Recipient Type' -notmatch "Group") -and ($AdditionalRecipients.PrimarySMTPAddress -notcontains $RecipientPermissions[$j].'Grantor Primary SMTP')) {
                    $AdditionalRecipients += New-Object PsObject -Property @{ PrimarySMTPAddress = $RecipientPermissions[$j].'Grantor Primary SMTP'; OU = $RecipientPermissions[$j].'Grantor OU' }
                    $AdditionalRecipientsString += (($RecipientPermissions[$j].'Grantor Primary SMTP') + ';' + ($RecipientPermissions[$j].'Grantor Recipient Type') + ';' + ($RecipientPermissions[$j].'Grantor Environment') + ';' + ($RecipientPermissions[$j].'Grantor OU') + [Environment]::NewLine)
                }
                if ($RecipientPermissions[$j].'Grantor InitialOrAdditional' -eq $null) {
                    $AllRecipientsGMLFileEdgeString += [Environment]::NewLine + '    edge' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        source ' + $AdditionalRecipients.primarysmtpaddress.IndexOf($RecipientPermissions[$j].'Grantor Primary SMTP') + [Environment]::NewLine + '        target ' + $i + [Environment]::NewLine + '        label "grants ' + $RecipientPermissions[$j].'Permission' + '"' + [Environment]::NewLine + '    ]'
                }
                if ($i -lt $InitialRecipientsCount) {
                    $RecipientPermissions[$j].'Trustee InitialOrAdditional' = 'Initial'
                } else {
                    $RecipientPermissions[$j].'Trustee InitialOrAdditional' = 'Additional'
                    if (-not $OUs.ContainsKey($RecipientPermissions[$j].'Trustee OU')) {
                        $OUs.add($RecipientPermissions[$j].'Trustee OU', $OUNumber)
                        $OUNumber++
                    }
                }
            }
        }
    }
    if ($i -lt $InitialRecipientsCount) {
        $AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id ' + $i + [Environment]::NewLine + '        label "' + $AdditionalRecipients[$i].PrimarySMTPAddress + '"' + [Environment]::NewLine + '        gid -32768' + [Environment]::NewLine + '    ]'
    } else {
        $AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id ' + $i + [Environment]::NewLine + '        label "' + $AdditionalRecipients[$i].PrimarySMTPAddress + '"' + [Environment]::NewLine + '        gid ' + $OUs.get_item($AdditionalRecipients[$i].OU) + [Environment]::NewLine + '    ]'
    }
    if (($i -lt $InitialRecipientsCount) -and ($InitialRecipientsstring -notlike '*' + [Environment]::NewLine + $AdditionalRecipients[$i].PrimarySMTPAddress + ';*')) {$InitialRecipientsString += (($AdditionalRecipients[$i].PrimarySMTPAddress) + ';unknown (not in permissions export file);unknown (not in permissions export file);unknown (not in permissions export file)'  + [Environment]::NewLine)}
}
write-host
write-host ('Exporting recipient dependencies to graph file ''' + $AllRecipientsGMLFile + '''') -NoNewline
$AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id -32768' + [Environment]::NewLine + '        label "Recipients from input file"' + [Environment]::NewLine + '        isGroup 1' + [Environment]::NewLine + '    ]'

$y = import-csv ".\ous.csv" -Delimiter ";"
foreach ($x in $OUs.GetEnumerator()) {
    if ($y.OU.IndexOf($x.name) -eq -1) {
        $AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id ' + $x.value + [Environment]::NewLine + '        label "' + $x.name + '"' + [Environment]::NewLine + '        isGroup 1' + [Environment]::NewLine + '    ]'
    } else {
        $AllRecipientsGMLFileString += [Environment]::NewLine + '    node' + [Environment]::NewLine + '    [' + [Environment]::NewLine + '        id ' + $x.value + [Environment]::NewLine + '        label "' + $y[$y.OU.IndexOf($x.name)].'Friendly Name' + '"' + [Environment]::NewLine + '        isGroup 1' + [Environment]::NewLine + '    ]'
    }
}


$AllRecipientsGMLFileString += [Environment]::NewLine + $AllRecipientsGMLFileEdgeString
$AllRecipientsGMLFileString = ($AllRecipientsGMLFileString -replace '&', '&amp;')
#$AllRecipientsGMLFileString = (($AllRecipientsGMLFileString -replace 'gid xxx', ("gid {0}" -f ($AdditionalRecipients.count))) + [Environment]::NewLine + "]")
$AllRecipientsGMLFileString | out-file $AllRecipientsGMLFile -Append -encoding "Default"
write-host "."

write-host ('Exporting full list of recipients to ''' + $ExportAllRecipientsFile + '''') -NoNewline
($InitialRecipientsString + $AdditionalRecipientsString) -replace ([Environment]::NewLine + [Environment]::NewLine), ([Environment]::NewLine) | out-file $ExportAllRecipientsFile -append
write-host "."

write-host ('Exporting only additional recipients to ''' + $ExportAdditionalRecipientsFile + '''') -NoNewline
"Primary SMTP Address;Recipient Type;Environment;OU" | out-file $ExportAdditionalRecipientsFile
$AdditionalRecipientsString | sort-object | out-file $ExportAdditionalRecipientsFile -append
write-host "."

write-host ('Exporting modified recipient permissions file to ''' + $RecipientPermissionsFileNew + '''') -NoNewline
$RecipientPermissions = ($RecipientPermissions | select-object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Grantor OU', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee OU', 'Permission', 'Folder', 'Grantor InitialOrAdditional', 'Trustee InitialOrAdditional', 'Root cause for additional mailboxes')
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
$RecipientPermissions | where-object {(($_.'Grantor InitialOrAdditional' -ne $null) -or ($_.'Trustee InitialOrAdditional' -ne $null))} | export-csv $RecipientPermissionsFileNew -delimiter ";" -NoTypeInformation -Force
write-host "."

write-host ('Creating summary file ''' + $SummaryFile + '''') -NoNewline
("{0:000000}" -f $InitialRecipientsCount) + ' initial recipients to migrate.' >> $SummaryFile
("{0:000000}" -f ($AdditionalRecipients.count - $InitialRecipientsCount)) + ' additional recipients to migrate.' >> $SummaryFile
("{0:000000}" -f $AdditionalRecipients.count) + ' total recipients to migrate.' >> $SummaryFile
("{0:000000}" -f ($RecipientPermissions | where-object {$_.'Root cause for additional mailboxes' -eq 'yes'} | measure-object).count) + ' root cause recipient permissions.' >> $SummaryFile
write-host "."

write-host ('Copying original input file to ''' + $InitialRecipientsFileNew + '''') -NoNewline
copy-item $InitialRecipientsFile $InitialRecipientsFileNew
write-host "."


write-host ('End of script.')