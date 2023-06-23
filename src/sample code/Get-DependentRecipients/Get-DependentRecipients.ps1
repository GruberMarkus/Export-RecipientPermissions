[CmdletBinding(PositionalBinding = $false)]

param (
    # Input: Primary SMTP addresses of users you want to move. Every line in this file containing the @-sign is imported.
    $InitialRecipientsFile = '.\Input_InitialRecipients.csv',

    # Input: Export file of Export-RecipientPermissions.ps1
    $RecipientPermissionsFile = '.\Input_Export-RecipientPermissions_Output.csv', 

    # Input: List of primary SMTP addresses to ignore (service accounts, for example). Wildcards are not allowed.
    $PrimarySMTPAddressesToIgnore = @('xxx@domain.com', 'yyy@domain.com'),

    # Output: List of all recipients, initial and additional (a.k.a. dependent)
    $ExportAllRecipientsFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AllRecipients.txt',

    # Output: List of initial (a.k.a. dependent) recipients
    $ExportInitialRecipientsFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_InitialRecipients.txt',

    # Output: List of additional (a.k.a. dependent) recipients
    $ExportAdditionalRecipientsFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_AdditionalRecipients.txt',

    # Output: File containing initial and additional recipients, their permission and information why they are dependent recipients
    $ExportRecipientPermissionsFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_Permissions.csv',

    # Output: Graphical representation of the file above in Graph Modeling Language (GML). Can be viewed with yWorks yEd, for example.
    $ExportAllRecipientsGMLFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_GML.gml',

    # Output: Summary file showing number of initial, addtional and total recipients as well as the number of root cause permissions
    $ExportSummaryFile = '.\Get-DependentRecipients_Output\Get-DependentRecipients_Output_Summary.txt',

    # Interval to update the job progress
    [int][ValidateRange(1, [int]::MaxValue)]$UpdateInterval = 1000
)


Clear-Host

Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

Set-Location $PSScriptRoot

if ($PSVersionTable.PSEdition -eq 'desktop') {
    $UTF8Encoding = 'UTF8'
} else {
    $UTF8Encoding = 'UTF8BOM'
}

$InitialRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InitialRecipientsFile)
$RecipientPermissionsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecipientPermissionsFile)
$ExportAllRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAllRecipientsFile)
$ExportAdditionalRecipientsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAdditionalRecipientsFile)
$ExportAllRecipientsGMLFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportAllRecipientsGMLFile)
$ExportRecipientPermissionsFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportRecipientPermissionsFile)
$ExportSummaryFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportSummaryFile)

New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAllRecipientsFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAdditionalRecipientsFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportAllRecipientsGMLFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportRecipientPermissionsFile) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ExportSummaryFile) | Out-Null

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
if (Test-Path $ExportAllRecipientsGMLFile) { (Remove-Item $ExportAllRecipientsGMLFile) }
if (Test-Path $ExportRecipientPermissionsFile) { (Remove-Item $ExportRecipientPermissionsFile) }
if (Test-Path $ExportSummaryFile) { (Remove-Item $ExportSummaryFile) }


Write-Host
Write-Host "Import recipient permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$RecipientPermissionsFile'"
$RecipientPermissions = [system.collections.arraylist]::new()
$RecipientPermissions.AddRange((Import-Csv $RecipientPermissionsFile -Delimiter ';' -Encoding $UTF8Encoding | Sort-Object -Property 'Grantor Primary SMTP'))
$RecipientPermissions | Add-Member -MemberType NoteProperty -Name 'Grantor InitialOrAdditional' -Value $null
$RecipientPermissions | Add-Member -MemberType NoteProperty -Name 'Trustee InitialOrAdditional' -Value $null
$RecipientPermissions | Add-Member -MemberType NoteProperty -Name 'Root cause for additional mailboxes' -Value $null
$RecipientPermissions | Add-Member -MemberType NoteProperty -Name 'GML Source ID' -Value $null
$RecipientPermissions | Add-Member -MemberType NoteProperty -Name 'GML Target ID' -Value $null
$RecipientPermissions | Add-Member -MemberType NoteProperty -Name 'Edge created' -Value $null



Write-Host
Write-Host "Filter recipient permissions @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$a = $RecipientPermissions.count - 1
$b = 0
for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
    if (($b % $UpdateInterval -eq 0) -or ($b -eq $a)) {
        Write-Host (("`b" * 100) + ('  {0:0000000}/{1:0000000} @{2}@' -f $b, $a, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
        if (($b -eq 0) -or ($b -eq $a)) { Write-Host }
    }
    $b++

    # Remove permissions with non-resolvable trustees
    if (-not $RecipientPermissions[$x].'Trustee Primary SMTP') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove permissions with primary SMTP addresses to ignore
    if (($RecipientPermissions[$x].'Grantor Primary SMTP' -iin $PrimarySMTPAddressesToIgnore) -or ($RecipientPermissions[$x].'Trustee Primary SMTP' -iin $PrimarySMTPAddressesToIgnore)) {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove permissions where grantor and trustee are the same
    if ($RecipientPermissions[$x].'Grantor Primary SMTP' -eq $RecipientPermissions[$x].'Trustee Primary SMTP') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove permissions where grantor or trustee is already in the cloud
    if (($RecipientPermissions[$x].'Grantor Environment' -ieq 'Cloud') -or ($RecipientPermissions[$x].'Trustee Environment' -ieq 'Cloud')) {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove permissions with groups as trustees (optional)
    #if ($RecipientPermissions[$x].'Trustee Recipient Type' -ilike "*group") {
    #    $RecipientPermissions.RemoveAt($x)
    #    $x--
    #    continue
    #}


    # See https://docs.microsoft.com/en-us/Exchange/permissions for details on which permissions work cross-premises
    # Also, test on your own with permissions set before and after migrating a mailbox


    # Remove permissions containing "FullAccess" permission if this permission works cross-premises
    if ($RecipientPermissions[$x].'Permission' -ieq 'FullAccess') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove permissions containing "SendOnBehalf" permission if this permission works cross-premises
    if ($RecipientPermissions[$x].'Permission' -ieq 'SendOnBehalf') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove folder permissions if this permission works cross-premises
    if ($RecipientPermissions[$x].'Folder') {
        $RecipientPermissions.RemoveAt($x)
        $x--
        continue
    }


    # Remove lines containing "SendAs" permission if this permission works cross-premises
    #if ($RecipientPermissions[$x].'Permission' -ieq 'SendAs') {
    #    $RecipientPermissions.RemoveAt($x)
    #   $x--
    #   continue
    #}


    # Remove lines with "ManagedBy" permission if this permission works cross-premises
    #if ($RecipientPermissions[$x].'Permission' -ieq 'ManagedBy') {
    #    $RecipientPermissions.remove($x)
    #    $x--
    #    continue
    #}


}

$RecipientPermissions.TrimToSize()
Write-Host ('  {0:0000000} permissions to consider' -f $RecipientPermissions.count)
if ($RecipientPermissions.count -le 0) {
    Write-Host
    Write-Host 'Nothing to do'
    exit
}


Write-Host
Write-Host "Create lookup tables @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  Unique IDs for GML node IDs @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$UniqueSmtps = (@() + $RecipientPermissions.'Grantor Primary SMTP' + $RecipientPermissions.'Trustee Primary SMTP')
$GmlIds = @{}
for ($x = 0; $x -lt $UniqueSmtps.count; $x++) {
    $GmlIds."$($uniquesmtps[$x])" = $x
}

Write-Host "  RecipientPermissions indexes forech SMTP address @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$SmtpToRecipientpermissionsIndex = @{}

for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
    if (($x % $UpdateInterval -eq 0) -or ($x -eq ($RecipientPermissions.count - 1))) {
        Write-Host (("`b" * 100) + ('    {0:0000000}/{1:0000000} @{2}@' -f $x, ($RecipientPermissions.count - 1), $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
        if (($x -eq 0) -or ($x -eq ($RecipientPermissions.count - 1))) { Write-Host }
    }

    if (-not $SmtpToRecipientpermissionsIndex."$($RecipientPermissions[$x].'Grantor Primary SMTP')") {
        $SmtpToRecipientpermissionsIndex."$($RecipientPermissions[$x].'Grantor Primary SMTP')" = [system.collections.arraylist]::new(1)
    }
    $null = $SmtpToRecipientpermissionsIndex."$($RecipientPermissions[$x].'Grantor Primary SMTP')".add($x)

    if (-not $SmtpToRecipientpermissionsIndex."$($RecipientPermissions[$x].'Trustee Primary SMTP')") {
        $SmtpToRecipientpermissionsIndex."$($RecipientPermissions[$x].'Trustee Primary SMTP')" = [system.collections.arraylist]::new(1)
    }
    $null = $SmtpToRecipientpermissionsIndex."$($RecipientPermissions[$x].'Trustee Primary SMTP')".add($x)

}


Write-Host
Write-Host "Import initial recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$InitialRecipientsFile' "
$InitialRecipients = @(Get-Content $InitialRecipientsFile | Where-Object { $_ -like '*@*' } | Sort-Object -Unique)
Write-Host '  Match case sensitivity with permissions file'
for ($x = 0; $x -lt $UniqueSmtps.count; $x++) {
    if (($x % $UpdateInterval -eq 0) -or ($x -eq ($UniqueSmtps.count - 1))) {
        Write-Host (("`b" * 100) + ('    {0:0000000}/{1:0000000} @{2}@' -f $x, ($UniqueSmtps.count - 1), $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
        if (($x -eq 0) -or ($x -eq ($UniqueSmtps.count - 1))) { Write-Host }
    }

    for ($y = 0; $y -lt $InitialRecipients.count; $y++) {
        if ($UniqueSmtps[$x] -ieq $InitialRecipients[$y]) {
            $InitialRecipients[$y] = $UniqueSmtps[$x]
            break
        } elseif ($UniqueSmtps[$x] -ieq $InitialRecipients[$y]) {
            $InitialRecipients[$y] = $UniqueSmtps[$x]
            break
        }
    }
}

$InitialRecipientsCount = $InitialRecipients.count
$AdditionalRecipients = $InitialRecipients


Write-Host
Write-Host "Add metadata to each permission @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

for ($i = 0; $i -lt $RecipientPermissions.count; $i++) {
    if (($i % $UpdateInterval -eq 0) -or ($i -eq ($RecipientPermissions.count - 1))) {
        Write-Host (("`b" * 100) + ('  {0:0000000}/{1:0000000} @{2}@' -f $i, ($RecipientPermissions.count - 1), $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
        if (($i -eq 0) -or ($i -eq ($RecipientPermissions.count - 1))) { Write-Host }
    }

    # Grantor InitialOrAdditional
    if ($RecipientPermissions[$i].'Grantor Primary SMTP' -in $InitialRecipients) {
        $RecipientPermissions[$i].'Grantor InitialOrAdditional' = 'Initial'
    }


    # Grantor InitialOrAdditional
    if ($RecipientPermissions[$i].'Trustee Primary SMTP' -in $InitialRecipients) {
        $RecipientPermissions[$i].'Trustee InitialOrAdditional' = 'Initial'
    }

    # GML Source IDs
    $RecipientPermissions[$i].'GML Source ID' = $GmlIds."$($RecipientPermissions[$i].'Grantor Primary SMTP')"
    $RecipientPermissions[$i].'GML Target ID' = $GmlIds."$($RecipientPermissions[$i].'Trustee Primary SMTP')"
}


Write-Host
Write-Host "Get GML data from permissions and add additional metadata, recipient by recipient @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host '  This may take long, depending on number of permissions and distribution across recipients'

$NodesCreated = @()
$EdgesCreated = @()
$ExportAllRecipientsGMLFileString = [System.Collections.ArrayList]::new($RecipientPermissions.count * 2)
$null = $ExportAllRecipientsGMLFileString.add([String]('graph' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'hierarchic 1' + [Environment]::NewLine + 'directed 1'))
$ExportAllRecipientsGMLFileEdgeString = [System.Collections.ArrayList]::new($RecipientPermissions.count)
$RootCausePermissionsCount = 0

for ($x = 0; $x -lt $AdditionalRecipients.count; $x++) {
    if (($x % $UpdateInterval -eq 0) -or ($x -eq ($AdditionalRecipients.count - 1))) {
        Write-Host (("`b" * 100) + ('  {0:0000000}/{1:0000000} @{2}@' -f $x, ($AdditionalRecipients.count - 1), $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
        if (($x -eq 0) -or ($x -eq ($AdditionalRecipients.count - 1))) { Write-Host }
    }

    foreach ($i in $SmtpToRecipientpermissionsIndex."$($AdditionalRecipients[$x])") {
        if (-not $RecipientPermissions[$i].'Edge Created') {
            if (
                ($RecipientPermissions[$i].'Grantor InitialOrAdditional' -eq 'Initial') -or
                ($RecipientPermissions[$i].'Trustee InitialOrAdditional' -eq 'Initial') -or
                ($RecipientPermissions[$i].'Grantor Primary SMTP' -in $AdditionalRecipients) -or
                ($RecipientPermissions[$i].'Trustee Primary SMTP' -in $AdditionalRecipients)
            ) {

                # GML node for Grantor
                if ($RecipientPermissions[$i].'Grantor Primary SMTP' -notin $NodesCreated) {
                    if ($RecipientPermissions[$i].'Grantor Primary SMTP' -in $InitialRecipients) {
                        $null = $ExportAllRecipientsGMLFileString.add(('node' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'id ' + $RecipientPermissions[$i].'GML Source ID' + [Environment]::NewLine + 'label "' + [System.Net.WebUtility]::HtmlEncode($RecipientPermissions[$i].'Grantor Primary SMTP') + '"' + [Environment]::NewLine + 'gid -32768' + [Environment]::NewLine + ']'))
                    } else {
                        $null = $ExportAllRecipientsGMLFileString.add(('node' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'id ' + $RecipientPermissions[$i].'GML Source ID' + [Environment]::NewLine + 'label "' + [System.Net.WebUtility]::HtmlEncode($RecipientPermissions[$i].'Grantor Primary SMTP') + '"' + [Environment]::NewLine + 'gid -32767' + [Environment]::NewLine + ']'))
                    }

                    if ($RecipientPermissions[$i].'Grantor Primary SMTP' -notin $AdditionalRecipients) {
                        $AdditionalRecipients += $RecipientPermissions[$i].'Grantor Primary SMTP'
                    }

                    $NodesCreated += $RecipientPermissions[$i].'Grantor Primary SMTP'
                }


                # GML node for trustee
                if ($RecipientPermissions[$i].'Trustee Primary SMTP' -notin $NodesCreated) {
                    if ($RecipientPermissions[$i].'Trustee Primary SMTP' -in $InitialRecipients) {
                        $null = $ExportAllRecipientsGMLFileString.add(('node' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'id ' + $RecipientPermissions[$i].'GML Target ID' + [Environment]::NewLine + 'label "' + [System.Net.WebUtility]::HtmlEncode($RecipientPermissions[$i].'Trustee Primary SMTP') + '"' + [Environment]::NewLine + 'gid -32768' + [Environment]::NewLine + ']'))
                    } else {
                        $null = $ExportAllRecipientsGMLFileString.add(('node' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'id ' + $RecipientPermissions[$i].'GML Target ID' + [Environment]::NewLine + 'label "' + [System.Net.WebUtility]::HtmlEncode($RecipientPermissions[$i].'Trustee Primary SMTP') + '"' + [Environment]::NewLine + 'gid -32767' + [Environment]::NewLine + ']'))
                    }

                    if ($RecipientPermissions[$i].'Trustee Primary SMTP' -notin $AdditionalRecipients) {
                        $AdditionalRecipients += $RecipientPermissions[$i].'Trustee Primary SMTP'
                    }
                    
                    $NodesCreated += $RecipientPermissions[$i].'Trustee Primary SMTP'
                }


                # GML edge for permission
                $null = $ExportAllRecipientsGMLFileEdgeString.add(('edge' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'source ' + $RecipientPermissions[$i].'GML Source ID' + [Environment]::NewLine + 'target ' + $RecipientPermissions[$i].'GML Target ID' + [Environment]::NewLine + 'label "grants ' + [System.Net.WebUtility]::HtmlEncode($RecipientPermissions[$i].'Permission') + $(if ($RecipientPermissions[$i].'Folder') { " on $([System.Net.WebUtility]::HtmlEncode($RecipientPermissions[$i].'Folder'))" }) + '"' + [Environment]::NewLine + ']'))
                $RecipientPermissions[$i].'Edge Created' = 'yes'


                # Grantor/Trustee InitialOrAdditional (initial has already been set before)
                if ($RecipientPermissions[$i].'Grantor Primary SMTP' -notin $InitialRecipients) {
                    $RecipientPermissions[$i].'Grantor InitialOrAdditional' = 'Additional'
                }

                if ($RecipientPermissions[$i].'Trustee Primary SMTP' -notin $InitialRecipients) {
                    $RecipientPermissions[$i].'Trustee InitialOrAdditional' = 'Additional'
                }

                # Root cause for additional mailboxes
                if ($RecipientPermissions[$i].'Grantor InitialOrAdditional' -ne $RecipientPermissions[$i].'Trustee InitialOrAdditional') {
                    $RecipientPermissions[$i].'Root cause for additional mailboxes' = 'Yes'
                    $RootCausePermissionsCount++
                } else {
                    $RecipientPermissions[$i].'Root cause for additional mailboxes' = 'No'
                }
            }
        }
    }
}


Write-Host
Write-Host "Create GML (Graph Modeling Language) file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$ExportAllRecipientsGMLFile'"
$null = $ExportAllRecipientsGMLFileString.add(('node' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'id -32768' + [Environment]::NewLine + 'label "Initial Recipients"' + [Environment]::NewLine + 'isGroup 1' + [Environment]::NewLine + ']'))
$null = $ExportAllRecipientsGMLFileString.add(('node' + [Environment]::NewLine + '[' + [Environment]::NewLine + 'id -32767' + [Environment]::NewLine + 'label "Additional recipients"' + [Environment]::NewLine + 'isGroup 1' + [Environment]::NewLine + ']'))
$null = $ExportAllRecipientsGMLFileString.addrange($ExportAllRecipientsGMLFileEdgeString)
$null = $ExportAllRecipientsGMLFileString.add((']'))
$ExportAllRecipientsGMLFileEdgeString = $null
[IO.File]::WriteAllLines($ExportAllRecipientsGMLFile, $ExportAllRecipientsGMLFileString)


Write-Host
Write-Host "Export modified recipient permissions file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host '  Filter relevant permissions'
$a = $RecipientPermissions.Count - 1
$b = 0
for ($x = 0; $x -lt $RecipientPermissions.Count; $x++) {
    if (($b % $UpdateInterval -eq 0) -or ($b -eq $a)) {
        Write-Host (("`b" * 100) + ('    {0:0000000}/{1:0000000} @{2}@' -f $b, $a, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline
        if (($b -eq 0) -or ($b -eq $a)) { Write-Host }
    }
    $b++

    if ((-not $RecipientPermissions[$x].'Grantor InitialOrAdditional') -or (-not $RecipientPermissions[$x].'Trustee InitialOrAdditional')) {
        $RecipientPermissions.RemoveAt($x)
        $x--
    }
}
Write-Host "  '$ExportRecipientPermissionsFile'"
$RecipientPermissions | Select-Object * -ExcludeProperty 'GML Source ID', 'GML Target ID', 'Edge created' | Export-Csv $ExportRecipientPermissionsFile -Delimiter ';' -NoTypeInformation -Force -Encoding $UTF8Encoding


Write-Host
Write-Host "Export list of initial recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$ExportInitialRecipientsFile'"
$InitialRecipients | Out-File $ExportInitialRecipientsFile -Encoding $UTF8Encoding -Force


Write-Host
Write-Host "Export list of additional recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$ExportAdditionalRecipientsFile'"
$AdditionalRecipients[$($InitialRecipients.Count)..$($AdditionalRecipients.count)] | Out-File $ExportAdditionalRecipientsFile -Encoding $UTF8Encoding -Force


Write-Host
Write-Host "Export list of all recipients @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$ExportAllRecipientsFile'"
$AdditionalRecipients | Out-File $ExportAllRecipientsFile -Encoding $UTF8Encoding -Force


Write-Host
Write-Host "Create summary @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  '$ExportSummaryFile'"
@'
{0:0000000} initial recipients
{1:0000000} root cause permissions
{2:0000000} additional recipients
{3:0000000} total recipients
'@ -f $InitialRecipients.count, $RootCausePermissionsCount, ($AdditionalRecipients.Count - $InitialRecipients.count), $AdditionalRecipients.count | Out-File $ExportSummaryFile -Encoding $UTF8Encoding -Force


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
