# Compare two result files from Export-RecipientPermissions.ps1
#
# Changes are marked in the column 'Change' with
#   'Deleted' if a line exists in the old file but not in the new one
#   'New' if a line exists in the new file but not in the old one
#   'Unchanged' if a line exists as well in the old file as in the new file
#
# Optionally display changes on screen
#
# Optionally export changes to CSV file


[CmdletBinding(PositionalBinding = $false)]


Param(
    # Path to the CSV file from the older run of Export-RecipientPermissions
    $oldCsv = '.\Export-RecipientPermissions_Output_old.csv',

    # Path to the CSV file from the newer run of Export-RecipientPermissions
    $newCsv = '.\Export-RecipientPermissions_Output_new.csv',

    # Display results on screen before creating file showing changes
    $DisplayResults = $true,

    # Path for export file showing changes
    # Set to '' or $null to not create this file
    $ChangeFile = '.\comparison.csv'
)


Clear-Host


Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Set-Location $PSScriptRoot

if ($PSVersionTable.PSEdition -ieq 'desktop') {
    $UTF8Encoding = 'UTF8'
} else {
    $UTF8Encoding = 'UTF8BOM'
}


Write-Host
Write-Host "Import old CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$oldLines = Import-Csv $oldCsv -Delimiter ';' -Encoding $UTF8Encoding


Write-Host
Write-Host "Import new CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$newLines = Import-Csv $newCsv -Delimiter ';' -Encoding $UTF8Encoding


Write-Host
Write-Host "Compare CSV files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  Create compared dataset @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$Dataset = Compare-Object -ReferenceObject $oldLines -DifferenceObject $newLines -Property $newLines[0].psobject.properties.name -IncludeEqual -PassThru
$Dataset | Add-Member -MemberType NoteProperty -Name 'Change' -Value $null

Write-Host "  Modify and sort dataset @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
foreach ($DatasetObject in $Dataset) {
    if ($DatasetObject.sideindicator -ieq '<=') {
        $DatasetObject.'Change' = 'Deleted'
    } elseif ($DatasetObject.sideindicator -ieq '=>') {
        $DatasetObject.'Change' = 'New'
    } else {
        $DatasetObject.'Change' = 'Unchanged'
    }
}

$Dataset = $Dataset | Select-Object * -ExcludeProperty 'SideIndicator'

$Dataset = $Dataset | Sort-Object -Property 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Folder', 'Permission', 'Trustee Original Identity', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Change'


Write-Host
Write-Host "Display results @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($DisplayResults) {
    $GrantorPrimarySmtpOrder = [system.collections.arraylist]::new($Dataset.'Grantor Primary SMTP')
    $GrantorPrimarySmtpReverseOrder = [system.collections.arraylist]::new($GrantorPrimarySmtpOrder.count)
    $GrantorPrimarySmtpList = [ordered]@{}

    for ($x = $GrantorPrimarySmtpOrder.count; $x -ge 0; $x--) {
        $null = $GrantorPrimarySmtpReverseOrder.add($GrantorPrimarySmtpOrder[$x])
        $GrantorPrimarySmtpList.'$($GrantorPrimarySmtpOrder[$x])' = $null
    }

    foreach ($GrantorPrimarySmtp in ($GrantorPrimarySmtpList.keys)) {
        Write-Host "  $($GrantorPrimarySmtp)"
        foreach ($DatasetObject in $Dataset[$($GrantorPrimarySmtpOrder.IndexOf($GrantorPrimarySmtp))..$($GrantorPrimarySmtpReverseOrder.count - 1 - $GrantorPrimarySmtpReverseOrder.IndexOf($GrantorPrimarySmtp))]) {
            if ($DatasetObject.Change -eq 'Deleted') {
                Write-Host ("    Deleted: '$($DatasetObject.'Trustee Original Identity')' (E-Mail '$($DatasetObject.'Trustee Primary SMTP')') no longer has the '$($DatasetObject.'Permission')' right" + $(if ($DatasetObject.'Folder') { " on folder '$($DatasetObject.'Folder')'" }))
            } elseif ($DatasetObject.change -eq 'New') {
                Write-Host ("    New: '$($DatasetObject.'Trustee Original Identity')' (E-Mail '$($DatasetObject.'Trustee Primary SMTP')) now has the '$($DatasetObject.'Permission')' right" + $(if ($DatasetObject.'Folder') { " on folder '$($DatasetObject.'Folder')'" }))
            } else {
                Write-Host ("    Unchanged: '$($DatasetObject.'Trustee Original Identity')' (E-Mail '$($DatasetObject.'Trustee Primary SMTP')') still has the '$($DatasetObject.'Permission')' right" + $(if ($DatasetObject.'Folder') { " on folder '$($DatasetObject.'Folder')'" }))
            }
        }
    }
} else {
    Write-Host '  Not required with current settings'
}


Write-Host
Write-Host "Create export file showing changes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($ChangeFile) {
    Write-Host "  '$($ChangeFile)'"
    $Dataset | Export-Csv -Path $ChangeFile -Delimiter ';' -Encoding $UTF8Encoding -NoTypeInformation -Force
} else {
    Write-Host '  Not required with current configuration settings'
}


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"