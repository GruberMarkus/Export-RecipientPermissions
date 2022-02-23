# Compare two result files from Export-RecipientPermissions.ps1
# to see which permissions have changed over time
#
# Changes are marked in the column 'Change' with
#   'Deleted' if a line exists in the old file but not in the new one
#   'New' if a line exists in the new file but not in the old one
#   'Unchanged' if a line exists as well in the old file as in the new file


# Path to the CSV file from the older run of Export-RecipientPermissions
$oldCsv = '.\Export-RecipientPermissions_Output_old.csv'

# Path to the CSV file from the newer run of Export-RecipientPermissions
$newCsv = '.\Export-RecipientPermissions_Output_new.csv'


#
# No changes from here on
#

Clear-Host

Set-Location $PSScriptRoot

Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"


Write-Host
Write-Host "Import old CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$oldLines = Import-Csv $oldCsv -Delimiter ';'

Write-Host
Write-Host "Import new CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$newLines = Import-Csv $newCsv -Delimiter ';'

Write-Host
Write-Host "Compare CSV files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
Write-Host "  Create dataset @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$Dataset = Compare-Object -ReferenceObject $oldLines -DifferenceObject $newLines -Property $newLines[0].psobject.properties.name -IncludeEqual -PassThru
$Dataset | Add-Member -MemberType NoteProperty -Name 'Change' -Value $null

Write-Host "  Modify dataset @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
foreach ($DatasetObject in $Dataset) {
    if ($DatasetObject.sideindicator -ieq '<=') {
        $DatasetObject.'Change' = 'Deleted'
    } elseif ($DatasetObject.sideindicator -ieq '=>') {
        $DatasetObject.'Change' = 'New'
    } else {
        $DatasetObject.'Change' = 'Unchanged'
    }
}


Write-Host
Write-Host "Display results @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
foreach ($GrantorPrimarySmtp in ($Dataset.'Grantor Primary SMTP' | Select-Object -Unique)) {
    Write-Host "$($GrantorPrimarySmtp)"
    foreach ($DatasetObject in $Dataset) {
        if ($DatasetObject.'Grantor Primary SMTP' -ieq $GrantorPrimarySMTP) {
            if ($DatasetObject.Change -ieq 'deleted') {
                Write-Host ("  Deleted: $($DatasetObject.'Trustee Primary SMTP') no longer has the '$($DatasetObject.'Permission')' right" + $(if ($DatasetObject.'Folder') { " on folder '$($DatasetObject.'Folder')'" }))
                $DatasetObject.'Change' = 'Deleted'
            } elseif ($DatasetObject.change -ieq 'new') {
                Write-Host ("  New: $($DatasetObject.'Trustee Primary SMTP') now has the '$($DatasetObject.'Permission')' right" + $(if ($DatasetObject.'Folder') { " on folder '$($DatasetObject.'Folder')'" }))
            } else {
                Write-Host ("  Unchanged: $($DatasetObject.'Trustee Primary SMTP') still has the '$($DatasetObject.'Permission')' right" + $(if ($DatasetObject.'Folder') { " on folder '$($DatasetObject.'Folder')'" }))
            }
        }
    }
}


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
