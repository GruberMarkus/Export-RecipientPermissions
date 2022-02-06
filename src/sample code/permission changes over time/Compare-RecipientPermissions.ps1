# Compare two result files from Export-RecipientPermissions.ps1
# to see which permissions have changed over time
#
# Only shows changes, equal lines are ignored
# Changes are marked in the column 'Change' with
#   'Deleted' if a line exists in the old file but not in the new one
#   'New' if a line exists in the new file but not in the old one


# Path to the CSV file from the older run of Export-RecipientPermissions
$oldCsv = ".\Export-RecipientPermissions_Output_old.csv"

# Path to the CSV file from the newer run of Export-RecipientPermissions
$newCsv = ".\Export-RecipientPermissions_Output_new.csv"


#
# No changes from here on
#

$oldLines = Import-Csv $oldCsv -Delimiter ";"
$newLines = Import-Csv $newCsv -Delimiter ";"

$Dataset = @()
$Dataset += compare-object -ReferenceObject $oldLines -DifferenceObject $newLines -Property $newLines[0].psobject.properties.name -IncludeEqual -PassThru
$Dataset | Add-Member -MemberType NoteProperty -Name 'Change' -Value $null
$Dataset | foreach {
    if ($_.sideindicator -ieq '<=') {
        $_.change = "Deleted"
    } elseif ($_.sideindicator -ieq '=>') {
        $_.change = "New"
    } else {
        $_.change = "Unchanged"
    }
}
$Dataset = $Dataset | select -property * -ExcludeProperty 'SideIndicator' -Unique | Sort-Object -Property 'Grantor Primary SMTP', 'Trustee Primary SMTP', 'Folder', 'Permission'


foreach ($Grantor in ($Dataset.'Grantor Primary SMTP' | select -Unique)) {
    write-host "$Grantor"
    $Dataset | where {($_.'Grantor Primary SMTP' -ieq $Grantor) -and ($_.Change -ieq 'New')} | foreach {
        write-host ("  New: $($_.'Trustee Primary SMTP') now has the '$($_.'Permission')' right" + $(if ($_.'Folder') { " on folder '$($_.'Folder')'"}))
    }

    $Dataset | where {($_.'Grantor Primary SMTP' -ieq $Grantor) -and ($_.Change -ieq 'Deleted')} | foreach {
        write-host ("  Deleted: $($_.'Trustee Primary SMTP') no longer has the '$($_.'Permission')' right" + $(if ($_.'Folder') { " on folder '$($_.'Folder')'"}))
    }

    $Dataset | where {($_.'Grantor Primary SMTP' -ieq $Grantor) -and ($_.Change -ieq 'Unchanged')} | foreach {
        write-host ("  Unchanged: $($_.'Trustee Primary SMTP') still has the '$($_.'Permission')' right" + $(if ($_.'Folder') { " on folder '$($_.'Folder')'"}))
    }
}
