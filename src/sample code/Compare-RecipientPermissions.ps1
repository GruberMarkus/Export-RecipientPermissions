cls
set-location "C:\Users\grube\Downloads"

# Compare two result files from Export-RecipientPermissions.ps1
# to see which permissions have changed over time
#
# Only shows changes, equal lines are ignored
# Changes are marked in the column 'Change' with
#   'Deleted' if a line exists in the old file but not in the new one
#   'New' if a line exists in the new file but not in the old one


# Path to the CSV file from the older run of Export-RecipientPermissions
$oldCsv = ".\Export-RecipientPermissions_Output_Sample from previous run.csv"

# Path to the CSV file from the newer run of Export-RecipientPermissions
$newCsv = ".\Export-RecipientPermissions_Output_Sample from current run.csv"


#
# No changes from here on
#

$oldLines = Import-Csv $oldCsv -Delimiter ";"
$newLines = Import-Csv $newCsv -Delimiter ";"

$x = @()
$x += compare-object -ReferenceObject $oldLines -DifferenceObject $newLines -Property $newLines[0].psobject.properties.name -PassThru
$x | Add-Member -MemberType NoteProperty -Name 'Change' -Value $null
$x | foreach {
    if ($_.sideindicator -ieq '<=') {
        $_.change = "New"
    } else {
        $_.change = "Deleted"
    }
}
$x = $x | select -property * -ExcludeProperty 'SideIndicator'

$x | convertto-csv -notypeinformation