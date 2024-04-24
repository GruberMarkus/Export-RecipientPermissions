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

    # Path for export file showing changes
    # Set to '' or $null to not create this file
    $ChangeFile = '.\comparison.csv'
)


Clear-Host


Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

if ($PSVersionTable.PSEdition -ieq 'desktop') {
    $UTF8Encoding = 'UTF8'
} else {
    $UTF8Encoding = 'UTF8BOM'
}


Write-Host
Write-Host "Check CSV files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$CompareLinebyLine = $false

if (
    (Get-Content $oldcsv -Head 1 -Encoding utf8) -ieq (Get-Content $oldcsv -Head 1 -Encoding utf8)
) {
    $CompareLinebyLine = $true

    Write-Host '  Headers of old and new CSV file are identical.'
    Write-Host '  Using faster and less memory intensive line-by-line comparison.'
} else {
    Write-Host '  Headers of old and new CSV file are not identical.'
    Write-Host '  Using slower and more memory-intensive object-based comparison.'
}


Write-Host
Write-Host "Import old CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($CompareLinebyLine) {
    $oldLines = Get-Content $oldCsv -Encoding $UTF8Encoding
} else {
    $oldLines = Import-Csv $oldCsv -Delimiter ';' -Encoding $UTF8Encoding
}


Write-Host
Write-Host "Import new CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($CompareLinebyLine) {
    $newLines = Get-Content $newCsv -Encoding $UTF8Encoding
} else {
    $newLines = Import-Csv $newCsv -Delimiter ';' -Encoding $UTF8Encoding
}


Write-Host
Write-Host "Compare CSV files @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  Create compared dataset @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($CompareLinebyLine) {
    $Dataset = Compare-Object -ReferenceObject $oldLines -DifferenceObject $newLines -IncludeEqual
    $Dataset[0].InputObject += ';"Change"'
} else {
    $Dataset = Compare-Object -ReferenceObject $oldLines -DifferenceObject $newLines -Property $newLines[0].psobject.properties.name -IncludeEqual -PassThru
    $Dataset | Add-Member -MemberType NoteProperty -Name 'Change' -Value $null
}

Write-Host "  Modify and sort dataset @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
for ($x = $(if ($CompareLinebyLine) { 1 } else { 0 }); $x -lt $dataset.count; $x++) {
    $DatasetObject = $Dataset[$x]

    if ($DatasetObject.sideindicator -ieq '<=') {
        if ($CompareLinebyLine) {
            $DatasetObject.InputObject += ';"Deleted"'
        } else {
            $DatasetObject.'Change' = 'Deleted'
        }
    } elseif ($DatasetObject.sideindicator -ieq '=>') {
        if ($CompareLinebyLine) {
            $DatasetObject.InputObject += ';"New"'
        } else {
            $DatasetObject.'Change' = 'New'
        }
    } else {
        if ($CompareLinebyLine) {
            $DatasetObject.InputObject += ';"Unchanged"'
        } else {
            $DatasetObject.'Change' = 'Unchanged'
        }
    }

    $Dataset[$x] = $DatasetObject
}

if (-not $CompareLinebyLine) {
    $Dataset = $Dataset | Select-Object * -ExcludeProperty 'SideIndicator'
}


Write-Host
Write-Host "Create export file showing changes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($ChangeFile) {
    Write-Host "  '$($ChangeFile)'"

    if ($CompareLinebyLine) {
        (
            ($Dataset[0].inputobject) +
            [System.Environment]::NewLine +
            ((@(($Dataset[1..$($Dataset.count - 1)]).InputObject) | Sort-Object) -join [System.Environment]::NewLine)
        ) | Set-Content -Path $ChangeFile -Encoding $UTF8Encoding -Force
    } else {
        $Dataset = $Dataset | Sort-Object -Property 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Folder', 'Permission', 'Trustee Original Identity', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Change'
        $Dataset | Export-Csv -Path $ChangeFile -Delimiter ';' -Encoding $UTF8Encoding -NoTypeInformation -Force
    }
} else {
    Write-Host '  Not required with current configuration settings'
}


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
