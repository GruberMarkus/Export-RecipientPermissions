# This sample code shows how to use TrusteeFilter to find permissions
# which may be affected by SIDHistory removal
#
# Provide the DistinguishedNames of the AD objects from which
# the SIDHistory should be removed in the $DNs variable
#
# TrusteeFilter is defined as a PowerShell here-string
#   Pay attention how the reference variable $Trustee is prefixed with a backtick,
#   so it is passed to Export-RecipientPermissions.ps1 as literal string and
#   therefore resolved within the export script and not already before (which
#   would lead to an empty string in this sample code)
#
#   Also pay attention how the arrays $PrimarySMTPs and $OriginalIdentities
#   are converted into array defining strings and are therefore passed as
#   values and not as references to Export-RecipientPermissions.ps1
#
# GrantorFilter behaves exactly like TrusteeFilter, only the reference variable
# is $Grantor instead of $Trustee
#
# You may want to adjust the file paths to Export-RecipientPermissions.ps1,
# ExportFile, ErrorFile and DebugFile in the last lines of this sample script

$DNs = (
    'CN=ObjectA,OU=OU3,OU=OU2,OU=OU1,DC=excample,DC=com',
    'CN=ObjectB,OU=OU3,OU=OU2,OU=OU1,DC=excample,DC=com'
)


Clear-Host


Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


Write-Host
Write-Host "Query AD for data from objects with SIDHistory to remove @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


$PrimarySMTPs = @()

$OriginalIdentities = @()

$count = 0

$Search = New-Object DirectoryServices.DirectorySearcher
$Search.PageSize = 1000

foreach ($DN in $DNs) {
    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$(($DN -split ',DC=')[1..999] -join '.')")
    $Search.filter = "distinguishedname=$($DN)"

    @('objectsid', 'sidhistory', 'distinguishedname', 'displayname', 'proxyaddresses', 'msds-principalname', 'legacyexchangedn') | ForEach-Object {
        $null = $search.propertiesToLoad.Add($_)
    }

    $result = $search.findone().properties

    if ($result.objectsid) {
        $OriginalIdentities += (New-Object System.Security.Principal.SecurityIdentifier $($result.objectsid), 0).value.tostring()
    }

    if ($result.sidhistory) {
        $result.sidhistory | ForEach-Object {
            $OriginalIdentities += (New-Object System.Security.Principal.SecurityIdentifier $_, 0).value.tostring()
        }
    }

    $OriginalIdentities += $result.distinguishedname

    $OriginalIdentities += $result.displayname

    if ($result.proxyaddresses) {
        $result.proxyaddresses | Where-Object { $_ -ilike 'smtp:*' } | ForEach-Object {
            $PrimarySMTPs += $_ -replace '^smtp:', ''
        }

        $result.proxyaddresses | Where-Object { $_ -ilike 'x500:*' } | ForEach-Object {
            $OriginalIdentities += $_ -replace '^x500:', ''
        }
    }

    $OriginalIdentities += $result.'msds-principalname'

    $OriginalIdentities += $result.legacyexchangedn

    if (($count % 100) -eq 0) {
        Write-Host (("`b" * 100) + ('  {0:0000000}/{1:0000000} @{2}@' -f $count, $DNs.count, $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'))) -NoNewline

        if (($count -eq 0) -or ($count -eq $DNs.count)) {
            Write-Host
        }
    }

    $count++
}

$PrimarySMTPs = $PrimarySMTPs | Where-Object { $_ }
$OriginalIdentities = $OriginalIdentities | Where-Object { $_ }


Write-Host
Write-Host "Export permissions from Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  Save to '.\export\Export-RecipientPermissions_Result_SIDHistory-Removal.csv'"

$params = @{
    ExportFromOnPrem          = $true
    UseDefaultCredential      = $true
    ExchangeConnectionUriList = 'http://server1.example.com/powershell/', 'http://server2.example.com/powershell/'
    GrantorFilter             = ''
    TrusteeFilter             = @"
if (`$Trustee.PrimarySmtpAddress) {
    $`Trustee.PrimarySmtpAddress -iin $('("' + (@($PrimarySMTPs) -join '", "') + '")')
} else {
    $`Trustee -iin $('("' + (@($OriginalIdentities) -join '", "') + '")')
}
"@
    ExportFile                = '.\export\Export-RecipientPermissions_Result_SIDHistory-Removal.csv'
    ErrorFile                 = '.\export\Export-RecipientPermissions_Error_SIDHistory-Removal.csv'
    DebugFile                 = ''
}


& ..\..\Export-RecipientPermissions.ps1 @params
