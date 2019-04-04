Param(
    # Environments to consider: Office 365 (Exchange Online) and/or Exchange on premises
    [boolean]$ExportFromOnPrem = $true, # Highly recommended to enable this for fast initial recipient enumeration
    [boolean]$ExportFromCloud = $true,

    # Permission types to export
    [boolean]$ExportAccessRights = $true, # Rights like "FullAccess" and "ReadAccess" to the entire mailbox
    [boolean]$ExportFullAccessPerTrustee = $true, # Additionally export a list showing who has full access to which mailbox
    [boolean]$ExportSendAs = $true, # Send as
    [boolean]$ExportSendOnBehalf = $true, # Send on behalf
    [boolean]$ExportManagedBy = $true, # Only valid for groups
    [boolean]$ExportFolderPermissions = $false, # Export permissions set on specific mailbox folders. This will take very long.

    # Name of the permission export file
    [string]$ExportFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Output.csv",

    # Name of the error file
    [string]$ErrorFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Errors.txt",

    # Name of the transcript file
    [string]$TranscriptFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Transcript.txt",

    # Name of temporary recipient file
    [string]$TempRecipientFile = ".\Export-RecipientPermissions_Output\Export-RecipientPermissions_Recipients.csv",

    # Folder to additionally store files created when $ExportFullAccessPerTrustee = $true. This folder must already exist at runtime. Set to "" when not needed.
    [string]$TargetFolder = "\\server.domain\share\folder",

    # Parallelization
    # Watch RAM and CPU usage
    [int]$NumberOfJobsParallel = 30, # Each job is a separate session towards Exchange on-prem and Office 365, so watch your maximum concurreny settings
    [int]$RecipientsPerJob = 100, # More recipients save time as jobs run longer, but the risk of a problem with the O365 connection is higher

    # User name and password are stored in secure string format
    [string]$CredentialPasswordFile = ".\Export-RecipientPermissions_CredentialPassword.txt",
    [string]$CredentialUsernameFile = ".\Export-RecipientPermissions_CredentialUsername.txt"
)


#
# Do not change anything from here on.
#

Function Pause($M="Press any key to continue . . . "){If($psISE){$S=New-Object -ComObject "WScript.Shell";$B=$S.Popup("Click OK to continue.",0,"Script Paused",0);Return};write-host -NoNewline $M;$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode){$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")};write-host}

$script:SessionCloud = $null

function Connect-ExchangeOnPrem {
    $Stoploop = $false
    [int]$Retrycount = 0
    do {
	    try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
            . $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
            Connect-ExchangeServer -auto -ErrorAction Stop
		    Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop | out-null
            $Stoploop = $true
        } catch {
            if ($Retrycount -le 3) {
                write-host "Could not connect to Exchange on-prem. Trying again in 70 seconds."
                Start-Sleep -Seconds 70
                $Retrycount = $Retrycount + 1
            } else {
                write-host "Could not connect to Exchange on-prem after three retires. Exiting."
                exit
                $Stoploop = $true
            }
        }
    } While ($Stoploop -eq $false)
}

function Connect-ExchangeOnline {
    $Stoploop = $false
    [int]$Retrycount = 0
    do {
	    try {
            $test = $null
            $test = (get-pssession | where-object {($_.name -like "O365Session") -and ($_.state -like "opened")})
            if ($test -eq $null) {
                $CloudUser = get-content $CredentialUsernameFile
                $CloudPassword = get-content $CredentialPasswordFile | convertto-securestring
                $UserCredential = new-object -typename System.Management.Automation.PSCredential -argumentlist $CloudUser, $CloudPassword
                $script:SessionCloud = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -Name "O365Session" -ErrorAction Stop
                (Invoke-Command -Session $script:SessionCloud -ScriptBlock {Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop}) | out-null
            }
            $Stoploop = $true
        } catch {
            if ($Retrycount -le 3) {
                write-host "Could not connect to Exchange Online. Trying again in 70 seconds."
                Start-Sleep -Seconds 70
                $Retrycount = $Retrycount + 1
            } else {
                write-host "Could not connect to Exchange Online after three retires. Exiting."
                exit
                $Stoploop = $true
            }
        }
    } While ($Stoploop -eq $false)
}

function write-HostColored() {
    <#
    .SYNOPSIS
    A wrapper around write-Host that supports selective coloring of
    substrings.

    .DESCRIPTION
    In addition to accepting a default foreground and background color,
    you can embed one or more color specifications in the string to write, 
    using the following syntax:
    #<fgcolor>[:<bgcolor>]#<text>#

    <fgcolor> and <bgcolor> must be valid [ConsoleColor] values, such as 'green' or 'white' (case does not matter).
    Everything following the color specification up to the next '#' or, impliclitly, the end of the string
    is written in that color.

    Note that nesting of color specifications is not supported.
    As a corollary, any token that immediately follows a color specification is treated
    as text to write, even if it happens to be a technically valid color spec too.
    This allows you to use, e.g., 'The next word is #green#green#.', without fear
    of having the second '#green' be interpreted as a color specification as well.

    .PARAMETER ForegroundColor
    Specifies the default text color for all text portions
    for which no embedded foreground color is specified.

    .PARAMETER BackgroundColor
    Specifies the default background color for all text portions
    for which no embedded background color is specified.

    .PARAMETER NoNewline
    Output the specified string withpout a trailing newline.

    .NOTES
    While this function is convenient, it will be slow with many embedded colors, because,
    behind the scenes, write-Host must be called for every colored span.

    .EXAMPLE
    write-HostColored "#green#Green foreground.# Default colors. #blue:white#Blue on white."

    .EXAMPLE
    '#black#Black on white (by default).#Blue# Blue on white.' | Write-HostColored -BackgroundColor White

    #>
    [CmdletBinding(ConfirmImpact='None', SupportsShouldProcess=$false, SupportsTransactions=$false)]
    param(
        [parameter(Position=0, ValueFromPipeline=$true)]
        [string[]] $Text
        ,
        [switch] $NoNewline
        ,
        [ConsoleColor] $BackgroundColor =  $host.UI.RawUI.BackgroundColor
        ,
        [ConsoleColor] $ForegroundColor = $host.UI.RawUI.ForegroundColor
    )

    begin {
        # If text was given as an operand, it'll be an array.
        # Like write-Host, we flatten the array into a single string
        # using simple string interpolation (which defaults to separating elements with a space,
        # which can be changed by setting $OFS).
        if ($Text -ne $null) {
            $Text = "$Text"
        }
    }

    process {
        if ($Text) {

            # Start with the foreground and background color specified via
            # -ForegroundColor / -BackgroundColor, or the current defaults.
            $curFgColor = $ForegroundColor
            $curBgColor = $BackgroundColor

            # Split message into tokens by '#'.
            # A token between to '#' instances is either the name of a color or text to write (in the color set by the previous token).
            $tokens = $Text.split("#")

            # Iterate over tokens.            
            $prevWasColorSpec = $false
            foreach($token in $tokens) {

                if (-not $prevWasColorSpec -and $token -match '^([a-z]+)(:([a-z]+))?$') { # a potential color spec.
                    # If a token is a color spec, set the color for the next token to write.
                    # Color spec can be a foreground color only (e.g., 'green'), or a foreground-background color pair (e.g., 'green:white')
                    try {
                        $curFgColor = [ConsoleColor]  $matches[1]
                        $prevWasColorSpec = $true
                    } catch {}
                    if ($matches[3]) {
                        try {
                            $curBgColor = [ConsoleColor]  $matches[3]
                            $prevWasColorSpec = $true
                        } catch {}
                    }
                    if ($prevWasColorSpec) {
                        continue                    
                    }
                }

                $prevWasColorSpec = $false

                if ($token) {
                    # A text token: write with (with no trailing line break).
                    # !! In the ISE - as opposed to a regular PowerShell console window,
                    # !! $host.UI.RawUI.ForegroundColor and $host.UI.RawUI.ForegroundColor inexcplicably 
                    # !! report value -1, which causes an error when passed to write-Host.
                    # !! Thus, we only specify the -ForegroundColor and -BackgroundColor parameters
                    # !! for values other than -1.
                    $argsHash = @{}
                    if ([int] $curFgColor -ne -1) { $argsHash += @{ 'ForegroundColor' = $curFgColor } }
                    if ([int] $curBgColor -ne -1) { $argsHash += @{ 'BackgroundColor' = $curBgColor } }
                    write-host -NoNewline @argsHash $token
                }

                # Revert to default colors.
                $curFgColor = $ForegroundColor
                $curBgColor = $BackgroundColor

            }
        }
        # Terminate with a newline, unless suppressed
        if (-not $NoNewLine) { write-Host }
    }
}

$error.clear()

Set-Location $PSScriptRoot

Clear-Host

$TargetFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TargetFolder)
$Exportfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Exportfile)
$ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
$TranscriptFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TranscriptFile)
$TempRecipientFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TempRecipientFile)
$CredentialPasswordFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CredentialPasswordFile)
$CredentialUsernameFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CredentialUsernameFile)
New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile) | out-null
New-Item -ItemType Directory -Force -Path (Split-Path -Path $TempRecipientFile) | out-null
if (Test-Path $Exportfile) {(Remove-Item $ExportFile -force)}
if (Test-Path $Errorfile) {(Remove-Item $ErrorFile -force)}
if (Test-Path $TranscriptFile) {(Remove-Item $TranscriptFile -force)}
if (Test-Path $TempRecipientFile) {(Remove-Item $TempRecipientFile -force)}
if (($ExportFullAccessPerTrustee -eq $true) -and ($ExportAccessRights -eq $true)) {
    if (test-path ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee")) {
        Remove-Item ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee") -Force -Recurse
    }
    New-Item -ItemType Directory -Force -Path ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee") | out-null
}

start-transcript -path ($TranscriptFile + "_temp") -force

if ($ExportFromCloud -eq $true) {
    if ((Test-Path $CredentialUsernameFile) -and (Test-Path $CredentialPasswordFile)) { } else {
        write-host 'Please enter cloud user name for later use.'
        read-host | out-file $CredentialUsernameFile
        write-host 'Please enter cloud admin password for later use.'
        read-host -assecurestring | convertfrom-securestring | out-file $CredentialPasswordFile
    }
}

# Test on-prem connection
if (($ExportFromOnPrem -eq $true)) {
	Try {
        Connect-ExchangeOnPrem
        write-host 'On-prem connection working.'
	}
	Catch
	{
		write-host 'On-prem connection does not work. Error executing ''Get-Recipient -ResultSize 1''. Exiting.'
		write-host 'Please start the script on an Exchange server with appropriate permissions.'
        $ExportFromOnPrem = $false
        exit
	}
}

if (($ExportFromCloud -eq $true)) {
	Try {
        write-host "Connecting to Exchange Online."
        Connect-ExchangeOnline
        write-host 'Cloud connection working.'
	}
	Catch
	{
		write-host 'Cloud connection does not work. Error executing ''Get-Recipient -ResultSize 1''. Exiting.'
        $ExportFromCloud = $false
        exit
	}
}

# Export list of objects
if ($ExportFromOnPrem -eq $true) {
    write-host 'Enumerating on-prem recipients. This may take a long time.'
} else {
    if ($ExportFromCloud -eq $true) {
        write-host 'Enumerating cloud recipients. This may take a long time.'
    } else {
        write-host 'Neither on-prem nor cloud connection configured or possible. Exiting.'
        exit
    }
}


if ($ExportFromOnPrem -eq $true) {
    get-recipient -recipienttype MailUniversalSecurityGroup, DynamicDistributionGroup, UserMailbox, MailUniversalDistributionGroup, MailUniversalSecurityGroup, MailNonUniversalGroup, MailUser  -resultsize unlimited  -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    <#
    get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype PublicFolder -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype DynamicDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype UserMailbox -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailUniversalDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailNonUniversalGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailUser -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    get-recipient -recipienttype MailContact -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    #>
} else {
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalSecurityGroup, DynamicDistributionGroup, UserMailbox, MailUniversalDistributionGroup, MailUniversalSecurityGroup, MailNonUniversalGroup, MailUser -resultsize unlimited -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    <#
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype PublicFolder -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype DynamicDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype UserMailbox -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalDistributionGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUniversalSecurityGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailNonUniversalGroup -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailUser -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    (Invoke-Command -Session $script:SessionCloud -ScriptBlock {get-recipient -recipienttype MailContact -resultsize 1000 -WarningAction silentlyContinue | select-object DistinguishedName}) | select-object DistinguishedName | export-csv $TempRecipientFile -Append -NoTypeInformation -Force -Delimiter ';'
    #>
}


if (($ExportFromCloud -eq $true)) {
    write-host 'Disconnecting from cloud services.'
    Remove-PSSession $script:SessionCloud
    #if ((test-path (Split-Path -Path $script:SessionCloudPath.path)) -eq $true) {
    #    Remove-Item (Split-Path -Path $script:SessionCloudPath.path) -Recurse -Force
    #}
}

# Import list of objects
$Recipients = (import-csv $TempRecipientFile)
$RecipientCount = $Recipients.length
$count = 1



$Batch=0
for ($RecipientStartID = 0; $RecipientStartID -lt $RecipientCount; $RecipientStartID += $RecipientsPerJob) {
    $RecipientEndID=($RecipientStartID + $RecipientsPerJob -1)
    $Batch++
}

write-host "$RecipientCount recipients found. Reading permissions in $Batch batches of $RecipientsPerJob recipients each."
write-host "$NumberOfJobsParallel batches will run in parallel. Output is updated at completion of a single batch."

get-job | Remove-Job -force
$Batch=1
for ($RecipientStartID = 0; $RecipientStartID -lt $RecipientCount; $RecipientStartID += $RecipientsPerJob) {
    $RecipientEndID=($RecipientStartID + $RecipientsPerJob - 1)
    $running = @(Get-Job -state running)
    foreach ($x in (Get-Job -state Completed)) {
        if (test-path ($Exportfile + '_temp' + $x.name)) { (get-content ($Exportfile + '_temp' + $x.name)) | Write-HostColored }
    }
    if ($running.Count -ge $NumberOfJobsParallel) {
        # wait and receive
        while($true) {
            if (@(Get-Job -state running).count -lt $NumberOfJobsParallel) {
                foreach ($x in (Get-Job -state Completed)) {
                    $temp = $null
                    $TempPath = $null
                    # show temp job output file, delete output file
                    $TempPath = ($Exportfile + '_temp' + $x.Name)
                    if (test-path $TempPath) {
                        $temp = get-content $TempPath
                        $temp | Write-HostColored
                        Remove-Item $TempPath -force
                    }
                    $temp = $null
                    $TempPath = $null

                    # append temp error file and delete temp file
                    $TempPath = ($Errorfile + '_temp' + $x.Name)
                    $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
                    if (Test-Path $TempPath) {
                        $temp = get-content $TempPath
                        $temp | Out-File $Errorfile -Append -Force
                        Remove-Item $TempPath -force
                    }
                    $temp = $null
                    $TempPath = $null

                    # append temp transcript file and delete temp file
                    $TempPath = ($Transcriptfile + '_temp' + $x.Name)
                    $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
                    if (Test-Path $TempPath) {
                        $temp = get-content $TempPath
                        $temp | Out-File $Transcriptfile -Append -Force
                        Remove-Item $TempPath -force
                    }

                    # append temp export file and delete temp file
                    $TempPath = ($Exportfile + '_temp' + $x.Name)
                    $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
                    if (Test-Path $TempPath) {
                        $temp = import-csv $TempPath -delimiter ";"
                        $temp | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                        Remove-Item $TempPath -force
                    }
                    $temp = $null
                    $TempPath = $null
                    Remove-Job -job $x -force
                }
                break
            } else {
                [System.GC]::Collect() # garbage collection
                start-sleep -s 5
            }
        }
    }
    Start-Job {
        param(
            $RecipientStartID,
            $RecipientEndID,
            $Exportfile,
            $ErrorFile,
            $TempRecipientFile,
            $ExportFromOnPrem,
            $ExportFromCloud,
            $CredentialPasswordFile,
            $CredentialUsernameFile,
            $ExportAccessRights,
            $ExportSendAs,
            $ExportSendOnBehalf,
            $ExportManagedBy,
            $ExportFolderPermissions,
            $ExportFullAccessPerTrustee,
            $TranscriptFile
        )
        start-sleep -s (get-random -minimum 0 -maximum 20)
        Set-Location $PSScriptRoot
        $Exportfile = $Exportfile + '_temp' + $RecipientStartID
        $ErrorFile = $ErrorFile + '_temp' + $RecipientStartID
        $TranscriptFile = $TranscriptFile + '_temp' + $RecipientStartID
        $Exportfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Exportfile)
        $ErrorFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ErrorFile)
        $TranscriptFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TranscriptFile)
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $Exportfile) | out-null
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $ErrorFile) | out-null
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $TranscriptFile) | out-null
        if (Test-Path $Exportfile) {(Remove-Item $ExportFile -force)}
        if (Test-Path $Errorfile) {(Remove-Item $ErrorFile -Force)}
        start-transcript -path $TranscriptFile -force
        write-host ("RecipientStartID: " + $RecipientStartID)
        write-host ("RecipientEndID: " + $RecipientEndID)
        write-host ("Time: " + (get-date))

        $script:BatchSessionCloud = $null
        function Connect-ExchangeOnPrem {
            $Stoploop = $false
            [int]$Retrycount = 0
            do {
	            try {
                    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
                    . $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                    Connect-ExchangeServer -auto -ErrorAction Stop
		            Get-Recipient -ResultSize 1 -wa silentlycontinue -ea stop | out-null
                    $Stoploop = $true
                } catch {
                    if ($Retrycount -le 3) {
                        write-host ("Time: " + (get-date))
                        write-host "Could not connect to Exchange on-prem. Trying again in 70 seconds."
                        Start-Sleep -Seconds 70
                        $Retrycount = $Retrycount + 1
                    } else {
                        write-host ("Time: " + (get-date))
                        write-host "Could not connect to Exchange on-prem after three retires. Exiting."
                        exit
                        $Stoploop = $true
                    }
                }
            } While ($Stoploop -eq $false)
        }

        function Connect-ExchangeOnline {
            $Stoploop = $false
            [int]$Retrycount = 0
            do {
	            try {
                    $test = $null
                    $test = (get-pssession | where-object {($_.name -like "O365BatchSession") -and ($_.state -like "opened")})
                    if ($test -eq $null) {
                        $CloudUser = get-content $CredentialUsernameFile
                        $CloudPassword = get-content $CredentialPasswordFile | convertto-securestring
                        $UserCredential = new-object -typename System.Management.Automation.PSCredential -argumentlist $CloudUser, $CloudPassword
                        $script:BatchSessionCloud = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -Name "O365BatchSession" -ErrorAction Stop
                    }
                    $Stoploop = $true
                } catch {
                    if ($Retrycount -le 3) {
                        write-host ("Time: " + (get-date))
                        write-host "Could not connect to Office 365. Trying again in 70 seconds."
                        Start-Sleep -Seconds 70
                        $Retrycount = $Retrycount + 1
                    } else {
                        write-host ("Time: " + (get-date))
                        write-host "Could not connect to Office 365 after three retires. Exiting."
                        exit
                        $Stoploop = $true
                    }
                }
            } While ($Stoploop -eq $false)
        }
       

        $Recipients = (import-csv $TempRecipientFile)
        $RecipientCount=$Recipients.count
        $Count = $RecipientStartID + 1
        #$BatchID = ($RecipientStartID / ($RecipientEndID - $RecipientStartID + 1)) + 1
        if (($ExportFromCloud -eq $true)) {Connect-ExchangeOnline}
        if (($ExportFromOnPrem -eq $true)) {Connect-ExchangeOnPrem}
        $ErrorCount=0
        for ($RecipientStartID; $RecipientStartID -le $RecipientEndID; $RecipientStartID++) {
            write-host ((write-host ("Time: " + (get-date))) + "; RecipientID: " + $RecipientStartID)
            if ($RecipientStartID -ge $Recipients.length) {break}
            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
            if ($ExportFromOnPrem -eq $true) {
                $Recipient=get-recipient $Recipients[$RecipientStartID].DistinguishedName -resultsize 1
            } else {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $Recipient=(Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient $($args[0]) -resultsize 1} -argumentlist $Recipients[$RecipientStartID].DistinguishedName)
            }
            $GrantorDisplayName = ""
            $GrantorPrimarySMTP = ""
            $GrantorRecipientType = ""
            $GrantorRecipientTypeDetails = ""
            $GrantorLegacyExchangeDN = ""
            $GrantorOU = ""
            $ALias = ""
            $x = $null
            if ($ExportFromOnPrem -eq $true) {
                if ($Recipient.RecipientTypeDetails -like "Remote*") {$GrantorCloudOrOnPrem = 'Cloud'} else {$GrantorCloudOrOnPrem = 'On-Prem'}
                if ($Recipient.RecipientTypeDetails -like "*Group") {$GrantorCloudOrOnPrem = 'On-Prem'}
            } else {
                if ($Recipient.RecipientTypeDetails -like "Remote*") {$GrantorCloudOrOnPrem = 'On-Prem'} else {$GrantorCloudOrOnPrem = 'Cloud'}
                if ($Recipient.RecipientTypeDetails -like "*Group") {$GrantorCloudOrOnPrem = 'Cloud'}
            }
            $GrantorDisplayName = $Recipient.DisplayName.tostring()
            $GrantorPrimarySMTP = $Recipient.PrimarySMTPAddress.tostring()
            $GrantorRecipientType = $Recipient.RecipientType.tostring()
            $GrantorRecipientTypeDetails = $Recipient.RecipientTypeDetails.tostring()
            $GrantorOU = $Recipient.OrganizationalUnit.tostring()
            $Alias = $Recipient.name.tostring()
            if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $RecipientTemp=(Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient $($args[0]) -resultsize 10} -ArgumentList $GrantorPrimarySMTP)
                if ($recipientTemp.count) {
                    foreach ($x in $recipientTemp) {
                        if ($x.recipienttypedetails -like "*mailbox*") {
                            $GrantorDN = $x.DistinguishedName.tostring()
                        }
                    }
                } else {
                    $GrantorDN = $recipienttemp.DistinguishedName.tostring()
                }
            } else {
                $GrantorDN = $Recipient.DistinguishedName.tostring()
            }
            $x = $null
            $recipientTemp = $null

            $Text = ("{0:000000}/{1:000000}: " -f $count, $RecipientCount) + $GrantorPrimarySMTP + ', ' + $GrantorRecipienttype + '/' + $GrantorRecipientTypeDetails + ', ' + $GrantorCloudOrOnPrem

            if (($Recipient.Recipienttype -eq "PublicFolder") -or ($Recipient.Recipienttype -eq "MailContact")) {$Text += (", recipient type $GrantorRecipientType not supported."); continue}


            # Access Rights (full access etc.)
            if ($ExportAccessRights -eq $true) {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                if (($GrantorRecipientType -NotMatch 'group')) {
                    $Text += ', AccessRights'
                    if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                        try {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            $TrusteeRightsQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-MailboxPermission -identity $($args[0]) -resultsize unlimited -wa stop -ea stop} -ArgumentList $GrantorDN) | where-object {($_.IsInherited -eq $false) -and ($_.user -notlike '*NT AUTHORITY\SELF*')}
                            $GrantorLegacyExchangeDN = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-Mailbox -identity $($args[0]) -resultsize 1 -wa stop -ea stop | select-object LegacyExchangeDN} -ArgumentList $GrantorDN).LegacyExchangeDN
                        } catch {
                        }
                    } else {
                        try {
                            $TrusteeRightsQuery = (Get-MailboxPermission -identity $GrantorDN -resultsize unlimited -wa stop -ea stop | where-object {($_.IsInherited -eq $false) -and ($_.user -notlike '*NT AUTHORITY\SELF*')})
                            $GrantorLegacyExchangeDN = (Get-Mailbox -identity $GrantorDN -resultsize 1 -wa stop -ea stop).LegacyExchangeDN
                        } catch {
                        }
                    }
                    $TrusteeIdentityOriginal = @($TrusteeRightsQuery | select-object @{Name = 'Trustee'; Expression = {$_.User}})
                    if ($error.count -eq 0) {
                        foreach ($TrusteeIdentity in $TrusteeIdentityOriginal) {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            try {
                                $TrusteeIdentityQuery = (get-recipient ($TrusteeIdentity.trustee.tostring()) -resultsize 1 -wa stop -ea stop)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'Cloud'} else {$TrusteeCloudOrOnPrem = 'On-Prem'}
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'On-Prem'}
                            } catch {
                                try {
                                    if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                                    $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop} -ArgumentList $TrusteeIdentity.trustee.tostring())
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'On-Prem'} else {$TrusteeCloudOrOnPrem = 'Cloud'}
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'Cloud'}
                                } catch {
                                    continue
                                }
                            }
                            $error.clear()
                            $TrusteeRecipientType = $null
                            $TrusteeRecipientTypeDetails = $null
                            $TrusteeDisplayName = $null
                            $TrusteePrimarySMTP = $null
                            $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                            $TrusteeRightsQuery | where-object {($_.user -like $TrusteeIdentity.trustee.ToString())} | select-object @{name = 'Grantor Primary SMTP'; expression = {$GrantorPrimarySMTP}}, @{name = 'Grantor Display Name'; expression = {$GrantorDisplayName}}, @{name = 'Grantor Recipient Type'; expression = {$GrantorRecipientType + '/' + $GrantorRecipientTypeDetails}}, @{name = 'Grantor Environment'; expression = {$GrantorCloudOrOnPrem}}, @{Name = 'Trustee Primary SMTP'; Expression = {$TrusteePrimarySMTP}}, @{Name = 'Trustee Display Name'; Expression = {$TrusteeDisplayName}}, @{Name = 'Trustee Original Identity'; Expression = {$TrusteeIdentity.trustee.ToString()}}, @{name = 'Trustee Recipient Type'; expression = {$TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails}}, @{name = 'Trustee Environment'; expression = {$TrusteeCloudOrOnPrem}}, @{name = 'Permission(s)'; expression = {[string]::join(', ', @($_.AccessRights))}}, @{Name = 'Folder Name'; Expression = {''}}, @{Name = 'Grantor LegacyExchangeDN'; Expression ={$GrantorLegacyExchangeDN}}, @{Name = 'Grantor OU'; Expression ={$GrantorOU}}, @{Name = 'Trustee OU'; Expression ={$TrusteeOU}} | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                        }
                        if ($error) {
                            $ErrorCount++
                            $Text += ' #white:red#ERROR#'
                            "==============================" | out-file $ErrorFile -Append
                            ("{0:000000}/{1:000000}: {2}, AccessRights" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | out-file $ErrorFile -Append
                            for ($e = ($error.count - 1); $e -ge 0; $e--) {
                                $error[$e] | out-file $ErrorFile -Append
                                "" | out-file $ErrorFile -Append
                            }
                            "" | out-file $ErrorFile -Append; "" | out-file $ErrorFile -Append
                        }
                        $ErrorActionPreference = "Continue"
                        $WarningPreference = "Continue"
                        $error.clear()
                    }
                }
                $GrantorLegacyExchangeDN = ""
            }


            # Send As
            if ($ExportSendAs -eq $true) {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', SendAs'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    try {
                        if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                        $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-RecipientPermission -identity $($args[0]) -resultsize unlimited -wa stop -ea stop} -ArgumentList $GrantorDN) | where-object {($_.Trustee -notlike '*NT AUTHORITY\SELF*') -and ($_.AccessRights -like '*SendAs*')}
                    } catch {
                    }
                } else {
                    try {
                        $TrusteeIdentityQuery = (Get-ADPermission -identity $GrantorDN -wa stop -ea stop | where-object {($_.user -notlike '*NT AUTHORITY\SELF*') -and ($_.ExtendedRights -like '*Send-As*')} | select-object *, @{Name="trustee";Expression={$_."user"}})
                    } catch {
                    }
                }
                $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object trustee)
                if ($error.count -eq 0) {
                    foreach ($TrusteeIdentity in $TrusteeIdentityOriginal.trustee) {
                        if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                        $TrusteeIdentity = $TrusteeIdentity.tostring()
                        try {
                            $TrusteeIdentityQuery = (get-recipient $TrusteeIdentity -resultsize 1 -wa stop -ea stop)
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'Cloud'} else {$TrusteeCloudOrOnPrem = 'On-Prem'}
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'On-Prem'}
                        } catch {
                            try {
                                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient ($($args[0])) -resultsize 1 -wa stop -ea stop} -ArgumentList $TrusteeIdentity)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'On-Prem'} else {$TrusteeCloudOrOnPrem = 'Cloud'}
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'Cloud'}
                            } catch {
                                continue
                            }
                        }
                        $error.clear()
                        $TrusteeIdentityOriginal = $TrusteeIdentity
                        $TrusteeRecipientType = $null
                        $TrusteeRecipientTypeDetails = $null
                        $TrusteeDisplayName = $null
                        $TrusteePrimarySMTP = $null
                        $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                        $TrusteeIdentityQuery | select-object @{name = 'Grantor Primary SMTP'; expression = {$GrantorPrimarySMTP}}, @{name = 'Grantor Display Name'; expression = {$GrantorDisplayName}}, @{name = 'Grantor Recipient Type'; expression = {$GrantorRecipientType + '/' + $GrantorRecipientTypeDetails}}, @{name = 'Grantor Environment'; expression = {$GrantorCloudOrOnPrem}}, @{Name = 'Trustee Primary SMTP'; Expression = {$TrusteePrimarySMTP}}, @{Name = 'Trustee Display Name'; Expression = {$TrusteeDisplayName}}, @{Name = 'Trustee Original Identity'; Expression = {$TrusteeIdentityOriginal}}, @{name = 'Trustee Recipient Type'; expression = {$TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails}}, @{name = 'Trustee Environment'; expression = {$TrusteeCloudOrOnPrem}}, @{Name = 'Permission(s)'; Expression = {'SendAs'}}, @{Name = 'Folder Name'; Expression = {''}}, @{Name = 'Grantor LegacyExchangeDN'; Expression ={$GrantorLegacyExchangeDN}}, @{Name = 'Grantor OU'; Expression ={$GrantorOU}}, @{Name = 'Trustee OU'; Expression ={$TrusteeOU}} | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | out-file $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, SendAs" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | out-file $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | out-file $ErrorFile -Append
                        "" | out-file $ErrorFile -Append
                    }
                    "" | out-file $ErrorFile -Append; "" | out-file $ErrorFile -Append
                }
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
            }


            # Send On Behalf
            if (($ExportSendOnBehalf -eq $true)) {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', SendOnBehalf'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    if (($GrantorRecipientType -match 'group') -and ($GrantorRecipientType -notmatch 'DynamicDistributionGroup')) {
                        if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                        try {
                            $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-distributiongroup -identity $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $GrantorDN) | where-object {$_.GrantSendOnBehalfto -ne ''}
                        } catch {
                        }
                        $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object @{Name = 'Trustee'; Expression = {$_.GrantSendonBehalfto}})
                    } else {
                        if (($GrantorRecipientType -like 'DynamicDistributionGroup')) {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            try {
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-recipient -identity $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $GrantorDN) | where-object {$_.GrantSendOnBehalfto -ne ''}
                            } catch {
                            }
                            $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object @{Name = 'Trustee'; Expression = {$_.GrantSendonBehalfto}})
                        } else {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            try {
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-Mailbox -identity $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $GrantorDN) | where-object {$_.GrantSendOnBehalfto -ne ''}
                            } catch {
                            }
                        }
                    }
                } else {
                    if (($GrantorRecipientType -match 'group') -and ($GrantorRecipientType -notmatch 'DynamicDistributionGroup')) {
                        try {$TrusteeIdentityQuery = (Get-distributiongroup -identity $GrantorDN -resultsize 1 -wa stop -ea stop| where-object {$_.GrantSendOnBehalfto -ne ''})} catch {}
                        $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object @{Name = 'Trustee'; Expression = {$_.GrantSendonBehalfto}})
                    } else {
                        if (($GrantorRecipientType -like 'DynamicDistributionGroup')) {
                            try {$TrusteeIdentityQuery = (Get-recipient -identity $GrantorDN -resultsize 1 -wa stop -ea stop| where-object {$_.GrantSendOnBehalfto -ne ''})} catch {}
                            $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object @{Name = 'Trustee'; Expression = {$_.GrantSendonBehalfto}})
                        } else {
                            try {$TrusteeIdentityQuery = (Get-Mailbox -identity $GrantorDN -resultsize 1 -wa stop -ea stop| where-object {$_.GrantSendOnBehalfto -ne ''})} catch {}
                        }
                    }
                }
                $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object @{Name = 'Trustee'; Expression = {$_.GrantSendonBehalfto}})
                if ($error.count -eq 0) {
                    foreach ($TrusteeIdentity in $TrusteeIdentityOriginal.trustee) {
                        if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                        try {
                            $TrusteeIdentityQuery = (get-recipient ($TrusteeIdentity) -resultsize 1 -wa stop -ea stop)
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'Cloud'} else {$TrusteeCloudOrOnPrem = 'On-Prem'}
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'On-Prem'}
                        } catch {
                            try {
                                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $TrusteeIdentity)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'On-Prem'} else {$TrusteeCloudOrOnPrem = 'Cloud'}
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'Cloud'}
                            } catch {
                                continue
                            }
                        }
                        $error.clear()
                        $TrusteeIdentityOriginal = $TrusteeIdentity
                        $TrusteeRecipientType = $null
                        $TrusteeRecipientTypeDetails = $null
                        $TrusteeDisplayName = $null
                        $TrusteePrimarySMTP = $null
                        $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                        $TrusteeIdentityQuery | select-object @{name = 'Grantor Primary SMTP'; expression = {$GrantorPrimarySMTP}}, @{name = 'Grantor Display Name'; expression = {$GrantorDisplayName}}, @{name = 'Grantor Recipient Type'; expression = {$GrantorRecipientType + '/' + $GrantorRecipientTypeDetails}}, @{name = 'Grantor Environment'; expression = {$GrantorCloudOrOnPrem}}, @{Name = 'Trustee Primary SMTP'; Expression = {$TrusteePrimarySMTP}}, @{Name = 'Trustee Display Name'; Expression = {$TrusteeDisplayName}}, @{Name = 'Trustee Original Identity'; Expression = {$TrusteeIdentityOriginal}}, @{name = 'Trustee Recipient Type'; expression = {$TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails}}, @{name = 'Trustee Environment'; expression = {$TrusteeCloudOrOnPrem}}, @{Name = 'Permission(s)'; Expression = {'SendOnBehalf'}}, @{Name = 'Folder Name'; Expression = {''}}, @{Name = 'Grantor LegacyExchangeDN'; Expression ={$GrantorLegacyExchangeDN}}, @{Name = 'Grantor OU'; Expression ={$GrantorOU}}, @{Name = 'Trustee OU'; Expression ={$TrusteeOU}} | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | out-file $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, SendOnBehalf" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | out-file $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | out-file $ErrorFile -Append
                        "" | out-file $ErrorFile -Append
                    }
                    "" | out-file $ErrorFile -Append; "" | out-file $ErrorFile -Append
                }
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
            }


            # Managed By
            if (($ExportManagedBy -eq $true) -and ($GrantorRecipientType -Match 'group')) {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', ManagedBy'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    try {
                        if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                        $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-Recipient -identity $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $GrantorDN) | where-object {$_.ManagedBy -ne ''}
                    } catch {
                    }
                } else {
                    try {
                        $TrusteeIdentityQuery = (Get-Recipient -identity $GrantorDN -resultsize 1 -wa stop -ea stop| where-object {$_.ManagedBy -ne ''} | select-object *, @{Name="trustee";Expression={$_."user"}})
                    } catch {
                    }
                }
                $TrusteeIdentityOriginal = @($TrusteeIdentityQuery | select-object @{Name = 'Trustee'; Expression = {$_.ManagedBy}})
                if ($error.count -eq 0) {
                    foreach ($TrusteeIdentity in $TrusteeIdentityOriginal.trustee) {
                        if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                        $trusteeidentity = $trusteeidentity.tostring()
                        try {
                            $TrusteeIdentityQuery = (get-user ($TrusteeIdentity) -resultsize 1 -wa stop -ea stop)
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'Cloud'} else {$TrusteeCloudOrOnPrem = 'On-Prem'}
                            if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'On-Prem'}
                        } catch {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            try {
                                $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $TrusteeIdentity)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'On-Prem'} else {$TrusteeCloudOrOnPrem = 'Cloud'}
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'Cloud'}
                            } catch {
                                continue
                            }
                        }
                        $error.clear()
                        $TrusteeIdentityOriginal = $TrusteeIdentity
                        $TrusteeRecipientType = $null
                        $TrusteeRecipientTypeDetails = $null
                        $TrusteeDisplayName = $null
                        $TrusteePrimarySMTP = $null
                        $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                        if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                        $TrusteeIdentityQuery | select-object @{name = 'Grantor Primary SMTP'; expression = {$GrantorPrimarySMTP}}, @{name = 'Grantor Display Name'; expression = {$GrantorDisplayName}}, @{name = 'Grantor Recipient Type'; expression = {$GrantorRecipientType + '/' + $GrantorRecipientTypeDetails}}, @{name = 'Grantor Environment'; expression = {$GrantorCloudOrOnPrem}}, @{Name = 'Trustee Primary SMTP'; Expression = {$TrusteePrimarySMTP}}, @{Name = 'Trustee Display Name'; Expression = {$TrusteeDisplayName}}, @{Name = 'Trustee Original Identity'; Expression = {$TrusteeIdentityOriginal}}, @{name = 'Trustee Recipient Type'; expression = {$TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails}}, @{name = 'Trustee Environment'; expression = {$TrusteeCloudOrOnPrem}}, @{Name = 'Permission(s)'; Expression = {'ManagedBy'}}, @{Name = 'Folder Name'; Expression = {''}}, @{Name = 'Grantor LegacyExchangeDN'; Expression ={$GrantorLegacyExchangeDN}}, @{Name = 'Grantor OU'; Expression ={$GrantorOU}}, @{Name = 'Trustee OU'; Expression ={$TrusteeOU}} | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | out-file $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, ManagedBy" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | out-file $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | out-file $ErrorFile -Append
                        "" | out-file $ErrorFile -Append
                    }
                    "" | out-file $ErrorFile -Append; "" | out-file $ErrorFile -Append
                }
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
            }


            # Folder permissions
            if (($ExportFolderPermissions -eq $true) -and ($GrantorRecipientType -NotMatch 'group')) {
                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                $ErrorActionPreference = "SilentlyContinue"
                $WarningPreference = "SilentlyContinue"
                $error.clear()
                $Text += ', Folders'
                if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                    if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                    try {
                        $Folders = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-MailboxFolderStatistics -identity $($args[0])} -argumentlist $GrantorDN) | foreach-object {$_.folderpath} | foreach-object{$_.replace('/','\')}
                    } catch {
                    }
                } else {
                    try {
                        $Folders = Get-MailboxFolderStatistics -identity $GrantorDN | foreach-object {$_.folderpath} | foreach-object{$_.replace('/','\')}
                    } catch {
                    }
                }
                $FolderCount = 1
                if ($error.count -eq 0) {
                    ForEach ($Folder in $Folders) {
                        $FolderKey = $Alias + ':' + $Folder
                        $Permissions = $null
                        if ($GrantorCloudOrOnPrem -eq 'Cloud') {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            $Permissions = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {Get-MailboxFolderPermission -identity $($args[0]) -wa silentlycontinue -ea silentlycontinue} -argumentlist $FolderKey) | where-object {$_.user.usertype -notlike 'Default' -and $_.user.usertype -notlike 'Anonymous' -and $_.user.displayname -notlike $Recipient.DisplayName}
                        } else {
                            $Permissions = Get-MailboxFolderPermission -identity $FolderKey -wa silentlycontinue -ea silentlycontinue | where-object {$_.user.usertype -notlike 'Default' -and $_.user.usertype -notlike 'Anonymous' -and $_.user.displayname -notlike $Recipient.DisplayName}
                        }
                        if ($permissions -eq $null) {continue}
                        foreach ($TrusteeIdentity in $Permissions.user) {
                            if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                            $trusteeidentity = $trusteeidentity.tostring()
                            try {
                                $TrusteeIdentityQuery = (get-recipient ($TrusteeIdentity) -resultsize 1 -wa stop -ea stop)
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'Cloud'} else {$TrusteeCloudOrOnPrem = 'On-Prem'}
                                if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'On-Prem'}
                            } catch {
                                if ($ExportFromCloud -eq $true) {Connect-ExchangeOnline}
                                try {
                                    $TrusteeIdentityQuery = (Invoke-Command -Session $script:BatchSessionCloud -ScriptBlock {get-recipient $($args[0]) -resultsize 1 -wa stop -ea stop} -argumentlist $TrusteeIdentity)
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "Remote*") {$TrusteeCloudOrOnPrem = 'On-Prem'} else {$TrusteeCloudOrOnPrem = 'Cloud'}
                                    if ($TrusteeIdentityQuery.RecipientTypeDetails -like "*Group") {$TrusteeCloudOrOnPrem = 'Cloud'}
                                } catch {
                                    continue
                                }
                            }
                            $error.clear()
                            $TrusteeIdentityOriginal = $TrusteeIdentity
                            $TrusteeRecipientType = $null
                            $TrusteeRecipientTypeDetails = $null
                            $TrusteeDisplayName = $null
                            $TrusteePrimarySMTP = $null
                            $TrusteeOU = $TrusteeIdentityQuery.OrganizationalUnit
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientType") { $TrusteeRecipientType = $TrusteeIdentityQuery.recipienttype.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "RecipientTypeDetails") { $TrusteeRecipientTypeDetails = $TrusteeIdentityQuery.recipienttypeDetails.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "DisplayName") { $TrusteeDisplayName = $TrusteeIdentityQuery.displayname.tostring() }
                            if ($TrusteeIdentityQuery.PSobject.Properties.name -contains "PrimarySMTPAddress") { $TrusteePrimarySMTP = $TrusteeIdentityQuery.primarysmtpaddress.tostring() }
                            $TrusteeIdentityQuery | select-object @{name = 'Grantor Primary SMTP'; expression = {$GrantorPrimarySMTP}}, @{name = 'Grantor Display Name'; expression = {$GrantorDisplayName}}, @{name = 'Grantor Recipient Type'; expression = {$GrantorRecipientType + '/' + $GrantorRecipientTypeDetails}}, @{name = 'Grantor Environment'; expression = {$GrantorCloudOrOnPrem}}, @{Name = 'Trustee Primary SMTP'; Expression = {$TrusteePrimarySMTP}}, @{Name = 'Trustee Display Name'; Expression = {$TrusteeDisplayName}}, @{Name = 'Trustee Original Identity'; Expression = {$TrusteeIdentityOriginal}}, @{name = 'Trustee Recipient Type'; expression = {$TrusteeRecipientType + '/' + $TrusteeRecipientTypeDetails}}, @{name = 'Trustee Environment'; expression = {$TrusteeCloudOrOnPrem}}, @{Name = 'Permission(s)'; Expression = {[string]::join(', ', @($Permissions | where-object {$_.User -like $trusteeidentity}).accessrights)}}, @{Name = 'Folder Name'; Expression = {$Folder}}, @{Name = 'Grantor LegacyExchangeDN'; Expression ={$GrantorLegacyExchangeDN}}, @{Name = 'Grantor OU'; Expression ={$GrantorOU}}, @{Name = 'Trustee OU'; Expression ={$TrusteeOU}} | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
                        }
                        $FolderCount++
                    }
                }
                if ($error) {
                    $ErrorCount++
                    $Text += ' #white:red#ERROR#'
                    "==============================" | out-file $ErrorFile -Append
                    ("{0:000000}/{1:000000}: {2}, FolderPermissions" -f $count, $RecipientCount, $GrantorPrimarySMTP).tostring().toupper() | out-file $ErrorFile -Append
                    for ($e = ($error.count - 1); $e -ge 0; $e--) {
                        $error[$e] | out-file $ErrorFile -Append
                        "" | out-file $ErrorFile -Append
                    }
                    "" | out-file $ErrorFile -Append; "" | out-file $ErrorFile -Append
                $ErrorActionPreference = "Continue"
                $WarningPreference = "Continue"
                $error.clear()
                }
            }
            $count++
            $text | out-file ($Exportfile + "_Job") -Append -Force
            [System.GC]::Collect() # garbage collection
        }
        if (($ExportFromCloud -eq $true)) {
            Remove-PSSession $script:BatchSessionCloud
        }
        write-host "Done."
    } -Name ("$RecipientStartID" + "_Job") -ArgumentList $RecipientStartID, $RecipientEndID, $Exportfile, $ErrorFile, $TempRecipientFile, $ExportFromOnPrem, $ExportFromCloud, $CredentialPasswordFile, $CredentialUsernameFile, $ExportAccessRights, $ExportSendAs, $ExportSendOnBehalf, $ExportManagedBy, $ExportFolderPermissions, $ExportFullAccessPerTrustee, $TranscriptFile | out-null
    $Batch = $Batch + 1
}

# Wait for all remaining jobs to complete and results are ready to be received
while($true) {
    foreach ($x in (Get-Job -state Completed)) {
        $temp = $null
        $TempPath = $null
        # show temp job output file, delete output file
        $TempPath = ($Exportfile + '_temp' + $x.Name)
        if (test-path $TempPath) {
            $temp = get-content $TempPath
            $temp | Write-HostColored
            Remove-Item $TempPath -force

        }
        $temp = $null
        $TempPath = $null

        # append temp error file and delete temp file
        $TempPath = ($Errorfile + '_temp' + $x.Name)
        $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
        if (Test-Path $TempPath) {
            $temp = get-content $TempPath
            $temp | Out-File $Errorfile -Append -Force
            Remove-Item $TempPath -force
        }
        $temp = $null
        $TempPath = $null

        # append temp transcript file and delete temp file
        $TempPath = ($Transcriptfile + '_temp' + $x.Name)
        $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
        if (Test-Path $TempPath) {
            $temp = get-content $TempPath
            $temp | Out-File $Transcriptfile -Append -Force
            Remove-Item $TempPath -force
        }
        $temp = $null
        $TempPath = $null

        # append temp export file and delete temp file
        $TempPath = ($Exportfile + '_temp' + $x.Name)
        $TempPath = $TempPath.substring(0, ($tempPath.length - 4))
        if (Test-Path $TempPath) {
            $temp = import-csv $TempPath -delimiter ";"
            $temp | export-csv $ExportFile -Append -NoTypeInformation -Force -Delimiter ';'
            Remove-Item $TempPath -force
        }
        $temp = $null
        $TempPath = $null
        Remove-Job -job $x -force
    }
    [System.GC]::Collect() # garbage collection
    start-sleep -s 5

    # end loop when no more completed jobs and no more running jobs are left
    if ((@(Get-Job -state running).count -eq 0) -and (@(Get-Job -state completed).count -eq 0)) { break }
}
if (($ExportAccessRights -eq $true) -and ($ExportFullAccessPerTrustee -eq $true)) {
    write-host 'Creating full access permission files per trustee.'
    $AllowedChars = @("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    $PrimarySMTPAddressesToIgnore = @("xxx@domain.com", "yyy@domain.com") #List of primary SMTP addresses to ignore (service account, for example). Wildcards are not allowed.
    $RecipientPermissions = import-csv $ExportFile -Delimiter ';' | Select-Object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name', 'Permission(s)', 'Grantor Environment', 'Trustee Environment' | Sort-Object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name', 'Permission(s)', 'Grantor Environment', 'Trustee Environment'
    for ($x = 0; $x -lt $RecipientPermissions.count; $x++) {
        if (($RecipientPermissions[$x].'Permission(s)' -like "*FullAccess*") -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne '') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -ne $RecipientPermissions[$x].'Grantor Primary SMTP') -and ($RecipientPermissions[$x].'Trustee Primary SMTP' -notin $PrimarySMTPAddressesToIgnore) -and ($RecipientPermissions[$x].'Grantor Primary SMTP' -notin $PrimarySMTPAddressesToIgnore) -and ($RecipientPermissions[$x].'Grantor Environment' -ne $RecipientPermissions[$x].'Trustee Environment')) {
            if ($AllowedChars.Contains($RecipientPermissions[$x].'Trustee Primary SMTP'.substring(0, 1).tolower()) -eq $true) {
                $FileName = 'prefix_' + $RecipientPermissions[$x].'Trustee Primary SMTP'.substring(0, 1).tolower() + '.csv'
            } else {
                $FileName = 'prefix__.csv'
            }

            $RecipientPermissions[$x].'Trustee Primary SMTP' = $RecipientPermissions[$x].'Trustee Primary SMTP'.ToLower()
            $RecipientPermissions[$x].'Grantor Primary SMTP' = $RecipientPermissions[$x].'Grantor Primary SMTP'.ToLower()
            $RecipientPermissions[$x].'Grantor LegacyExchangeDN' = $RecipientPermissions[$x].'Grantor LegacyExchangeDN'
            $RecipientPermissions[$x].'Grantor Display Name' = $RecipientPermissions[$x].'Grantor Display Name'
            $RecipientPermissions[$x] | select-object 'Trustee Primary SMTP', 'Grantor Primary SMTP', 'Grantor LegacyExchangeDN', 'Grantor Display Name' | export-csv ((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee\' + $FileName) -append -force -notypeinformation -delimiter ";"
        }
    }

    if ($TargetFolder -ne "") {
        if (test-path $TargetFolder) {
            Get-ChildItem ((Split-Path -Path $Exportfile) + "\FullAccessPerTrustee\") -Filter 'prefix_*.csv' -file | Foreach-Object {
                $x = import-csv $_.fullname -delimiter ";" | select-object * -unique
                $x | export-csv $_.fullname -NoTypeInformation -Force -Delimiter ';'
                $x = $null
                if (test-path ($TargetFolder + "\" + $_.Name)) {
                    # File exists at target, compare MD5 hashes with source.
                    if ((Get-FileHash $_.FullName -Algorithm MD5).hash -eq (Get-FileHash ($TargetFolder + '\' + $_.Name) -Algorithm MD5).hash) {
                        # MD5 hashes are equal, file does not need to be copied
                    } else {
                        # MD5 hashes are not equal, file needs to be copied.
                        copy-item $_.fullname $TargetFolder -force
                    }
                } else {
                    # File does not exist at target, copy file.
                    copy-item $_.fullname $TargetFolder -force
                }
            }

            Get-ChildItem $TargetFolder -Filter 'prefix_*.csv' -file | Foreach-Object {
                 if (-not (test-path (((Split-Path -Path $Exportfile) + '\FullAccessPerTrustee\' + $_.Name)))) {
                    # File does not exist at source. Delete at target.
                    Remove-Item $_.FullName -force
                }
            }
        } else {
            write-host "Folder $TargetFolder does not exist."
        }
    }
}

write-host 'Cleaning output file.'
$RecipientPermissions = import-csv $ExportFile -Delimiter ';' | Select-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission(s)', 'Folder Name', 'Grantor OU', 'Trustee OU' | Sort-Object 'Grantor Primary SMTP', 'Grantor Display Name', 'Grantor Recipient Type', 'Grantor Environment', 'Trustee Primary SMTP', 'Trustee Display Name', 'Trustee Recipient Type', 'Trustee Environment', 'Trustee Original Identity', 'Permission(s)', 'Folder Name', 'Grantor OU', 'Trustee OU'
$RecipientPermissions | export-csv $ExportFile -NoTypeInformation -Force -Delimiter ';'

if (Test-Path $TempRecipientFile) {(Remove-Item $TempRecipientFile -force)}

stop-transcript
$TempPath = ($Transcriptfile + '_temp')
if (Test-Path $TempPath) {
    $temp = get-content $TempPath
    $temp | out-file $TranscriptFile -Append -Force
    Remove-Item $TempPath -force
}
$temp = $null
$TempPath = $null

write-host 'Script completed.'