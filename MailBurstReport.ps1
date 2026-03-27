<#
MAIL BURST REPORT - CSV (Excel-friendly) - PowerShell 5.1 compatible

Purpose:
- Looks back over the last X minutes
- Groups messages by Sender + Subject
- Reports only if message count reaches the defined threshold
- Excludes messages from configured internal sender domains
- Writes results to a CSV file
- Writes nothing if no threshold is exceeded
#>

param(
    [int]$WindowMinutes = 1440,
    [int]$Threshold = 20,
    [string]$OutDir = "C:\Reports",
    [string]$BaseName = "MailBurstReport",
    [ValidateSet("RECEIVE","SEND","DELIVER")]
    [string]$EventId = "RECEIVE",
    [string[]]$ExcludeSenderDomains = @("internal.example")
)

function Normalize-Subject {
    <#
    .SYNOPSIS
    Normalizes a message subject for grouping.

    .DESCRIPTION
    Removes common reply/forward prefixes such as RE:, FW:, and FWD:.
    Also normalizes repeated whitespace so similar subjects are grouped together.

    .PARAMETER Subject
    The original message subject from the tracking log.

    .OUTPUTS
    System.String
    A cleaned subject string. Returns "(no subject)" if the value is empty.

    .EXAMPLE
    Normalize-Subject -Subject "RE:   FW: Monthly Report"
    #>
    param([string]$Subject)

    if ([string]::IsNullOrWhiteSpace($Subject)) {
        return "(no subject)"
    }

    $s = $Subject.Trim()

    while ($s -match '^(RE|FW|FWD)\s*:\s*') {
        $s = ($s -replace '^(RE|FW|FWD)\s*:\s*', '').Trim()
    }

    $s = ($s -replace '\s+', ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($s)) {
        return "(no subject)"
    }

    return $s
}

function Is-ExcludedSender {
    <#
    .SYNOPSIS
    Checks whether a sender belongs to an excluded domain list.

    .DESCRIPTION
    Compares the sender address against the configured excluded domains.
    This is useful for skipping internal or otherwise trusted sender domains.

    .PARAMETER Sender
    The sender value from the message tracking log entry.

    .PARAMETER Domains
    A list of domains that should be excluded from the report.

    .OUTPUTS
    System.Boolean
    Returns $true if the sender matches an excluded domain; otherwise $false.

    .EXAMPLE
    Is-ExcludedSender -Sender "user@internal.example" -Domains @("internal.example")
    #>
    param(
        [object]$Sender,
        [string[]]$Domains
    )

    if ($Sender -eq $null) {
        return $true
    }

    $senderStr = $Sender.ToString().ToLower()

    foreach ($d in $Domains) {
        if ([string]::IsNullOrWhiteSpace($d)) {
            continue
        }

        $domain = $d.ToLower().Trim()

        if ($senderStr -like "*@$domain") {
            return $true
        }
    }

    return $false
}

# Define time window
$EndTime   = Get-Date
$StartTime = $EndTime.AddMinutes(-1 * $WindowMinutes)

# Create a unique output file name for each run
$runStamp = $EndTime.ToString("yyyy-MM-dd_HHmmss")
$OutCsv   = Join-Path $OutDir ("{0}_{1}.csv" -f $BaseName, $runStamp)

# Ensure output directory exists
if (-not (Test-Path $OutDir)) {
    New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
}

# Read message tracking logs
$logs = Get-MessageTrackingLog `
    -Start $StartTime `
    -End $EndTime `
    -EventId $EventId `
    -ResultSize Unlimited `
    -ErrorAction SilentlyContinue

# Exclude configured sender domains
$logs = $logs | Where-Object {
    -not (Is-ExcludedSender -Sender $_.Sender -Domains $ExcludeSenderDomains)
}

# Group by Sender + normalized Subject
$bursts = $logs | ForEach-Object {

    $senderValue = $_.Sender
    if ($senderValue -eq $null -or [string]::IsNullOrWhiteSpace($senderValue.ToString())) {
        $senderValue = "(unknown)"
    }

    [PSCustomObject]@{
        Sender  = $senderValue.ToString()
        Subject = Normalize-Subject -Subject $_.MessageSubject
    }

} | Group-Object -Property Sender, Subject |
    Where-Object { $_.Count -ge $Threshold } |
    Sort-Object -Property Count -Descending

# If nothing is found, exit without creating output
if ($bursts -eq $null -or $bursts.Count -eq 0) {
    exit 0
}

# Build export rows
$rows = foreach ($b in $bursts) {
    [PSCustomObject]@{
        ReportTime       = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
        WindowStart      = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
        WindowEnd        = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
        EventId          = $EventId
        Threshold        = $Threshold
        Count            = $b.Count
        Sender           = $b.Group[0].Sender
        Subject          = $b.Group[0].Subject
        ExcludedDomains  = ($ExcludeSenderDomains -join ",")
    }
}

# Export result to CSV
$rows | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
