function Write-NyanProgress {
    param (
        [int]$PercentComplete,
        [string]$Status,
        [string]$Activity = "Processing"
    )
    $rainbowColors = @(
        [ConsoleColor]::Red,
        [ConsoleColor]::DarkYellow,
        [ConsoleColor]::Yellow,
        [ConsoleColor]::Green,
        [ConsoleColor]::Cyan,
        [ConsoleColor]::Blue,
        [ConsoleColor]::Magenta
    )
    $nyanCat = "[=^.^=]"
    $barWidth = 40
    $filledWidth = [math]::Floor(($PercentComplete / 100) * $barWidth)
    $emptyWidth = $barWidth - $filledWidth

    $maxWidth = [Console]::WindowWidth - 1

    Write-Host "`r" -NoNewline
    Write-Host "$Activity " -NoNewline -ForegroundColor Cyan

    for ($i = 0; $i -lt $filledWidth; $i++) {
        $colorIndex = $i % $rainbowColors.Count
        Write-Host ([char]0x2588) -NoNewline -ForegroundColor $rainbowColors[$colorIndex]
    }

    Write-Host $nyanCat -NoNewline -ForegroundColor Yellow
    Write-Host ([string]([char]0x2591) * $emptyWidth) -NoNewline -ForegroundColor DarkGray
    Write-Host " $PercentComplete%" -NoNewline -ForegroundColor White

    $usedWidth = $Activity.Length + 1 + $filledWidth + $nyanCat.Length + $emptyWidth + " $PercentComplete%".Length

    if ($Status) {
        $statusPrefix = " - "
        $availableForStatus = $maxWidth - $usedWidth - $statusPrefix.Length
        if ($availableForStatus -gt 0) {
            $truncStatus = if ($Status.Length -gt $availableForStatus) { $Status.Substring(0, $availableForStatus) } else { $Status }
            Write-Host "$statusPrefix$truncStatus" -NoNewline -ForegroundColor Gray
            $usedWidth += $statusPrefix.Length + $truncStatus.Length
        }
    }

    $remaining = $maxWidth - $usedWidth
    if ($remaining -gt 0) {
        Write-Host (" " * $remaining) -NoNewline
    }
}

function Write-NyanComplete {
    param ([string]$Activity = "Processing")
    Write-Host ""
    Write-Host "$Activity complete! =^.^=" -ForegroundColor Green
    Write-Host ""
}

function Get-HuntEmailPagedCollection {
    param (
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [hashtable]$Headers
    )

    $Items = [System.Collections.Generic.List[object]]::new()
    $NextUri = $Uri

    do {
        $Response = if ($Headers) {
            Invoke-MgGraphRequest -Method GET -Uri $NextUri -Headers $Headers -ErrorAction Stop
        }
        else {
            Invoke-MgGraphRequest -Method GET -Uri $NextUri -ErrorAction Stop
        }

        foreach ($Item in @($Response.value)) {
            $Items.Add($Item)
        }

        $NextUri = $Response.'@odata.nextLink'
    } while ($NextUri)

    return $Items
}

function Get-HuntEmailFolderMap {
    param (
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $FolderMap = @{}
    $PendingUris = [System.Collections.Generic.Queue[string]]::new()
    $PendingUris.Enqueue("https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailFolders?`$select=id,displayName,childFolderCount&`$top=100")

    while ($PendingUris.Count -gt 0) {
        $NextUri = $PendingUris.Dequeue()
        do {
            $Response = Invoke-MgGraphRequest -Method GET -Uri $NextUri -ErrorAction Stop
            foreach ($Folder in @($Response.value)) {
                $FolderMap[$Folder.id] = $Folder.displayName
                if ($Folder.childFolderCount -gt 0) {
                    $PendingUris.Enqueue("https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailFolders/$($Folder.id)/childFolders?`$select=id,displayName,childFolderCount&`$top=100")
                }
            }

            $NextUri = $Response.'@odata.nextLink'
        } while ($NextUri)
    }

    return $FolderMap
}

function huntEmail {
    <#
    .SYNOPSIS
        Searches mailboxes for emails matching specified criteria using MS Graph.

    .DESCRIPTION
        Sweeps across one, multiple, or all user mailboxes to find emails matching
        criteria like sender, subject, date range, keywords, attachment names, or
        InternetMessageId. Uses MS Graph API for fast, flexible searching across the org.

        Results include InternetMessageId (universal across all copies) and attachment
        metadata for forensic fingerprinting.

        Requires an active MS Graph connection with Mail.Read and User.Read.All
        permissions (Application).

    .PARAMETER Sender
        Email address of the sender to search for.

    .PARAMETER Subject
        Subject line to search for (partial match supported).

    .PARAMETER Keywords
        Keywords to search for in the email body.

    .PARAMETER InternetMessageId
        RFC 2822 Message-ID header value. Universal across all copies of the same email.
        Accepts with or without angle brackets. Searches tenant-wide by default.

    .PARAMETER StartDate
        Start of the date range to search (default: last 24 hours).

    .PARAMETER EndDate
        End of the date range to search (default: now).

    .PARAMETER UserPrincipalName
        One or more specific mailboxes to search. If not provided, searches all licensed mailboxes.

    .PARAMETER AttachmentName
        Search for emails with attachments matching this name.

    .PARAMETER MaxResults
        Maximum results per mailbox after filtering (default: 50). Use 0 for no limit.

    .PARAMETER Detailed
        Show full detail for each result including InternetMessageId, attachment info, and body preview.

    .PARAMETER ExportEml
        Path to a folder where full .eml (MIME) exports will be saved for each result.
        Only works with -Detailed. Creates the folder if it doesn't exist.

    .EXAMPLE
        huntEmail -Sender "attacker@evil.com"

    .EXAMPLE
        huntEmail -Sender "phish@bad.com" -StartDate "2026-02-15" -EndDate "2026-02-18"

    .EXAMPLE
        hunt -InternetMessageId "<ABC123@mail.gmail.com>"

    .EXAMPLE
        hunt -Subject "Invoice" -Detailed -ExportEml "C:\IR\case001"

    .EXAMPLE
        hunt -Keywords "password reset" -Sender "noreply@suspicious.com" | Export-Csv results.csv

    .NOTES
        Author: Security Operations
        Created: 2/19/2026
        Updated: 2/20/2026 - Added InternetMessageId search, attachment metadata, EML export
        Part of: Email Threat Response Toolkit
        Permissions: Mail.Read, Mail.ReadWrite, User.Read.All (Application)
    #>

    [CmdletBinding()]
    [Alias("hunt")]
    param (
        [Parameter()]
        [string]$Sender,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [string]$Keywords,

        [Parameter()]
        [string]$InternetMessageId,

        [Parameter()]
        [datetime]$StartDate = (Get-Date).AddHours(-24),

        [Parameter()]
        [datetime]$EndDate = (Get-Date),

        [Parameter()]
        [string[]]$UserPrincipalName,

        [Parameter()]
        [string]$AttachmentName,

        [Parameter()]
        [int]$MaxResults = 50,

        [Parameter()]
        [switch]$Detailed,

        [Parameter()]
        [string]$ExportEml,

        [Parameter()]
        [switch]$Nyan
    )

    begin {
        try {
            $Context = Get-MgContext
            if (-not $Context) { throw "Not connected" }
        }
        catch {
            Write-Host "[ERROR] Not connected to MS Graph. Run:" -ForegroundColor Red
            Write-Host '  . .\Connect-HuntEmailGraph.ps1' -ForegroundColor Yellow
            Write-Host "  or: Connect-MgGraph -TenantId <tenant> -ClientId <client> -CertificateThumbprint <thumb> -NoWelcome" -ForegroundColor Yellow
            Write-Host "  (See documentation\api_info.md and Certificate_Setup.md)" -ForegroundColor DarkGray
            return
        }

        if (-not $Sender -and -not $Subject -and -not $Keywords -and -not $AttachmentName -and -not $InternetMessageId) {
            Write-Host "[ERROR] Provide at least one search criteria: -Sender, -Subject, -Keywords, -AttachmentName, or -InternetMessageId" -ForegroundColor Red
            return
        }

        if ($ExportEml -and -not $Detailed) {
            Write-Host "[WARN] -ExportEml requires -Detailed. Enabling -Detailed automatically." -ForegroundColor Yellow
            $Detailed = [switch]$true
        }

        if ($ExportEml -and -not (Test-Path $ExportEml)) {
            New-Item -Path $ExportEml -ItemType Directory -Force | Out-Null
            Write-Host "[EXPORT] Created export folder: $ExportEml" -ForegroundColor DarkGray
        }

        if ($InternetMessageId) {
            $InternetMessageId = $InternetMessageId.Trim()
            if ($InternetMessageId -notmatch '^<.*>$') {
                $InternetMessageId = "<$InternetMessageId>"
            }
        }

        $ResultsList = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        # --- Build target mailbox list ---
        if ($UserPrincipalName) {
            $Mailboxes = $UserPrincipalName
            Write-Host "[SEARCH] Targeting $($Mailboxes.Count) specified mailbox(es)" -ForegroundColor Cyan
