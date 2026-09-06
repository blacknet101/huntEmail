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
    $clearLine = "`r" + (" " * 120) + "`r"

    Write-Host $clearLine -NoNewline
    Write-Host "`r$Activity " -NoNewline -ForegroundColor Cyan

    for ($i = 0; $i -lt $filledWidth; $i++) {
        $colorIndex = $i % $rainbowColors.Count
        Write-Host ([char]0x2588) -NoNewline -ForegroundColor $rainbowColors[$colorIndex]
    }

    Write-Host $nyanCat -NoNewline -ForegroundColor Yellow

    Write-Host ([string]([char]0x2591) * $emptyWidth) -NoNewline -ForegroundColor DarkGray

    Write-Host " $PercentComplete%" -NoNewline -ForegroundColor White

    if ($Status) {
        Write-Host " - $Status" -NoNewline -ForegroundColor Gray
    }
}

function Write-NyanComplete {
    param ([string]$Activity = "Processing")
    Write-Host ""
    Write-Host "$Activity complete! =^.^=" -ForegroundColor Green
    Write-Host ""
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

function timelineEmail {
    <#
    .SYNOPSIS
        Builds a delivery timeline for a phishing email or IOC across the org.

    .DESCRIPTION
        Given search criteria (sender, subject, date range), enumerates all mailboxes
        and builds a chronological timeline showing who received the email, when it
        was delivered, whether it was read, replied to, or forwarded.

        Accepts piped input from huntEmail results to build a timeline from
        previously found messages.

        Requires an active MS Graph connection with Mail.Read and User.Read.All
        permissions (Application).

    .PARAMETER SearchResults
        Piped results from huntEmail. If provided, builds timeline from those results.

    .PARAMETER Sender
        Email address of the sender to build timeline for.

    .PARAMETER Subject
        Subject line to match.

    .PARAMETER StartDate
        Start of the date range (default: last 48 hours).

    .PARAMETER EndDate
        End of the date range (default: now).

    .EXAMPLE
        timelineEmail -Sender "attacker@evil.com" -Subject "Urgent Invoice"

    .EXAMPLE
        huntEmail -Sender "phish@bad.com" | timelineEmail

    .EXAMPLE
        timeline -Sender "noreply@suspicious.com" -StartDate "2026-02-15" | Export-Csv timeline.csv

    .NOTES
        Author: Security Operations
        Created: 2/19/2026
        Part of: Email Threat Response Toolkit
    #>

    [CmdletBinding()]
    [Alias("timeline")]
    param (
        [Parameter(ValueFromPipeline)]
        [PSObject[]]$SearchResults,

        [Parameter()]
        [string]$Sender,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [datetime]$StartDate = (Get-Date).AddHours(-48),

        [Parameter()]
        [datetime]$EndDate = (Get-Date),

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

        $PipedResults = [System.Collections.Generic.List[PSObject]]::new()
        $Timeline = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        if ($SearchResults) {
            foreach ($R in $SearchResults) { $PipedResults.Add($R) }
        }
    }

    end {
        $AllResults = $null

        if ($PipedResults.Count -gt 0) {
            Write-Host "[TIMELINE] Building from $($PipedResults.Count) search results..." -ForegroundColor Cyan
            $AllResults = $PipedResults
        }
        elseif ($Sender -or $Subject) {
            Write-Host "[TIMELINE] Searching org for timeline data..." -ForegroundColor Cyan
            $SearchParams = @{
                StartDate = $StartDate
                EndDate   = $EndDate
                MaxResults = 0
            }
            if ($Sender) { $SearchParams['Sender'] = $Sender }
            if ($Subject) { $SearchParams['Subject'] = $Subject }

            $AllResults = huntEmail @SearchParams

            if (-not $AllResults -or $AllResults.Count -eq 0) {
                Write-Host "[TIMELINE] No matching emails found to build timeline." -ForegroundColor Yellow
                return
            }
        }
        else {
            Write-Host "[ERROR] Provide -Sender/-Subject or pipe results from huntEmail" -ForegroundColor Red
            return
        }

        $FolderCache = @{}
        $Total = $AllResults.Count
        $Current = 0

        foreach ($Result in $AllResults) {
            $Current++
            $Pct = [math]::Round(($Current / $Total) * 100)
            if ($Nyan) {
                Write-NyanProgress -PercentComplete $Pct -Status "[$Current/$Total] $($Result.Mailbox)" -Activity "Timeline"
            } else {
                Write-Progress -Activity "Building timeline" -Status "[$Current/$Total] $($Result.Mailbox)" -PercentComplete $Pct
            }

            $FolderName = if ($Result.Folder -and $Result.Folder -ne "Unknown") {
                $Result.Folder
            }
            elseif ($Result.FolderId) {
                if (-not $FolderCache.ContainsKey($Result.Mailbox)) {
                    try {
                        $FolderCache[$Result.Mailbox] = Get-HuntEmailFolderMap -UserPrincipalName $Result.Mailbox
                    }
                    catch {
                        $FolderCache[$Result.Mailbox] = @{}
                    }
                }
                if ($FolderCache[$Result.Mailbox][$Result.FolderId]) {
                    $FolderCache[$Result.Mailbox][$Result.FolderId]
                } else { "Unknown" }
            }
            else { "Unknown" }

            $Actions = @()
            if ($Result.IsRead) { $Actions += "READ" }
            if ($FolderName -eq "Deleted Items") { $Actions += "DELETED" }
            if ($FolderName -eq "Junk Email") { $Actions += "JUNKED" }
            if ($FolderName -eq "Sent Items") { $Actions += "SENT" }
            if ($FolderName -eq "Archive") { $Actions += "ARCHIVED" }

            $Timeline.Add([PSCustomObject]@{
                Timestamp   = $Result.Received
                Mailbox     = $Result.Mailbox
                From        = $Result.From
                Subject     = $Result.Subject
                Folder      = $FolderName
                IsRead      = $Result.IsRead
                HasAttach   = $Result.HasAttach
                Actions     = if ($Actions.Count -gt 0) { $Actions -join ", " } else { "DELIVERED" }
                MessageId   = $Result.MessageId
                FolderId    = $Result.FolderId
            })
        }

        if ($Nyan) {
            Write-NyanComplete -Activity "Timeline"
        } else {
            Write-Progress -Activity "Building timeline" -Completed
        }

        $Sorted = $Timeline | Sort-Object Timestamp

        # --- Display ---
        $ReadCount = @($Sorted | Where-Object { $_.IsRead }).Count
        $UnreadCount = @($Sorted | Where-Object { -not $_.IsRead }).Count
        $UniqueMailboxes = ($Sorted | Select-Object -ExpandProperty Mailbox -Unique).Count

        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host " EMAIL DELIVERY TIMELINE" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan

        if ($Sorted.Count -gt 0) {
            $First = $Sorted[0].Timestamp
            $Last = $Sorted[-1].Timestamp
            Write-Host "  From:        $($Sorted[0].From)" -ForegroundColor White
            Write-Host "  Subject:     $($Sorted[0].Subject)" -ForegroundColor White
            Write-Host "  First seen:  $First" -ForegroundColor White
            Write-Host "  Last seen:   $Last" -ForegroundColor White
            Write-Host "  Spread:      $(($Last - $First).ToString())" -ForegroundColor White
        }

        Write-Host "  Total hits:  $($Sorted.Count)" -ForegroundColor White
        Write-Host "  Mailboxes:   $UniqueMailboxes" -ForegroundColor White
        Write-Host "  Read:        $ReadCount" -ForegroundColor $(if ($ReadCount -gt 0) { "Red" } else { "Green" })
        Write-Host "  Unread:      $UnreadCount" -ForegroundColor $(if ($UnreadCount -gt 0) { "Yellow" } else { "Green" })
        Write-Host ""

        $Sorted | Format-Table Timestamp, Mailbox, Folder, Actions, HasAttach -AutoSize | Out-Host

        if ($ReadCount -gt 0) {
            Write-Host "  WARNING: $ReadCount user(s) opened this email." -ForegroundColor Red
            $ReadUsers = ($Sorted | Where-Object { $_.IsRead } | Select-Object -ExpandProperty Mailbox -Unique) -join ", "
            Write-Host "  Users who read it: $ReadUsers" -ForegroundColor Red
        }

        Write-Host "`n  Tip: Pipe to purgeEmail to purge all matched messages" -ForegroundColor DarkGray

        return $Sorted
    }
}
