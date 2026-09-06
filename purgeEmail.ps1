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

function purgeEmail {
    <#
    .SYNOPSIS
        Hard deletes emails from all folders in a user's mailbox via MS Graph.

    .DESCRIPTION
        Permanently removes emails from ALL mailbox folders including Inbox, Sent Items,
        Archive, Deleted Items, Junk, and any custom/nested folders using the
        Microsoft Graph permanent delete action.

        Accepts piped input from huntEmail or timelineEmail to purge previously
        found messages. Can also target specific messages by MessageId or search criteria.

        REQUIRES CONFIRMATION before executing. Use -Force to skip confirmation (use with caution).

        Requires: Mail.ReadWrite and User.Read.All permissions (Application).

    .PARAMETER SearchResults
        Piped results from huntEmail or timelineEmail containing Mailbox and MessageId.

    .PARAMETER UserPrincipalName
        Target mailbox(es) to purge from.

    .PARAMETER MessageId
        Specific message ID(s) to delete.

    .PARAMETER Sender
        Delete all emails from this sender in the target mailbox(es).

    .PARAMETER Subject
        Delete all emails matching this subject.

    .PARAMETER StartDate
        Start of date range for matching emails.

    .PARAMETER EndDate
        End of date range for matching emails.

    .PARAMETER MaxResults
        Maximum results per mailbox when building purge targets from search criteria.
        Use 0 for no limit.

    .PARAMETER Force
        Skip confirmation prompt. Use with extreme caution.

    .EXAMPLE
        huntEmail -Sender "attacker@evil.com" | purgeEmail

    .EXAMPLE
        huntEmail -Sender "phish@bad.com" | timelineEmail | purgeEmail

    .EXAMPLE
        purge -UserPrincipalName "user@company.com" -Sender "attacker@evil.com" -Force

    .NOTES
        Author: Security Operations
        Created: 2/19/2026
        Part of: Email Threat Response Toolkit
        THIS FUNCTION PERFORMS HARD DELETES. Messages cannot be recovered after purge.
    #>

    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [Alias("purge")]
    param (
        [Parameter(ValueFromPipeline)]
        [PSObject[]]$SearchResults,

        [Parameter()]
        [string[]]$UserPrincipalName,

        [Parameter()]
        [string[]]$MessageId,

        [Parameter()]
        [string]$Sender,

        [Parameter()]
        [string]$Subject,

        [Parameter()]
        [datetime]$StartDate,

        [Parameter()]
        [datetime]$EndDate,

        [Parameter()]
        [int]$MaxResults = 0,

        [Parameter()]
        [switch]$Force,

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
        $PurgeTargets = [System.Collections.Generic.List[PSObject]]::new()
        $PurgeLog = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        if ($SearchResults) {
            foreach ($R in $SearchResults) { $PipedResults.Add($R) }
        }
    }

    end {
        # --- Build purge target list ---
        if ($PipedResults.Count -gt 0) {
            foreach ($R in $PipedResults) {
                if ($R.Mailbox -and $R.MessageId) {
                    $PurgeTargets.Add([PSCustomObject]@{
                        Mailbox   = $R.Mailbox
                        MessageId = $R.MessageId
                        Subject   = $R.Subject
                        From      = $R.From
                    })
                }
            }
        }
        elseif ($UserPrincipalName -and $MessageId) {
            foreach ($UPN in $UserPrincipalName) {
                foreach ($MId in $MessageId) {
                    $PurgeTargets.Add([PSCustomObject]@{
                        Mailbox   = $UPN
                        MessageId = $MId
                        Subject   = "(direct ID)"
                        From      = "(direct ID)"
                    })
                }
            }
        }
        elseif ($Sender -or $Subject) {
            Write-Host "[PURGE] Searching for matching emails first..." -ForegroundColor Cyan
            $SearchParams = @{}
            if ($Sender) { $SearchParams['Sender'] = $Sender }
            if ($Subject) { $SearchParams['Subject'] = $Subject }
            if ($StartDate) { $SearchParams['StartDate'] = $StartDate }
            if ($EndDate) { $SearchParams['EndDate'] = $EndDate }
            if ($UserPrincipalName) { $SearchParams['UserPrincipalName'] = $UserPrincipalName }
            $SearchParams['MaxResults'] = $MaxResults

            $Results = huntEmail @SearchParams

            if (-not $Results -or $Results.Count -eq 0) {
                Write-Host "[PURGE] No matching emails found. Nothing to purge." -ForegroundColor Green
                return
            }

            foreach ($R in $Results) {
                $PurgeTargets.Add([PSCustomObject]@{
                    Mailbox   = $R.Mailbox
                    MessageId = $R.MessageId
                    Subject   = $R.Subject
                    From      = $R.From
                })
            }
        }
        else {
            Write-Host "[ERROR] Provide search results via pipe, -MessageId, or -Sender/-Subject criteria" -ForegroundColor Red
            return
        }

        if ($PurgeTargets.Count -eq 0) {
            Write-Host "[PURGE] No targets identified. Nothing to purge." -ForegroundColor Green
            return
        }

        # --- Confirmation ---
        $UniqueMailboxes = ($PurgeTargets | Select-Object -ExpandProperty Mailbox -Unique).Count

        Write-Host "`n========================================" -ForegroundColor Red
        Write-Host " PURGE CONFIRMATION" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "  Messages to delete: $($PurgeTargets.Count)" -ForegroundColor Yellow
        Write-Host "  Across mailboxes:   $UniqueMailboxes" -ForegroundColor Yellow
        if ($PurgeTargets[0].From -ne "(direct ID)") {
            Write-Host "  From:               $($PurgeTargets[0].From)" -ForegroundColor Yellow
            Write-Host "  Subject:            $($PurgeTargets[0].Subject)" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "  THIS IS A HARD DELETE. MESSAGES CANNOT BE RECOVERED." -ForegroundColor Red
        Write-Host ""

        if (-not $Force -and -not $WhatIfPreference) {
            $Confirm = Read-Host "  Type 'PURGE' to confirm deletion"
            if ($Confirm -ne "PURGE") {
                Write-Host "`n[ABORTED] Purge cancelled by user." -ForegroundColor Yellow
                return
            }
        }

        # --- Execute purge ---
        Write-Host "`n[PURGING] Starting hard delete..." -ForegroundColor Red

        $Total = $PurgeTargets.Count
        $Current = 0
        $SuccessCount = 0
        $FailCount = 0
        $SkipCount = 0

        foreach ($Target in $PurgeTargets) {
            $Current++
            $Pct = [math]::Round(($Current / $Total) * 100)
            if ($Nyan) {
                Write-NyanProgress -PercentComplete $Pct -Status "[$Current/$Total] $($Target.Mailbox)" -Activity "Purging"
            } else {
                Write-Progress -Activity "Purging emails" -Status "[$Current/$Total] $($Target.Mailbox)" -PercentComplete $Pct
            }

            try {
                $ActionTarget = "$($Target.Mailbox) :: $($Target.Subject) [$($Target.MessageId)]"
                if (-not $PSCmdlet.ShouldProcess($ActionTarget, "Permanently delete message")) {
                    $SkipCount++
                    $PurgeLog.Add([PSCustomObject]@{
                        Mailbox   = $Target.Mailbox
                        MessageId = $Target.MessageId
                        Subject   = $Target.Subject
                        Status    = "Skipped"
                        Error     = ""
                    })
                    continue
                }

                # Delete message (moves to Deleted Items or soft-deletes)
                $DeleteUri = "https://graph.microsoft.com/v1.0/users/$($Target.Mailbox)/messages/$($Target.MessageId)/permanentDelete"
                Invoke-MgGraphRequest -Method POST -Uri $DeleteUri -ErrorAction Stop

                $SuccessCount++
                $PurgeLog.Add([PSCustomObject]@{
                    Mailbox   = $Target.Mailbox
                    MessageId = $Target.MessageId
                    Subject   = $Target.Subject
                    Status    = "Purged"
                    Error     = ""
                })
            }
            catch {
                $FailCount++
                $PurgeLog.Add([PSCustomObject]@{
                    Mailbox   = $Target.Mailbox
                    MessageId = $Target.MessageId
                    Subject   = $Target.Subject
                    Status    = "Failed"
                    Error     = $_.Exception.Message
                })
                Write-Verbose "[FAIL] $($Target.Mailbox) - $($Target.MessageId): $_"
            }
        }

        if ($Nyan) {
            Write-NyanComplete -Activity "Purging"
        } else {
            Write-Progress -Activity "Purging emails" -Completed
        }

        # --- Summary ---
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host " PURGE RESULTS" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  Total processed: $Total" -ForegroundColor White
        Write-Host "  Purged:          $SuccessCount" -ForegroundColor Green
        if ($FailCount -gt 0) {
            Write-Host "  Failed:          $FailCount" -ForegroundColor Red
        }
        Write-Host " Mailboxes:       $UniqueMailboxes" -ForegroundColor White

        if ($FailCount -gt 0) {
            Write-Host "`n  Failed targets:" -ForegroundColor Red
            $PurgeLog | Where-Object { $_.Status -eq "Failed" } | Format-Table Mailbox, Subject, Error -AutoSize | Out-Host
        }

        Write-Host "`n  Purge log returned. Pipe to Export-Csv to save." -ForegroundColor DarkGray

        return $PurgeLog
    }
}
