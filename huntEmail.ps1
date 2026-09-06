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
