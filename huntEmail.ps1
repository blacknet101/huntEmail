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
