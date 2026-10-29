# huntEmail.psm1 — Email Threat Response Toolkit
# One download. Dot-source huntEmail.ps1 or Import-Module ./huntEmail.psd1.
#
# Hunt, timeline, and hard-delete phishing or HR-directed mail with Microsoft Graph
# app-only certificate auth. Read Microsoft 365 Copilot prompts and responses.
# This module does not purge Copilot content. Removal stays in Purview eDiscovery.
#
# Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048

Set-StrictMode -Version 1

$script:HuntEmailCopilotPermission = 'AiEnterpriseInteraction.Read.All'
$script:HuntEmailCopilotLearn = 'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions'
$script:HuntEmailCopilotAuditLearn = 'https://learn.microsoft.com/en-us/office/office-365-management-api/copilot-schema'
$script:HuntEmailPlaybookDocs = 'https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048'
$script:HuntEmailInvokeGraph = $null
$script:HuntEmailSearchAuditLog = $null
$script:HuntEmailExchangeConnectionTest = $null

# Interaction-history $filter values published on Learn (Teams Copilot export article).
# Purview eDiscovery item classes are broader (Outlook, PowerPoint, Loop, Whiteboard, ...).
# Those are not sent as $filter values because Learn says other filters are not supported.
$script:HuntEmailCopilotAppClasses = @{
    Word        = 'IPM.SkypeTeams.Message.Copilot.Word'
    Excel       = 'IPM.SkypeTeams.Message.Copilot.Excel'
    Teams       = 'IPM.SkypeTeams.Message.Copilot.Teams'
    BizChat     = 'IPM.SkypeTeams.Message.Copilot.BizChat'
    WebChat     = 'IPM.SkypeTeams.Message.Copilot.WebChat'
    CoworkChat  = 'IPM.SkypeTeams.Message.Copilot.CoworkChat'
}

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

    $maxWidth = 120
    try {
        $windowWidth = [Console]::WindowWidth
        if ($windowWidth -gt 20) { $maxWidth = $windowWidth - 1 }
    }
    catch {
        $maxWidth = 120
    }

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

function Invoke-HuntEmailGraphRequest {
    <#
    .SYNOPSIS
        Internal Graph request wrapper. Production calls Invoke-MgGraphRequest.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [hashtable]$Headers,

        [Parameter()]
        [string]$OutputFilePath
    )

    if ($script:HuntEmailInvokeGraph) {
        return & $script:HuntEmailInvokeGraph -Method $Method -Uri $Uri -Headers $Headers -OutputFilePath $OutputFilePath
    }

    $params = @{
        Method      = $Method
        Uri         = $Uri
        ErrorAction = 'Stop'
    }
    if ($Headers) { $params.Headers = $Headers }
    if ($OutputFilePath) { $params.OutputFilePath = $OutputFilePath }
    return (Invoke-MgGraphRequest @params)
}

function Get-HuntEmailGraphStatusCode {
    param ($ErrorRecord)

    $exception = $ErrorRecord.Exception
    foreach ($candidate in @($exception, $exception.InnerException)) {
        if (-not $candidate) { continue }
        if ($candidate.PSObject.Properties['Response'] -and $candidate.Response -and $candidate.Response.StatusCode) {
            return [int]$candidate.Response.StatusCode
        }
        if ($candidate.PSObject.Properties['StatusCode'] -and $null -ne $candidate.StatusCode) {
            try { return [int]$candidate.StatusCode } catch { Write-Verbose "Graph status code was not numeric." }
        }
    }

    $message = [string]$ErrorRecord.Exception.Message
    if ($message -match 'Response status code does not indicate success:\s*(\d{3})') {
        return [int]$Matches[1]
    }
    if ($message -match '\b(401|403|404)\b') {
        return [int]$Matches[1]
    }
    return $null
}

function Write-HuntEmailGraphConnectionHelp {
    Write-Host "[ERROR] Not connected to MS Graph. Run:" -ForegroundColor Red
    Write-Host "  Import-Module .\huntEmail.psd1" -ForegroundColor Yellow
    Write-Host "  Connect-HuntEmailGraph" -ForegroundColor Yellow
    Write-Host "  or: .\Connect-HuntEmailGraph.ps1" -ForegroundColor Yellow
    Write-Host "  or: Connect-MgGraph -TenantId <tenant> -ClientId <client> -CertificateThumbprint <thumb> -NoWelcome" -ForegroundColor Yellow
    Write-Host "  (See documentation\api_info.md and Certificate_Setup.md)" -ForegroundColor DarkGray
}

function Test-HuntEmailGraphConnection {
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
        Write-Host "[ERROR] Microsoft.Graph is not available. Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Red
        Write-HuntEmailGraphConnectionHelp
        return $false
    }

    try {
        $context = Get-MgContext -ErrorAction Stop
        if (-not $context) { throw "Not connected" }
        return $true
    }
    catch {
        Write-HuntEmailGraphConnectionHelp
        return $false
    }
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
    $Seen = @{}

    do {
        if ($Seen.ContainsKey($NextUri)) { break }
        $Seen[$NextUri] = $true

        $Response = Invoke-HuntEmailGraphRequest -Method GET -Uri $NextUri -Headers $Headers
        foreach ($Item in @($Response.value)) {
            if ($null -ne $Item) { $Items.Add($Item) }
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
    $encodedUser = [uri]::EscapeDataString($UserPrincipalName)
    $PendingUris.Enqueue("https://graph.microsoft.com/v1.0/users/$encodedUser/mailFolders?`$select=id,displayName,childFolderCount&`$top=100")

    while ($PendingUris.Count -gt 0) {
        $NextUri = $PendingUris.Dequeue()
        do {
            $Response = Invoke-HuntEmailGraphRequest -Method GET -Uri $NextUri
            foreach ($Folder in @($Response.value)) {
                if ($null -eq $Folder) { continue }
                $FolderMap[$Folder.id] = $Folder.displayName
                if ($Folder.childFolderCount -gt 0) {
                    $PendingUris.Enqueue("https://graph.microsoft.com/v1.0/users/$encodedUser/mailFolders/$([uri]::EscapeDataString([string]$Folder.id))/childFolders?`$select=id,displayName,childFolderCount&`$top=100")
                }
            }

            $NextUri = $Response.'@odata.nextLink'
        } while ($NextUri)
    }

    return $FolderMap
}

function ConvertTo-HuntEmailPlainText {
    param ([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return "" }

    $text = $Html -replace '(?is)<script.*?</script>', ' ' -replace '(?is)<style.*?</style>', ' '
    $text = $text -replace '(?s)<[^>]+>', ' '
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return (($text -replace '\s+', ' ').Trim())
}

function Get-HuntEmailSafeFileName {
    param ([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return "message" }

    $builder = [System.Text.StringBuilder]::new()
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $Name.ToCharArray()) {
        if ($invalid -contains $char -or [int]$char -lt 32) {
            [void]$builder.Append('_')
        }
        else {
            [void]$builder.Append($char)
        }
    }

    $clean = $builder.ToString().Trim()
    if (-not $clean) { $clean = "message" }
    if ($clean.Length -gt 80) { $clean = $clean.Substring(0, 80).Trim() }
    return $clean
}

function ConvertTo-HuntEmailUtcDateTime {
    param ([datetime]$Date)

    if ($Date.Kind -eq [DateTimeKind]::Unspecified) {
        return [DateTime]::SpecifyKind($Date, [DateTimeKind]::Local).ToUniversalTime()
    }
    return $Date.ToUniversalTime()
}

function ConvertTo-HuntEmailUtcGraphString {
    param ([datetime]$Date)
    return (ConvertTo-HuntEmailUtcDateTime -Date $Date).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
}

function Format-HuntEmailKqlToken {
    param ([string]$Value)

    $clean = ($Value -replace '"', ' ').Trim()
    if ($clean -match '[\s()]') { return '"' + $clean + '"' }
    return $clean
}

function Format-HuntEmailRecipient {
    param ($Recipients)

    $addresses = foreach ($recipient in @($Recipients)) {
        if ($null -eq $recipient) { continue }
        if ($recipient.emailAddress.address) { [string]$recipient.emailAddress.address }
    }
    return (($addresses | Where-Object { $_ }) -join '; ')
}

function Resolve-HuntEmailCopilotAppClass {
    param ([Parameter(Mandatory)][string]$Name)

    $trimmed = $Name.Trim()
    if ($script:HuntEmailCopilotAppClasses.ContainsKey($trimmed)) {
        return [string]$script:HuntEmailCopilotAppClasses[$trimmed]
    }

    foreach ($entry in $script:HuntEmailCopilotAppClasses.GetEnumerator()) {
        if ($trimmed -eq $entry.Value) { return [string]$entry.Value }
    }

    $supported = ($script:HuntEmailCopilotAppClasses.Keys | Sort-Object) -join ', '
    throw "Unsupported Copilot appClass '$trimmed'. Learn lists these interaction-history filters: $supported. Other Purview item classes (Outlook, PowerPoint, Loop, Whiteboard, and others) are not a documented `$filter value. Run huntCopilot without -AppClass and filter the AppClass property on the results."
}

function Write-HuntEmailCopilotForbidden {
    param (
        [string]$UserId,
        [int]$StatusCode
    )

    $codeText = if ($StatusCode) { "HTTP $StatusCode" } else { "access denied" }
    Write-Host "[ERROR] Microsoft Graph refused Copilot interaction history for $UserId ($codeText)." -ForegroundColor Red
    Write-Host "  Required application permission: $script:HuntEmailCopilotPermission (admin consent)." -ForegroundColor Yellow
    Write-Host "  Delegated permissions are not supported for getAllEnterpriseInteractions." -ForegroundColor Yellow
    Write-Host "  The target user also needs a Microsoft 365 Copilot license with the Microsoft Copilot with Graph-grounded chat service plan." -ForegroundColor Yellow
    Write-Host "  Learn: $script:HuntEmailCopilotLearn" -ForegroundColor DarkGray
    Write-Host "  Docs: $script:HuntEmailPlaybookDocs" -ForegroundColor DarkGray
}

function Get-HuntEmailCopilotInteraction {
    <#
    .SYNOPSIS
        GET /copilot/users/{id}/interactionHistory/getAllEnterpriseInteractions (Graph v1.0).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [datetime]$StartDate,

        [Parameter(Mandatory)]
        [datetime]$EndDate,

        [Parameter()]
        [string[]]$AppClass,

        [Parameter()]
        [ValidateRange(1, 100)]
        [int]$Top = 100,

        [Parameter()]
        [ValidateRange(1, 500)]
        [int]$PageLimit = 200
    )

    if ($StartDate -ge $EndDate) {
        throw "StartDate must be earlier than EndDate."
    }

    $encodedUser = [uri]::EscapeDataString($UserId)
    $classFilters = @($null)
    if ($AppClass) {
        $classFilters = foreach ($name in $AppClass) {
            Resolve-HuntEmailCopilotAppClass -Name $name
        }
    }

    $items = [System.Collections.Generic.List[object]]::new()
    $seenIds = @{}
    $truncated = $false

    foreach ($classItem in $classFilters) {
        # Learn requires both bounds and uses exclusive gt/lt. Widen by one second so the caller's range is inclusive.
        $filterStart = ConvertTo-HuntEmailUtcGraphString -Date (ConvertTo-HuntEmailUtcDateTime $StartDate).AddSeconds(-1)
        $filterEnd = ConvertTo-HuntEmailUtcGraphString -Date (ConvertTo-HuntEmailUtcDateTime $EndDate).AddSeconds(1)
        $filter = "createdDateTime gt $filterStart and createdDateTime lt $filterEnd"
        if ($classItem) {
            $safeClass = ([string]$classItem).Replace("'", "''")
            $filter = "appClass eq '$safeClass' and $filter"
        }

        $encodedFilter = [uri]::EscapeDataString($filter)
        $nextUri = "https://graph.microsoft.com/v1.0/copilot/users/$encodedUser/interactionHistory/getAllEnterpriseInteractions?`$top=$Top&`$filter=$encodedFilter"
        $seenUris = @{}
        $pages = 0

        while ($nextUri) {
            if ($seenUris.ContainsKey($nextUri)) { break }
            $seenUris[$nextUri] = $true
            $pages++
            if ($pages -gt $PageLimit) {
                $truncated = $true
                break
            }

            try {
                $response = Invoke-HuntEmailGraphRequest -Method GET -Uri $nextUri
            }
            catch {
                $status = Get-HuntEmailGraphStatusCode -ErrorRecord $_
                if ($status -in 401, 403 -or $_.Exception.Message -match 'Authorization_RequestDenied|Forbidden|ErrorAccessDenied') {
                    Write-HuntEmailCopilotForbidden -UserId $UserId -StatusCode $status
                    throw "HUNTEMAIL_COPILOT_FORBIDDEN: $script:HuntEmailCopilotPermission is not consented, or the Copilot license service plan is missing for $UserId."
                }
                if ($status -eq 400 -and $classItem) {
                    throw "Graph rejected the Copilot appClass filter '$classItem' (HTTP 400). Supported interaction-history filters: $(($script:HuntEmailCopilotAppClasses.Keys | Sort-Object) -join ', ')."
                }
                throw
            }

            foreach ($item in @($response.value)) {
                if ($null -eq $item) { continue }
                $key = [string]$item.id
                if (-not $key) { $key = [guid]::NewGuid().ToString() }
                if ($seenIds.ContainsKey($key)) { continue }
                $seenIds[$key] = $true
                $items.Add($item)
            }

            $nextUri = $response.'@odata.nextLink'
            if (-not $nextUri -and $pages -ge $PageLimit) { $truncated = $true }
        }

        if ($truncated) { break }
    }

    return [PSCustomObject]@{
        UserId        = $UserId
        Interactions  = $items
        Truncated     = $truncated
    }
}

function ConvertTo-HuntEmailCopilotRow {
    param (
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        $Item
    )

    $contentType = [string]$Item.body.contentType
    $body = [string]$Item.body.content
    $bodyText = if ($contentType -match 'html') { ConvertTo-HuntEmailPlainText -Html $body } else { $body }

    $from = $null
    if ($Item.from.user.displayName) { $from = [string]$Item.from.user.displayName }
    elseif ($Item.from.application.displayName) { $from = [string]$Item.from.application.displayName }
    elseif ($Item.from.application.id) { $from = [string]$Item.from.application.id }

    $contexts = foreach ($context in @($Item.contexts)) {
        if ($null -eq $context) { continue }
        $bits = @($context.contextType, $context.displayName, $context.contextReference) | Where-Object { $_ }
        if ($bits) { $bits -join ': ' }
    }

    $links = foreach ($link in @($Item.links)) {
        if ($null -eq $link) { continue }
        if ($link.displayName -and $link.linkUrl) { "$($link.displayName) <$($link.linkUrl)>" }
        elseif ($link.linkUrl) { [string]$link.linkUrl }
    }

    $attachmentNames = foreach ($attachment in @($Item.attachments)) {
        if ($attachment.name) { [string]$attachment.name }
    }

    $appClass = [string]$Item.appClass
    $appName = if ($appClass) { ($appClass -split '\.')[-1] } else { "" }

    [PSCustomObject]@{
        UserId            = $UserId
        CreatedDateTime   = [string]$Item.createdDateTime
        AppClass          = $appClass
        AppName           = $appName
        InteractionType   = [string]$Item.interactionType
        ConversationType  = [string]$Item.conversationType
        SessionId         = [string]$Item.sessionId
        RequestId         = [string]$Item.requestId
        InteractionId     = [string]$Item.id
        Locale            = [string]$Item.locale
        From              = $from
        BodyContentType   = $contentType
        Body              = $bodyText
        Contexts          = (($contexts | Where-Object { $_ }) -join ' | ')
        Links             = (($links | Where-Object { $_ }) -join ' | ')
        AttachmentNames   = (($attachmentNames | Where-Object { $_ }) -join ' | ')
    }
}

function Test-HuntEmailExchangeOnlineConnection {
    if ($script:HuntEmailExchangeConnectionTest) {
        return [bool](& $script:HuntEmailExchangeConnectionTest)
    }

    if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
        Write-Host "[ERROR] Search-UnifiedAuditLog is not available." -ForegroundColor Red
        Write-Host "  auditCopilot uses Exchange Online PowerShell, which is a separate sign-in from Connect-MgGraph." -ForegroundColor Yellow
        Write-Host "  Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
        Write-Host "  Connect-ExchangeOnline" -ForegroundColor Yellow
        Write-Host "  Then retry auditCopilot. The account needs the Exchange Online Audit Logs or View-Only Audit Logs role" -ForegroundColor Yellow
        Write-Host "  (included by default in Compliance Management and Organization Management)." -ForegroundColor Yellow
        Write-Error "Search-UnifiedAuditLog is not available. Install ExchangeOnlineManagement and run Connect-ExchangeOnline before auditCopilot."
        return $false
    }

    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        $connections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
        if ($connections.Count -eq 0) {
            Write-Host "[ERROR] Not connected to Exchange Online. Run Connect-ExchangeOnline, then retry auditCopilot." -ForegroundColor Red
            Write-Host "  Graph certificate auth does not create this session." -ForegroundColor Yellow
            Write-Error "Not connected to Exchange Online. Run Connect-ExchangeOnline, then retry auditCopilot."
            return $false
        }
    }

    return $true
}

function Invoke-HuntEmailAuditSearch {
    param ([Parameter(Mandatory)][hashtable]$Parameters)

    if ($script:HuntEmailSearchAuditLog) {
        return & $script:HuntEmailSearchAuditLog $Parameters
    }

    return Search-UnifiedAuditLog @Parameters
}

function ConvertTo-HuntEmailAuditRow {
    param ($Record)

    $audit = $Record.AuditData
    if ($audit -is [string] -and $audit) {
        $audit = $audit | ConvertFrom-Json
    }

    $eventData = $null
    if ($audit) { $eventData = $audit.CopilotEventData }

    $resources = @($eventData.AccessedResources)
    $messages = @($eventData.Messages)
    $contexts = @($eventData.Contexts)
    $plugins = @($eventData.AISystemPlugin)
    $models = @($eventData.ModelTransparencyDetails)

    $promptIds = foreach ($message in $messages) {
        if ($null -eq $message) { continue }
        if ($message.isPrompt -eq $true -and $message.Id) { [string]$message.Id }
    }
    $responseIds = foreach ($message in $messages) {
        if ($null -eq $message) { continue }
        if ($message.isPrompt -eq $false -and $message.Id) { [string]$message.Id }
    }

    $userId = $null
    if ($audit -and $audit.UserId) { $userId = [string]$audit.UserId }
    elseif ($Record.UserIds) { $userId = (@($Record.UserIds) -join '; ') }

    $creation = $null
    if ($audit -and $audit.CreationTime) { $creation = [string]$audit.CreationTime }
    elseif ($Record.CreationDate) { $creation = [string]$Record.CreationDate }

    [PSCustomObject]@{
        UserId                   = $userId
        CreationTime             = $creation
        Operation                = if ($audit.Operation) { [string]$audit.Operation } else { [string]$Record.Operations }
        RecordType               = if ($audit.RecordType) { [string]$audit.RecordType } else { [string]$Record.RecordType }
        Workload                 = if ($audit) { [string]$audit.Workload } else { "" }
        ClientRegion             = if ($audit -and $audit.ClientRegion) { [string]$audit.ClientRegion } else { "" }
        AppHost                  = if ($eventData) { [string]$eventData.AppHost } else { "" }
        AppIdentity              = if ($eventData -and $eventData.AppIdentity) { [string]$eventData.AppIdentity } else { "" }
        ThreadId                 = if ($eventData) { [string]$eventData.ThreadId } else { "" }
        PromptMessageIds         = (($promptIds | Where-Object { $_ }) -join ' | ')
        ResponseMessageIds       = (($responseIds | Where-Object { $_ }) -join ' | ')
        AccessedResources        = (($resources | Where-Object { $_.Name } | ForEach-Object { [string]$_.Name }) -join ' | ')
        AccessedResourceUrls     = (($resources | Where-Object { $_.SiteUrl } | ForEach-Object { [string]$_.SiteUrl }) -join ' | ')
        AccessedResourceTypes    = (($resources | Where-Object { $_.Type } | ForEach-Object { [string]$_.Type }) -join ' | ')
        AccessedResourceActions  = (($resources | Where-Object { $_.Action } | ForEach-Object { [string]$_.Action }) -join ' | ')
        SensitivityLabelIds      = (($resources | Where-Object { $_.SensitivityLabelId } | ForEach-Object { [string]$_.SensitivityLabelId } | Select-Object -Unique) -join ' | ')
        ContextIds               = (($contexts | Where-Object { $_.Id } | ForEach-Object { [string]$_.Id }) -join ' | ')
        ContextTypes             = (($contexts | Where-Object { $_.Type } | ForEach-Object { [string]$_.Type }) -join ' | ')
        AISystemPlugins          = (($plugins | Where-Object { $_.Id -or $_.Name } | ForEach-Object { @($_.Id, $_.Name) -join ':' }) -join ' | ')
        ModelProviders           = (($models | Where-Object { $_.ModelProviderName } | ForEach-Object { [string]$_.ModelProviderName }) -join ' | ')
        ModelNames               = (($models | Where-Object { $_.ModelName } | ForEach-Object { [string]$_.ModelName }) -join ' | ')
    }
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
        huntEmail -Sender "attacker@evil.example"

    .EXAMPLE
        huntEmail -Sender "phish@bad.example" -StartDate "2026-02-15" -EndDate "2026-02-18"

    .EXAMPLE
        hunt -InternetMessageId "<ABC123@mail.example>"

    .EXAMPLE
        hunt -Subject "Invoice" -Detailed -ExportEml "C:\IR\case001"

    .EXAMPLE
        hunt -Keywords "password reset" -Sender "noreply@suspicious.example" | Export-Csv results.csv

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
        $abort = $false
        if (-not (Test-HuntEmailGraphConnection)) {
            $abort = $true
            return
        }

        if (-not $Sender -and -not $Subject -and -not $Keywords -and -not $AttachmentName -and -not $InternetMessageId) {
            Write-Host "[ERROR] Provide at least one search criteria: -Sender, -Subject, -Keywords, -AttachmentName, or -InternetMessageId" -ForegroundColor Red
            $abort = $true
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
        $select = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,parentFolderId,internetMessageId,bodyPreview,conversationId'
        $explicitStart = $PSBoundParameters.ContainsKey('StartDate')
        $explicitEnd = $PSBoundParameters.ContainsKey('EndDate')
    }

    process {
        if ($abort) { return }

        if ($UserPrincipalName) {
            $Mailboxes = @($UserPrincipalName)
            Write-Host "[SEARCH] Targeting $($Mailboxes.Count) specified mailbox(es)" -ForegroundColor Cyan
        }
        else {
            Write-Host "[SEARCH] Enumerating all licensed mailboxes..." -ForegroundColor Cyan
            try {
                $userUri = "https://graph.microsoft.com/v1.0/users?`$filter=assignedLicenses/`$count ne 0 and accountEnabled eq true&`$count=true&`$select=userPrincipalName,displayName,mail&`$top=999"
                $users = Get-HuntEmailPagedCollection -Uri $userUri -Headers @{ ConsistencyLevel = "eventual" }
                $Mailboxes = @($users | Where-Object { $_.mail } | ForEach-Object { [string]$_.userPrincipalName })
                Write-Host "[SEARCH] Found $($Mailboxes.Count) mailboxes to search" -ForegroundColor Cyan
            }
            catch {
                Write-Host "[ERROR] Failed to enumerate mailboxes: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "  User.Read.All (Application) is required for a tenant-wide hunt. Pass -UserPrincipalName to search specific mailboxes." -ForegroundColor Yellow
                $abort = $true
                return
            }
        }

        if ($Mailboxes.Count -eq 0) {
            Write-Host "[SEARCH] No mailboxes to search." -ForegroundColor Yellow
            return
        }

        $current = 0
        foreach ($mailbox in $Mailboxes) {
            $current++
            $pct = [math]::Round(($current / $Mailboxes.Count) * 100)
            if ($Nyan) {
                Write-NyanProgress -PercentComplete $pct -Status "[$current/$($Mailboxes.Count)] $mailbox" -Activity "Hunt"
            }
            else {
                Write-Progress -Activity "Hunting mailboxes" -Status "[$current/$($Mailboxes.Count)] $mailbox" -PercentComplete $pct
            }

            $encodedUser = [uri]::EscapeDataString($mailbox)
            $folderMap = @{}
            try {
                $folderMap = Get-HuntEmailFolderMap -UserPrincipalName $mailbox
            }
            catch {
                Write-Verbose "Folder map failed for ${mailbox}: $($_.Exception.Message)"
                $folderMap = @{}
            }

            $mode = 'search'
            $headers = $null
            $baseUri = $null
            if ($InternetMessageId) {
                $mode = 'internetMessageId'
                $escapedId = $InternetMessageId.Replace("'", "''")
                $filterParts = @("internetMessageId eq '$escapedId'")
                if ($explicitStart) {
                    $filterParts += "receivedDateTime ge $(ConvertTo-HuntEmailUtcGraphString $StartDate)"
                }
                if ($explicitEnd) {
                    $filterParts += "receivedDateTime le $(ConvertTo-HuntEmailUtcGraphString $EndDate)"
                }
                $filter = [uri]::EscapeDataString(($filterParts -join ' and '))
                $baseUri = "https://graph.microsoft.com/v1.0/users/$encodedUser/messages?`$filter=$filter&`$select=$select&`$top=50"
            }
            else {
                $kql = [System.Collections.Generic.List[string]]::new()
                if ($Sender) { $kql.Add("from:$(Format-HuntEmailKqlToken $Sender)") }
                if ($Subject) { $kql.Add("subject:$(Format-HuntEmailKqlToken $Subject)") }
                if ($Keywords) { $kql.Add("body:$(Format-HuntEmailKqlToken $Keywords)") }
                if ($AttachmentName) { $kql.Add("attachment:$(Format-HuntEmailKqlToken $AttachmentName)") }
                $kql.Add(("received>={0}" -f (ConvertTo-HuntEmailUtcDateTime $StartDate).ToString('yyyy-MM-dd')))
                $kql.Add(("received<={0}" -f (ConvertTo-HuntEmailUtcDateTime $EndDate).ToString('yyyy-MM-dd')))
                $searchValue = [uri]::EscapeDataString('"' + ($kql -join ' AND ') + '"')
                $headers = @{ ConsistencyLevel = "eventual" }
                $baseUri = "https://graph.microsoft.com/v1.0/users/$encodedUser/messages?`$search=$searchValue&`$select=$select&`$top=50"
            }

            $matched = 0
            $nextUri = $baseUri
            $seenUris = @{}
            $fellBack = $false

            while ($nextUri) {
                if ($seenUris.ContainsKey($nextUri)) { break }
                $seenUris[$nextUri] = $true

                try {
                    $response = Invoke-HuntEmailGraphRequest -Method GET -Uri $nextUri -Headers $headers
                }
                catch {
                    $status = Get-HuntEmailGraphStatusCode -ErrorRecord $_
                    $canFallback = ($mode -eq 'search' -and -not $fellBack -and $matched -eq 0 -and ($status -eq 400 -or $_.Exception.Message -match 'Search|Parsing|syntax|BadRequest'))
                    if ($canFallback) {
                        Write-Host "[SEARCH] `$search was rejected for $mailbox. Falling back to a receivedDateTime filter." -ForegroundColor Yellow
                        $fellBack = $true
                        $mode = 'date'
                        $headers = $null
                        $dateFilter = [uri]::EscapeDataString("receivedDateTime ge $(ConvertTo-HuntEmailUtcGraphString $StartDate) and receivedDateTime le $(ConvertTo-HuntEmailUtcGraphString $EndDate)")
                        $nextUri = "https://graph.microsoft.com/v1.0/users/$encodedUser/messages?`$filter=$dateFilter&`$select=$select&`$top=50"
                        $seenUris = @{}
                        continue
                    }

                    if ($status -in 401, 403) {
                        Write-Host "[ERROR] Graph denied mailbox read for $mailbox (HTTP $status). Application permissions Mail.Read and User.Read.All are required." -ForegroundColor Red
                    }
                    else {
                        Write-Host "[ERROR] ${mailbox}: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    break
                }

                foreach ($message in @($response.value)) {
                    if ($null -eq $message) { continue }

                    $receivedRaw = $message.receivedDateTime
                    $received = $null
                    if ($receivedRaw) {
                        $received = ConvertTo-HuntEmailUtcDateTime ([datetime]$receivedRaw)
                    }

                    $applyDate = ($mode -ne 'internetMessageId') -or $explicitStart -or $explicitEnd
                    if ($applyDate -and $received) {
                        if ($received -lt (ConvertTo-HuntEmailUtcDateTime $StartDate)) { continue }
                        if ($received -gt (ConvertTo-HuntEmailUtcDateTime $EndDate)) { continue }
                    }

                    $fromAddress = [string]$message.from.emailAddress.address
                    if (-not $fromAddress) { $fromAddress = [string]$message.from.emailAddress.name }
                    if ($Sender) {
                        $senderPattern = "*$([WildcardPattern]::Escape($Sender))*"
                        if ($fromAddress -notlike $senderPattern) { continue }
                    }
                    if ($Subject) {
                        $subjectPattern = "*$([WildcardPattern]::Escape($Subject))*"
                        if (([string]$message.subject) -notlike $subjectPattern) { continue }
                    }

                    $attachmentNames = @()
                    $fetchedAttachments = $false
                    if ($message.hasAttachments) {
                        try {
                            $attachmentUri = "https://graph.microsoft.com/v1.0/users/$encodedUser/messages/$([uri]::EscapeDataString([string]$message.id))/attachments?`$select=id,name,contentType,size,isInline"
                            $attachmentResponse = Invoke-HuntEmailGraphRequest -Method GET -Uri $attachmentUri
                            $fetchedAttachments = $true
                            $attachmentNames = @(
                                $attachmentResponse.value |
                                    Where-Object { $_ -and -not $_.isInline -and $_.name } |
                                    ForEach-Object { [string]$_.name }
                            )
                        }
                        catch {
                            Write-Verbose "Attachment lookup failed for $($message.id): $($_.Exception.Message)"
                        }
                    }

                    if ($AttachmentName) {
                        $attachmentPattern = [regex]::Escape($AttachmentName)
                        $attachmentHit = $false
                        foreach ($name in $attachmentNames) {
                            if ($name -match $attachmentPattern) { $attachmentHit = $true; break }
                        }
                        if (-not $attachmentHit) { continue }
                    }

                    $bodyText = [string]$message.bodyPreview
                    if ($Detailed) {
                        try {
                            $bodyUri = "https://graph.microsoft.com/v1.0/users/$encodedUser/messages/$([uri]::EscapeDataString([string]$message.id))?`$select=body"
                            $bodyResponse = Invoke-HuntEmailGraphRequest -Method GET -Uri $bodyUri
                            $rawBody = [string]$bodyResponse.body.content
                            if ([string]$bodyResponse.body.contentType -match 'html') {
                                $bodyText = ConvertTo-HuntEmailPlainText -Html $rawBody
                            }
                            elseif ($rawBody) {
                                $bodyText = $rawBody
                            }
                        }
                        catch {
                            Write-Verbose "Body lookup failed for $($message.id): $($_.Exception.Message)"
                        }
                    }

                    if ($mode -eq 'date' -and $Keywords) {
                        $keywordPattern = "*$([WildcardPattern]::Escape($Keywords))*"
                        $haystack = "$($message.subject) $($message.bodyPreview) $bodyText"
                        if ($haystack -notlike $keywordPattern) { continue }
                    }

                    $folderName = "Unknown"
                    if ($message.parentFolderId -and $folderMap.ContainsKey([string]$message.parentFolderId)) {
                        $folderName = [string]$folderMap[[string]$message.parentFolderId]
                    }

                    $emlPath = $null
                    if ($ExportEml) {
                        $fileBase = Get-HuntEmailSafeFileName -Name ([string]$message.subject)
                        $emlPath = Join-Path $ExportEml "$fileBase.eml"
                        $suffix = 1
                        while (Test-Path -LiteralPath $emlPath) {
                            $emlPath = Join-Path $ExportEml "$fileBase-$suffix.eml"
                            $suffix++
                        }
                        $mimeUri = "https://graph.microsoft.com/v1.0/users/$encodedUser/messages/$([uri]::EscapeDataString([string]$message.id))/`$value"
                        Invoke-HuntEmailGraphRequest -Method GET -Uri $mimeUri -OutputFilePath $emlPath | Out-Null
                    }

                    $hasAttach = if ($fetchedAttachments) { [bool]($attachmentNames.Count -gt 0) } else { [bool]$message.hasAttachments }

                    $ResultsList.Add([PSCustomObject]@{
                        Mailbox           = $mailbox
                        MessageId         = [string]$message.id
                        InternetMessageId = [string]$message.internetMessageId
                        Subject           = [string]$message.subject
                        From              = $fromAddress
                        To                = Format-HuntEmailRecipient $message.toRecipients
                        Cc                = Format-HuntEmailRecipient $message.ccRecipients
                        Received          = $received
                        Sent              = if ($message.sentDateTime) { ConvertTo-HuntEmailUtcDateTime ([datetime]$message.sentDateTime) } else { $null }
                        Folder            = $folderName
                        FolderId          = [string]$message.parentFolderId
                        IsRead            = [bool]$message.isRead
                        HasAttach         = $hasAttach
                        AttachmentNames   = ($attachmentNames -join '; ')
                        BodyPreview       = [string]$message.bodyPreview
                        Body              = $bodyText
                        EmlPath           = $emlPath
                    })

                    $matched++
                    if ($MaxResults -gt 0 -and $matched -ge $MaxResults) { break }
                }

                if ($MaxResults -gt 0 -and $matched -ge $MaxResults) { break }
                $nextUri = $response.'@odata.nextLink'
            }
        }

        if ($Nyan) {
            Write-NyanComplete -Activity "Hunt"
        }
        else {
            Write-Progress -Activity "Hunting mailboxes" -Completed
        }
    }

    end {
        if ($abort) { return }

        Write-Host "[SEARCH] $($ResultsList.Count) message(s) matched." -ForegroundColor Cyan
        if ($Detailed) {
            foreach ($row in $ResultsList) {
                Write-Host "----------------------------------------" -ForegroundColor DarkGray
                Write-Host "Mailbox: $($row.Mailbox)  Folder: $($row.Folder)  Read: $($row.IsRead)" -ForegroundColor White
                Write-Host "From: $($row.From)" -ForegroundColor White
                Write-Host "Subject: $($row.Subject)" -ForegroundColor White
                Write-Host "InternetMessageId: $($row.InternetMessageId)" -ForegroundColor DarkGray
                if ($row.AttachmentNames) { Write-Host "Attachments: $($row.AttachmentNames)" -ForegroundColor White }
                if ($row.Body) { Write-Host $row.Body -ForegroundColor Gray }
                if ($row.EmlPath) { Write-Host "EML: $($row.EmlPath)" -ForegroundColor DarkGray }
            }
        }

        return $ResultsList.ToArray()
    }
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

        -IncludeCopilot asks the Copilot interaction history API whether each
        recipient mailbox used Microsoft 365 Copilot during the delivery window.
        It is read-only and requires AiEnterpriseInteraction.Read.All.

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

    .PARAMETER IncludeCopilot
        For each recipient mailbox, record whether Copilot interaction history
        has any item in the timeline window.

    .EXAMPLE
        timelineEmail -Sender "attacker@evil.example" -Subject "Urgent Invoice"

    .EXAMPLE
        huntEmail -Sender "phish@bad.example" | timelineEmail

    .EXAMPLE
        timeline -Sender "noreply@suspicious.example" -StartDate "2026-02-15" | Export-Csv timeline.csv

    .EXAMPLE
        timelineEmail -Sender "phish@bad.example" -IncludeCopilot

    .NOTES
        Author: Security Operations
        Created: 2/19/2026
        Part of: Email Threat Response Toolkit
        Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048
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
        [switch]$IncludeCopilot,

        [Parameter()]
        [switch]$Nyan
    )

    begin {
        $abort = $false
        if (-not (Test-HuntEmailGraphConnection)) {
            $abort = $true
            return
        }

        $PipedResults = [System.Collections.Generic.List[PSObject]]::new()
        $Timeline = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        if ($abort) { return }
        if ($SearchResults) {
            foreach ($R in $SearchResults) { $PipedResults.Add($R) }
        }
    }

    end {
        if ($abort) { return }
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

            $AllResults = @(huntEmail @SearchParams)

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
        $Total = @($AllResults).Count
        $Current = 0

        foreach ($Result in @($AllResults)) {
            $Current++
            $Pct = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100) } else { 100 }
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

        $Sorted = @($Timeline | Sort-Object Timestamp)

        if ($IncludeCopilot -and $Sorted.Count -gt 0) {
            $windowStart = $StartDate
            $windowEnd = $EndDate
            $stamps = @($Sorted | Where-Object { $_.Timestamp } | ForEach-Object { [datetime]$_.Timestamp })
            if ($stamps.Count -gt 0) {
                $windowStart = ($stamps | Measure-Object -Minimum).Minimum
                $windowEnd = ($stamps | Measure-Object -Maximum).Maximum
                if ($windowStart -eq $windowEnd) {
                    $windowStart = $windowStart.AddMinutes(-5)
                    $windowEnd = $windowEnd.AddMinutes(5)
                }
            }

            $copilotByMailbox = @{}
            $copilotDenied = $false
            $mailboxes = @($Sorted | Select-Object -ExpandProperty Mailbox -Unique)
            $mbIndex = 0
            foreach ($mb in $mailboxes) {
                $mbIndex++
                if ($Nyan) {
                    Write-NyanProgress -PercentComplete ([math]::Round(($mbIndex / $mailboxes.Count) * 100)) -Status $mb -Activity "Copilot"
                }
                if ($copilotDenied) { continue }
                try {
                    $copilotByMailbox[$mb] = Get-HuntEmailCopilotInteraction -UserId $mb -StartDate $windowStart -EndDate $windowEnd -Top 25 -PageLimit 1
                }
                catch {
                    if ($_.Exception.Message -like '*HUNTEMAIL_COPILOT_FORBIDDEN*') {
                        $copilotDenied = $true
                    }
                    else {
                        Write-Host "[WARN] Copilot lookup failed for ${mb}: $($_.Exception.Message)" -ForegroundColor Yellow
                        $copilotByMailbox[$mb] = $null
                    }
                }
            }
            if ($Nyan) { Write-NyanComplete -Activity "Copilot" }

            foreach ($row in $Sorted) {
                if ($copilotDenied) {
                    $row | Add-Member -NotePropertyName CopilotUsed -NotePropertyValue $null
                    $row | Add-Member -NotePropertyName CopilotInteractionCount -NotePropertyValue $null
                    $row | Add-Member -NotePropertyName CopilotAppClasses -NotePropertyValue $null
                    $row | Add-Member -NotePropertyName CopilotNote -NotePropertyValue "AiEnterpriseInteraction.Read.All was not consented"
                    continue
                }

                $info = $copilotByMailbox[$row.Mailbox]
                $count = 0
                $apps = ""
                $more = $false
                if ($info) {
                    $count = @($info.Interactions).Count
                    $apps = ((@($info.Interactions) | ForEach-Object { [string]$_.appClass } | Where-Object { $_ } | Select-Object -Unique) -join '; ')
                    $more = [bool]$info.Truncated
                }
                $row | Add-Member -NotePropertyName CopilotUsed -NotePropertyValue ($count -gt 0)
                $row | Add-Member -NotePropertyName CopilotInteractionCount -NotePropertyValue $count
                $row | Add-Member -NotePropertyName CopilotAppClasses -NotePropertyValue $apps
                $row | Add-Member -NotePropertyName CopilotMayHaveMore -NotePropertyValue $more
            }
        }

        $ReadCount = @($Sorted | Where-Object { $_.IsRead }).Count
        $UnreadCount = @($Sorted | Where-Object { -not $_.IsRead }).Count
        $UniqueMailboxes = @($Sorted | Select-Object -ExpandProperty Mailbox -Unique).Count

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
            if ($First -and $Last) {
                Write-Host "  Spread:      $(($Last - $First).ToString())" -ForegroundColor White
            }
        }

        Write-Host "  Total hits:  $($Sorted.Count)" -ForegroundColor White
        Write-Host "  Mailboxes:   $UniqueMailboxes" -ForegroundColor White
        Write-Host "  Read:        $ReadCount" -ForegroundColor $(if ($ReadCount -gt 0) { "Red" } else { "Green" })
        Write-Host "  Unread:      $UnreadCount" -ForegroundColor $(if ($UnreadCount -gt 0) { "Yellow" } else { "Green" })
        if ($IncludeCopilot) {
            $copilotUsers = @($Sorted | Where-Object { $_.CopilotUsed } | Select-Object -ExpandProperty Mailbox -Unique).Count
            Write-Host "  Copilot:     $copilotUsers mailbox(es) with interaction history in this window" -ForegroundColor White
        }
        Write-Host ""

        $Sorted | Format-Table Timestamp, Mailbox, Folder, Actions, HasAttach -AutoSize | Out-Host

        if ($ReadCount -gt 0) {
            Write-Host "  WARNING: $ReadCount user(s) opened this email." -ForegroundColor Red
            $ReadUsers = ($Sorted | Where-Object { $_.IsRead } | Select-Object -ExpandProperty Mailbox -Unique) -join ", "
            Write-Host "  Users who read it: $ReadUsers" -ForegroundColor Red
        }

        Write-Host "`n  Tip: Pipe to purgeEmail to purge all matched messages" -ForegroundColor DarkGray
        Write-Host "  Copilot content is not purged by this toolkit. Use Purview eDiscovery." -ForegroundColor DarkGray

        return $Sorted
    }
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

        This function deletes ordinary mailbox messages only. It does not delete
        Copilot interaction items. Copilot removal is a Purview eDiscovery task.

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
        huntEmail -Sender "attacker@evil.example" | purgeEmail

    .EXAMPLE
        huntEmail -Sender "phish@bad.example" | timelineEmail | purgeEmail

    .EXAMPLE
        purge -UserPrincipalName "user@example.com" -Sender "attacker@evil.example" -Force

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
        $abort = $false
        if (-not (Test-HuntEmailGraphConnection)) {
            $abort = $true
            return
        }

        $PipedResults = [System.Collections.Generic.List[PSObject]]::new()
        $PurgeTargets = [System.Collections.Generic.List[PSObject]]::new()
        $PurgeLog = [System.Collections.Generic.List[PSObject]]::new()
    }

    process {
        if ($abort) { return }
        if ($SearchResults) {
            foreach ($R in $SearchResults) { $PipedResults.Add($R) }
        }
    }

    end {
        if ($abort) { return }

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

            $Results = @(huntEmail @SearchParams)

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
        $UniqueMailboxes = @($PurgeTargets | Select-Object -ExpandProperty Mailbox -Unique).Count

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
                Invoke-HuntEmailGraphRequest -Method POST -Uri $DeleteUri

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

        Write-Verbose "Skipped: $SkipCount"

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

function Connect-HuntEmailGraph {
    <#
    .SYNOPSIS
        Connect to Microsoft Graph for the email investigation toolkit using certificate auth.

    .DESCRIPTION
        Uses a Microsoft Entra app registration and certificate thumbprint supplied
        via parameters or environment variables. Run this once before using
        huntEmail, timelineEmail, purgeEmail, or huntCopilot.

        If the cert is not yet in your Current User store, use -PfxPath and -Password
        to import it first (for example from C:\certs\huntEmail.pfx).

        Required values can be provided with parameters or these environment variables:
        HUNTEMAIL_TENANT_ID, HUNTEMAIL_CLIENT_ID, HUNTEMAIL_CERT_THUMBPRINT.

    .PARAMETER PfxPath
        Optional. Full path to the .pfx file. If provided and the thumbprint cert is not
        in Cert:\CurrentUser\My, the function imports the PFX then connects.

    .PARAMETER Password
        Optional. SecureString password for the PFX. Only used when -PfxPath is provided
        and the cert needs to be imported.

    .EXAMPLE
        Connect-HuntEmailGraph

    .EXAMPLE
        $pwd = Read-Host "PFX password" -AsSecureString
        Connect-HuntEmailGraph -TenantId "<tenant-id>" -ClientId "<client-id>" -Thumbprint "<thumbprint>" -PfxPath "C:\certs\huntEmail.pfx" -Password $pwd

    .NOTES
        See documentation\api_info.md for the configuration template.
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$TenantId = $env:HUNTEMAIL_TENANT_ID,

        [Parameter()]
        [string]$ClientId = $env:HUNTEMAIL_CLIENT_ID,

        [Parameter()]
        [string]$Thumbprint = $env:HUNTEMAIL_CERT_THUMBPRINT,

        [Parameter()]
        [string]$PfxPath,

        [Parameter()]
        [SecureString]$Password
    )

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Connect-MgGraph is not available. Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    if (-not $TenantId -or -not $ClientId -or -not $Thumbprint) {
        Write-Host "[ERROR] Missing Graph connection settings." -ForegroundColor Red
        Write-Host "  Provide -TenantId, -ClientId, and -Thumbprint" -ForegroundColor Yellow
        Write-Host "  or set HUNTEMAIL_TENANT_ID, HUNTEMAIL_CLIENT_ID, and HUNTEMAIL_CERT_THUMBPRINT." -ForegroundColor Yellow
        Write-Host "  See documentation\api_info.md for the configuration template." -ForegroundColor Gray
        throw "Missing Graph connection settings."
    }

    $certInStore = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumbprint }

    if (-not $certInStore -and $PfxPath -and $Password) {
        if (-not (Test-Path -LiteralPath $PfxPath)) {
            Write-Host "[ERROR] PFX not found: $PfxPath" -ForegroundColor Red
            throw "PFX not found: $PfxPath"
        }
        Write-Host "[CERT] Importing PFX to Current User\My..." -ForegroundColor Cyan
        try {
            Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $Password | Out-Null
            Write-Host "[CERT] Imported successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Failed to import PFX: $_" -ForegroundColor Red
            throw
        }
    }
    elseif (-not $certInStore -and $PfxPath) {
        Write-Host "[ERROR] Cert with thumbprint $Thumbprint not in store. Provide -Password to import from: $PfxPath" -ForegroundColor Red
        throw "Certificate $Thumbprint is not in the store and no PFX password was provided."
    }
    elseif (-not $certInStore) {
        $defaultPfx = "C:\certs\huntEmail.pfx"
        Write-Host "[ERROR] Cert with thumbprint $Thumbprint not in Cert:\CurrentUser\My." -ForegroundColor Red
        Write-Host "  Import first, e.g.:" -ForegroundColor Yellow
        Write-Host "    `$pwd = Read-Host 'PFX password' -AsSecureString" -ForegroundColor Gray
        Write-Host "    Import-HuntEmailCert -PfxPath '$defaultPfx' -Password `$pwd" -ForegroundColor Gray
        Write-Host "  Or: Connect-HuntEmailGraph -PfxPath '$defaultPfx' -Password `$pwd" -ForegroundColor Gray
        throw "Certificate $Thumbprint is not in Cert:\CurrentUser\My."
    }

    Write-Host "[GRAPH] Connecting to Microsoft Graph (certificate auth)..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        Write-Host "[GRAPH] Connected. ClientId: $($ctx.ClientId), AuthType: $($ctx.AuthType)" -ForegroundColor Green
        return $ctx
    }
    catch {
        Write-Host "[ERROR] Connect-MgGraph failed: $_" -ForegroundColor Red
        throw
    }
}

function Import-HuntEmailCert {
    <#
    .SYNOPSIS
        One-time import of the toolkit PFX into Current User\My.

    .DESCRIPTION
        Run this once per machine (or after cert rotation) to import the .pfx so that
        Connect-HuntEmailGraph can connect without prompting for a password each time.

        PFX location: C:\certs\huntEmail.pfx (from cert\createCert.ps1), or pass -PfxPath.

    .PARAMETER PfxPath
        Path to the .pfx file. Default: C:\certs\huntEmail.pfx

    .PARAMETER Password
        SecureString password for the PFX (for example, from your secret manager).

    .EXAMPLE
        $pwd = Read-Host "PFX password" -AsSecureString
        Import-HuntEmailCert -Password $pwd

    .EXAMPLE
        Import-HuntEmailCert -PfxPath "C:\certs\huntEmail.pfx" -Password (Read-Host "Password" -AsSecureString)
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$PfxPath = "C:\certs\huntEmail.pfx",

        [Parameter(Mandatory)]
        [SecureString]$Password
    )

    if (-not (Test-Path -LiteralPath $PfxPath)) {
        Write-Host "[ERROR] PFX not found: $PfxPath" -ForegroundColor Red
        Write-Host "  Create it with cert\createCert.ps1, or copy it from your secure location." -ForegroundColor Gray
        throw "PFX not found: $PfxPath"
    }

    Write-Host "[CERT] Importing to Cert:\CurrentUser\My..." -ForegroundColor Cyan
    try {
        $imported = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $Password -ErrorAction Stop
        Write-Host "[CERT] Imported. Thumbprint: $($imported.Thumbprint)" -ForegroundColor Green
        Write-Host "  Next: Connect-HuntEmailGraph" -ForegroundColor Gray
        return $imported
    }
    catch {
        Write-Host "[ERROR] Import failed: $_" -ForegroundColor Red
        throw
    }
}

function huntCopilot {
    <#
    .SYNOPSIS
        Read Microsoft 365 Copilot prompts and responses for one or more users.

    .DESCRIPTION
        Calls Microsoft Graph v1.0:
        GET /copilot/users/{id}/interactionHistory/getAllEnterpriseInteractions

        Standard /users/{id}/messages does not return these hidden Copilot items.
        The audit log (auditCopilot) has metadata and no prompt text. Pipe auditCopilot
        into huntCopilot to open the interaction content for the same users.

        Application permission: AiEnterpriseInteraction.Read.All (admin consent).
        Delegated permissions are not supported. The target user needs a Microsoft 365
        Copilot license with the Microsoft Copilot with Graph-grounded chat service plan.

        There is no Copilot purge function. Content removal is Purview eDiscovery only.

    .PARAMETER UserIds
        One or more user ids or UPNs. Also binds from pipeline properties UserId,
        UserPrincipalName, and Mailbox, including auditCopilot output.

    .PARAMETER StartDate
        Start of the createdDateTime window. Default: 7 days ago.
        When audit rows are piped and neither date is passed, the window is the
        min/max CreationTime of those rows.

    .PARAMETER EndDate
        End of the createdDateTime window. Default: now.
        Learn's filter is exclusive (gt/lt). This function widens each bound by one second.

    .PARAMETER AppClass
        Optional interaction-history app class filter. Supported values from Learn:
        Word, Excel, Teams, BizChat, WebChat, CoworkChat.
        Full IPM.SkypeTeams.Message.Copilot.<Name> values for those six are also accepted.

    .PARAMETER Top
        Page size. Learn recommends 100.

    .EXAMPLE
        huntCopilot -UserIds "user@example.com" -StartDate "2026-02-01" -EndDate "2026-02-18"

    .EXAMPLE
        auditCopilot -UserIds "user@example.com" -StartDate "2026-02-01" -EndDate "2026-02-18" | huntCopilot

    .EXAMPLE
        hcop -UserIds "user@example.com" -AppClass BizChat | Export-Csv copilot.csv -NoTypeInformation

    .NOTES
        Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048
        Learn: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
    #>

    [CmdletBinding()]
    [Alias('hcop')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('UserId', 'UserPrincipalName', 'Mailbox')]
        [string[]]$UserIds,

        [Parameter(ValueFromPipeline)]
        [psobject]$InputObject,

        [Parameter()]
        [datetime]$StartDate = (Get-Date).AddDays(-7),

        [Parameter()]
        [datetime]$EndDate = (Get-Date),

        [Parameter()]
        [string[]]$AppClass,

        [Parameter()]
        [ValidateRange(1, 100)]
        [int]$Top = 100,

        [Parameter()]
        [switch]$Nyan
    )

    begin {
        $abort = $false
        if (-not (Test-HuntEmailGraphConnection)) {
            $abort = $true
            return
        }

        $users = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $times = [System.Collections.Generic.List[datetime]]::new()
        $explicitStart = $PSBoundParameters.ContainsKey('StartDate')
        $explicitEnd = $PSBoundParameters.ContainsKey('EndDate')
    }

    process {
        if ($abort) { return }

        foreach ($id in @($UserIds)) {
            if ($id) { [void]$users.Add($id.Trim()) }
        }
        if ($InputObject -is [string] -and $InputObject) {
            [void]$users.Add($InputObject.Trim())
        }

        $stamp = $null
        if ($InputObject -and $InputObject.PSObject -and $InputObject.PSObject.Properties['CreationTime'] -and $InputObject.CreationTime) {
            $stamp = $InputObject.CreationTime
        }
        elseif ($InputObject -and $InputObject.PSObject -and $InputObject.PSObject.Properties['CreatedDateTime'] -and $InputObject.CreatedDateTime) {
            $stamp = $InputObject.CreatedDateTime
        }
        if ($stamp) {
            try { $times.Add([datetime]$stamp) } catch { Write-Verbose "Ignored an unparsable Copilot timestamp." }
        }
    }

    end {
        if ($abort) { return }
        if ($users.Count -eq 0) {
            Write-Host "[ERROR] huntCopilot needs at least one user id or UPN." -ForegroundColor Red
            return
        }

        if (-not $explicitStart -and -not $explicitEnd -and $times.Count -gt 0) {
            $StartDate = ($times | Measure-Object -Minimum).Minimum
            $EndDate = ($times | Measure-Object -Maximum).Maximum
            if ($StartDate -eq $EndDate) {
                $StartDate = $StartDate.AddMinutes(-5)
                $EndDate = $EndDate.AddMinutes(5)
            }
            Write-Host "[COPILOT] Using the piped audit time window: $StartDate -> $EndDate" -ForegroundColor Cyan
        }

        $rows = [System.Collections.Generic.List[object]]::new()
        $index = 0
        foreach ($userId in $users) {
            $index++
            $pct = [math]::Round(($index / $users.Count) * 100)
            if ($Nyan) {
                Write-NyanProgress -PercentComplete $pct -Status $userId -Activity "Copilot"
            }
            else {
                Write-Progress -Activity "Reading Copilot interactions" -Status $userId -PercentComplete $pct
            }

            $result = Get-HuntEmailCopilotInteraction -UserId $userId -StartDate $StartDate -EndDate $EndDate -AppClass $AppClass -Top $Top
            foreach ($item in @($result.Interactions)) {
                $rows.Add((ConvertTo-HuntEmailCopilotRow -UserId $userId -Item $item))
            }
            if ($result.Truncated) {
                Write-Warning "Copilot history for $userId was truncated at the page limit. Narrow the date range."
            }
        }

        if ($Nyan) { Write-NyanComplete -Activity "Copilot" }
        else { Write-Progress -Activity "Reading Copilot interactions" -Completed }

        Write-Host "[COPILOT] $($rows.Count) interaction(s) for $($users.Count) user(s)." -ForegroundColor Cyan
        return $rows.ToArray()
    }
}

function auditCopilot {
    <#
    .SYNOPSIS
        Read Purview unified audit events for Microsoft 365 Copilot interactions.

    .DESCRIPTION
        Wraps Search-UnifiedAuditLog -RecordType CopilotInteraction from Exchange
        Online PowerShell. Audit events contain metadata (user, time, app host,
        accessed resources, sensitivity labels) and do not contain prompt text.
        Pipe the results to huntCopilot to retrieve the prompts and responses.

        Requires a separate Connect-ExchangeOnline session. Microsoft Graph
        certificate auth is not this session.

        Paging uses one SessionId with -SessionCommand ReturnLargeSet until
        ResultIndex reaches ResultCount, or the cmdlet returns no rows.
        ReturnLargeSet is capped at 50,000 results by Exchange Online.

    .PARAMETER UserIds
        Optional user principal names. Passed to Search-UnifiedAuditLog -UserIds.

    .PARAMETER StartDate
        Start of the audit window. Default: 7 days ago. Sent to Exchange in UTC.

    .PARAMETER EndDate
        End of the audit window. Default: now. Sent to Exchange in UTC.

    .EXAMPLE
        auditCopilot -StartDate "2026-02-01" -EndDate "2026-02-18" -UserIds "user@example.com"

    .EXAMPLE
        acop -UserIds "user@example.com" -StartDate (Get-Date).AddDays(-2) | huntCopilot | Export-Csv copilot.csv -NoTypeInformation

    .NOTES
        Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048
        Schema: https://learn.microsoft.com/en-us/office/office-365-management-api/copilot-schema
        Cmdlet: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/search-unifiedauditlog
        Roles: Exchange Online Audit Logs or View-Only Audit Logs.
    #>

    [CmdletBinding()]
    [Alias('acop')]
    param (
        [Parameter()]
        [Alias('UserId', 'UserPrincipalName', 'Mailbox')]
        [string[]]$UserIds,

        [Parameter()]
        [datetime]$StartDate = (Get-Date).AddDays(-7),

        [Parameter()]
        [datetime]$EndDate = (Get-Date),

        [Parameter()]
        [switch]$Nyan
    )

    if (-not (Test-HuntEmailExchangeOnlineConnection)) {
        return
    }

    if ($StartDate -ge $EndDate) {
        Write-Host "[ERROR] StartDate must be earlier than EndDate." -ForegroundColor Red
        Write-Error "StartDate must be earlier than EndDate."
        return
    }

    $sessionId = [guid]::NewGuid().ToString()
    $rows = [System.Collections.Generic.List[object]]::new()
    $previousIndex = -1
    $page = 0

    do {
        $page++
        $parameters = @{
            StartDate      = ConvertTo-HuntEmailUtcDateTime $StartDate
            EndDate        = ConvertTo-HuntEmailUtcDateTime $EndDate
            RecordType     = 'CopilotInteraction'
            Operations     = 'CopilotInteraction'
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = 5000
            Formatted      = $true
        }
        if ($UserIds) { $parameters.UserIds = $UserIds }

        if ($Nyan) {
            Write-NyanProgress -PercentComplete ([math]::Min(99, $page * 5)) -Status "page $page" -Activity "Audit"
        }
        else {
            Write-Progress -Activity "Searching Copilot audit log" -Status "page $page ($($rows.Count) rows)" -PercentComplete ([math]::Min(99, $page * 5))
        }

        try {
            $batch = @(Invoke-HuntEmailAuditSearch -Parameters $parameters)
        }
        catch {
            Write-Host "[ERROR] Search-UnifiedAuditLog failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  Confirm Connect-ExchangeOnline and the Audit Logs or View-Only Audit Logs role." -ForegroundColor Yellow
            Write-Error "Search-UnifiedAuditLog failed: $($_.Exception.Message)"
            return
        }

        if ($batch.Count -eq 0 -or $null -eq $batch[0]) { break }

        foreach ($record in $batch) {
            if ($null -eq $record -or -not $record.AuditData) { continue }
            $rows.Add((ConvertTo-HuntEmailAuditRow -Record $record))
        }

        $last = $batch[-1]
        $resultIndex = 0
        $resultCount = 0
        if ($last -and $last.PSObject.Properties['ResultIndex'] -and $null -ne $last.ResultIndex) {
            $resultIndex = [int]$last.ResultIndex
        }
        if ($last -and $last.PSObject.Properties['ResultCount'] -and $null -ne $last.ResultCount) {
            $resultCount = [int]$last.ResultCount
        }

        $more = ($resultCount -gt 0) -and ($resultIndex -lt $resultCount) -and ($resultIndex -gt $previousIndex)
        $previousIndex = $resultIndex

        if ($rows.Count -ge 50000) {
            Write-Warning "Search-UnifiedAuditLog ReturnLargeSet stops at 50,000 results. Narrow -StartDate, -EndDate, or -UserIds."
            $more = $false
        }
    } while ($more)

    if ($Nyan) { Write-NyanComplete -Activity "Audit" }
    else { Write-Progress -Activity "Searching Copilot audit log" -Completed }

    Write-Host "[AUDIT] $($rows.Count) CopilotInteraction record(s). Pipe to huntCopilot for prompt and response text." -ForegroundColor Cyan
    return $rows.ToArray()
}

function Read-HuntEmailMenuLine {
    param (
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    $value = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    return $value.Trim()
}

function Start-HuntEmail {
    <#
    .SYNOPSIS
        Numbered menu that calls huntEmail toolkit functions.

    .DESCRIPTION
        Optional entry point. Direct function calls remain the primary interface.
        The menu prompts for key parameters and does not bypass purgeEmail safety rails.
    #>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'The menu only collects parameters. purgeEmail keeps its own ShouldProcess and PURGE confirmation.')]
    [CmdletBinding()]
    param ()

    Write-Host "huntEmail toolkit" -ForegroundColor Cyan
    Write-Host "Direct calls (huntEmail, timelineEmail, purgeEmail, huntCopilot, auditCopilot) are the primary interface." -ForegroundColor DarkGray
    Write-Host "Docs: $script:HuntEmailPlaybookDocs" -ForegroundColor DarkGray

    while ($true) {
        Write-Host ""
        Write-Host "  1) huntEmail (hunt)" -ForegroundColor White
        Write-Host "  2) timelineEmail (timeline)" -ForegroundColor White
        Write-Host "  3) purgeEmail (purge)" -ForegroundColor White
        Write-Host "  4) huntCopilot (hcop)  [read-only]" -ForegroundColor White
        Write-Host "  5) auditCopilot (acop) [read-only]" -ForegroundColor White
        Write-Host "  6) Connect-HuntEmailGraph" -ForegroundColor White
        Write-Host "  7) Import-HuntEmailCert" -ForegroundColor White
        Write-Host "  0) Exit" -ForegroundColor White
        $choice = Read-Host "Select"

        try {
            switch ($choice) {
                '1' {
                    $params = @{}
                    foreach ($pair in @(
                        @{ Name = 'Sender'; Prompt = 'Sender (blank to skip)' },
                        @{ Name = 'Subject'; Prompt = 'Subject (blank to skip)' },
                        @{ Name = 'Keywords'; Prompt = 'Keywords (blank to skip)' },
                        @{ Name = 'InternetMessageId'; Prompt = 'InternetMessageId (blank to skip)' },
                        @{ Name = 'UserPrincipalName'; Prompt = 'UserPrincipalName, comma-separated (blank = all licensed)' },
                        @{ Name = 'AttachmentName'; Prompt = 'AttachmentName (blank to skip)' },
                        @{ Name = 'ExportEml'; Prompt = 'ExportEml folder (blank to skip)' }
                    )) {
                        $entered = Read-HuntEmailMenuLine -Prompt $pair.Prompt
                        if ($entered) { $params[$pair.Name] = $entered }
                    }
                    if ($params.ContainsKey('UserPrincipalName')) {
                        $params.UserPrincipalName = @($params.UserPrincipalName -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                    $startText = Read-HuntEmailMenuLine -Prompt 'StartDate (blank = last 24 hours)'
                    $endText = Read-HuntEmailMenuLine -Prompt 'EndDate (blank = now)'
                    if ($startText) { $params.StartDate = [datetime]$startText }
                    if ($endText) { $params.EndDate = [datetime]$endText }
                    if ((Read-HuntEmailMenuLine -Prompt 'Detailed? (y/N)') -eq 'y') { $params.Detailed = $true }
                    huntEmail @params
                }
                '2' {
                    $params = @{}
                    $sender = Read-HuntEmailMenuLine -Prompt 'Sender (blank to skip)'
                    $subject = Read-HuntEmailMenuLine -Prompt 'Subject (blank to skip)'
                    if ($sender) { $params.Sender = $sender }
                    if ($subject) { $params.Subject = $subject }
                    $startText = Read-HuntEmailMenuLine -Prompt 'StartDate (blank = last 48 hours)'
                    $endText = Read-HuntEmailMenuLine -Prompt 'EndDate (blank = now)'
                    if ($startText) { $params.StartDate = [datetime]$startText }
                    if ($endText) { $params.EndDate = [datetime]$endText }
                    if ((Read-HuntEmailMenuLine -Prompt 'IncludeCopilot? (y/N)') -eq 'y') { $params.IncludeCopilot = $true }
                    timelineEmail @params
                }
                '3' {
                    Write-Host "purgeEmail still requires -WhatIf or a typed PURGE confirmation. The menu does not pass -Force." -ForegroundColor Yellow
                    $params = @{}
                    $sender = Read-HuntEmailMenuLine -Prompt 'Sender (blank to skip)'
                    $subject = Read-HuntEmailMenuLine -Prompt 'Subject (blank to skip)'
                    $upn = Read-HuntEmailMenuLine -Prompt 'UserPrincipalName, comma-separated (blank to skip)'
                    if ($sender) { $params.Sender = $sender }
                    if ($subject) { $params.Subject = $subject }
                    if ($upn) { $params.UserPrincipalName = @($upn -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                    $whatIf = Read-HuntEmailMenuLine -Prompt 'WhatIf first? (Y/n)'
                    if ($whatIf -ne 'n') { $params.WhatIf = $true }
                    purgeEmail @params
                }
                '4' {
                    $users = Read-HuntEmailMenuLine -Prompt 'UserIds, comma-separated'
                    if (-not $users) { Write-Host "At least one user is required." -ForegroundColor Yellow; continue }
                    $params = @{
                        UserIds = @($users -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                    $startText = Read-HuntEmailMenuLine -Prompt 'StartDate (blank = last 7 days)'
                    $endText = Read-HuntEmailMenuLine -Prompt 'EndDate (blank = now)'
                    $app = Read-HuntEmailMenuLine -Prompt 'AppClass Word,Excel,Teams,BizChat,WebChat,CoworkChat (blank = all)'
                    if ($startText) { $params.StartDate = [datetime]$startText }
                    if ($endText) { $params.EndDate = [datetime]$endText }
                    if ($app) { $params.AppClass = @($app -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                    huntCopilot @params
                }
                '5' {
                    $params = @{}
                    $users = Read-HuntEmailMenuLine -Prompt 'UserIds, comma-separated (blank = all)'
                    if ($users) { $params.UserIds = @($users -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                    $startText = Read-HuntEmailMenuLine -Prompt 'StartDate (blank = last 7 days)'
                    $endText = Read-HuntEmailMenuLine -Prompt 'EndDate (blank = now)'
                    if ($startText) { $params.StartDate = [datetime]$startText }
                    if ($endText) { $params.EndDate = [datetime]$endText }
                    auditCopilot @params
                }
                '6' {
                    $params = @{}
                    $tenant = Read-HuntEmailMenuLine -Prompt 'TenantId (blank = HUNTEMAIL_TENANT_ID)'
                    $client = Read-HuntEmailMenuLine -Prompt 'ClientId (blank = HUNTEMAIL_CLIENT_ID)'
                    $thumb = Read-HuntEmailMenuLine -Prompt 'Thumbprint (blank = HUNTEMAIL_CERT_THUMBPRINT)'
                    $pfx = Read-HuntEmailMenuLine -Prompt 'PfxPath (blank to skip)'
                    if ($tenant) { $params.TenantId = $tenant }
                    if ($client) { $params.ClientId = $client }
                    if ($thumb) { $params.Thumbprint = $thumb }
                    if ($pfx) { $params.PfxPath = $pfx }
                    if ($pfx -and (Read-HuntEmailMenuLine -Prompt 'Prompt for PFX password? (y/N)') -eq 'y') {
                        $params.Password = Read-Host 'PFX password' -AsSecureString
                    }
                    Connect-HuntEmailGraph @params
                }
                '7' {
                    $pfx = Read-HuntEmailMenuLine -Prompt 'PfxPath (blank = C:\certs\huntEmail.pfx)'
                    $params = @{ Password = (Read-Host 'PFX password' -AsSecureString) }
                    if ($pfx) { $params.PfxPath = $pfx }
                    Import-HuntEmailCert @params
                }
                '0' { return }
                default { Write-Host "Unknown selection." -ForegroundColor Yellow }
            }
        }
        catch {
            Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Export-ModuleMember -Function @(
    'Write-NyanProgress',
    'Write-NyanComplete',
    'Get-HuntEmailPagedCollection',
    'Get-HuntEmailFolderMap',
    'huntEmail',
    'timelineEmail',
    'purgeEmail',
    'Connect-HuntEmailGraph',
    'Import-HuntEmailCert',
    'huntCopilot',
    'auditCopilot',
    'Start-HuntEmail'
) -Alias @('hunt', 'timeline', 'purge', 'hcop', 'acop')
