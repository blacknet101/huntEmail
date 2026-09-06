# huntEmail.ps1 -- Script Breakdown

A section-by-section walkthrough of how `huntEmail` works under the hood.

**Project location:** project root  
**Script:** `huntEmail.ps1`  
**Related:** `purgeEmail.ps1`, `timelineEmail.ps1`, and docs in `documentation\`

**Last updated:** 2/24/2026 -- Reflects InternetMessageId search, attachment metadata, full body retrieval, EML export, Nyan mode, and PS 5.1/7 compatibility.

---

## 1. Nyan Cat Progress Bar

```powershell
function Write-NyanProgress { ... }
function Write-NyanComplete { ... }
```

Two helper functions draw a rainbow progress bar when `-Nyan` is used. They rely on `[char]0x2588` for the filled bar and `[char]0x2591` for the empty portion, with the cat `[=^.^=]` at the leading edge.

---

## 2. Function Declaration and Alias

```powershell
function huntEmail {
    [CmdletBinding()]
    [Alias("hunt")]
```

- The function is named `huntEmail` and also supports `hunt`.
- `[CmdletBinding()]` makes it an advanced function with common PowerShell parameters.

---

## 3. Parameters

```powershell
param (
    [string]$Sender,
    [string]$Subject,
    [string]$Keywords,
    [string]$InternetMessageId,
    [datetime]$StartDate = (Get-Date).AddHours(-24),
    [datetime]$EndDate = (Get-Date),
    [string[]]$UserPrincipalName,
    [string]$AttachmentName,
    [int]$MaxResults = 50,
    [switch]$Detailed,
    [string]$ExportEml,
    [switch]$Nyan
)
```

| Parameter | Purpose | Default |
|-----------|---------|---------|
| `$Sender` | Filter by sender email address | None |
| `$Subject` | Filter by subject line | None |
| `$Keywords` | Search the email body for these words | None |
| `$InternetMessageId` | Search by RFC 2822 Message-ID header | None |
| `$StartDate` | Beginning of the time window | 24 hours ago |
| `$EndDate` | End of the time window | Right now |
| `$UserPrincipalName` | Target specific mailbox(es) | All licensed mailboxes |
| `$AttachmentName` | Filter by attachment filename | None |
| `$MaxResults` | Max emails returned per mailbox | 50 |
| `$Detailed` | Show expanded detail and full body | Off |
| `$ExportEml` | Save raw `.eml` files for each result | None |
| `$Nyan` | Rainbow progress bar | Off |

At least one of `$Sender`, `$Subject`, `$Keywords`, `$AttachmentName`, or `$InternetMessageId` must be provided.

---

## 4. Begin Block -- Pre-Flight Checks

The begin block:

1. Verifies there is an active Microsoft Graph session with `Get-MgContext`
2. Validates that search criteria were provided
3. Auto-enables `-Detailed` when `-ExportEml` is set
4. Normalizes `InternetMessageId` with angle brackets when needed
5. Initializes a `List[PSObject]` for results

---

## 5. Process Block -- Main Flow

### 5a. Mailbox target list

- If `-UserPrincipalName` is supplied, only those mailboxes are searched.
- Otherwise the script enumerates all licensed, enabled users with mailboxes through Microsoft Graph.

### 5b. Query construction

There are three search modes:

1. `InternetMessageId` exact match using OData `$filter`
2. Keyword-style search using Graph `$search`
3. Date-based fallback using OData `$filter`

`internetMessageId` is included in the selected fields so it is always present in results.

### 5c. Mailbox loop

For each mailbox, the script:

- Builds the request URI
- Calls Microsoft Graph
- Applies client-side filtering for date, subject, and sender when needed

### 5d. Attachment processing

If a message has attachments, the script fetches attachment metadata and excludes inline content.

### 5e. Attachment-name matching

Attachment-name matching uses `[regex]::Escape($AttachmentName)` with `-match`, which correctly handles punctuation and other special characters in file names such as `Quarterly_Report_(Final).pdf`.

### 5f. Full body retrieval

When `-Detailed` is set, the script makes a separate Graph call to retrieve the full body and strips HTML tags for display.

### 5g. Result objects

Each result includes mailbox, subject, sender, recipients, timestamps, folder details, read state, attachments, preview or full body, Graph message ID, and InternetMessageId.

### 5h. EML export

When `-ExportEml` is set, the script downloads the raw MIME payload through the `/$value` endpoint and saves a sanitized `.eml` filename.

---

## 6. Operational Notes

- `InternetMessageId` is the best cross-mailbox identifier for the same email.
- `MessageId` is Graph's mailbox-specific ID and is required for delete or export operations.
- `-Detailed` adds more per-message API calls and is slower than summary mode.
- `-ExportEml` writes full message content to disk, so exported files should be handled as sensitive evidence.
