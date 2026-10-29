# huntEmail — Email Threat Response Toolkit

**huntEmail** is a PowerShell toolkit for Microsoft 365 email threat response. It uses Microsoft Graph app-only certificate auth to search mailboxes, build a delivery timeline, and hard-delete confirmed phishing or HR-directed mail. It also reads Microsoft 365 Copilot prompts and Purview audit metadata. Everything loads from one module.

Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048

That link is the Purview Copilot eDiscovery playbook. Copilot content is not deleted by this toolkit.

---

## Features

| Capability | Function / alias | What it does |
|------------|------------------|--------------|
| **Hunt** | `huntEmail` / `hunt` | Mailbox search by sender, subject, keywords, attachment name, Internet Message-ID, and date range. Optional detailed body and `.eml` export. |
| **Timeline** | `timelineEmail` / `timeline` | Who received it, when, and whether it was read, deleted, junked, sent, or archived. Accepts hunt results on the pipeline. `-IncludeCopilot` flags Copilot use in that window. |
| **Purge** | `purgeEmail` / `purge` | Hard-deletes matching **mail** messages. `-WhatIf`, high `ConfirmImpact`, and a typed `PURGE` confirmation. |
| **Copilot content** | `huntCopilot` / `hcop` | Read-only. Prompts and responses from the Graph AI interaction history API. |
| **Copilot audit** | `auditCopilot` / `acop` | Read-only. `Search-UnifiedAuditLog -RecordType CopilotInteraction`. Metadata only, no prompt text. |
| **Menu** | `Start-HuntEmail` | Optional numbered menu. Direct function calls stay the primary interface. |

Helpers in the same module:

- `Connect-HuntEmailGraph` — certificate Graph connection (parameters or `HUNTEMAIL_*` environment variables)
- `Import-HuntEmailCert` — one-time PFX import into `Cert:\CurrentUser\My`
- `cert/createCert.ps1` — one-time self-signed `huntEmail.cer` / `huntEmail.pfx` (kept separate from the module)

---

## No Copilot purge

There is no function that deletes Copilot content.

- Prompt and response removal is done only through **Purview eDiscovery**, per the playbook linked above. Microsoft's supported admin path is an eDiscovery case, not `purgeEmail`. `purgeEmail` only calls `POST /users/{id}/messages/{id}/permanentDelete` for ordinary mail.
- **Copilot Memories** are not covered by Purview retention policies or retention labels. Microsoft Learn states that those policies do not apply to Copilot memory, and deleting a chat does not delete saved memories. This toolkit does not hold or delete them.
- **Copilot Pages** and Copilot Notebooks live in a user-owned **SharePoint Embedded** container (the container's application name is `Loop`). They are not mailbox messages.

---

## Install and import

PowerShell 7+ and the Microsoft Graph PowerShell module:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

`auditCopilot` also needs Exchange Online PowerShell. That is a separate sign-in from `Connect-MgGraph`:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline
```

One download. Either of these loads every function and alias:

```powershell
Import-Module .\huntEmail.psd1
```

```powershell
. .\huntEmail.ps1
```

Running `.\huntEmail.ps1` (not dot-sourcing it) loads the module and opens `Start-HuntEmail`.

`timelineEmail.ps1` and `purgeEmail.ps1` are compatibility shims that import the module. `Connect-HuntEmailGraph.ps1` and `Import-HuntEmailCert.ps1` are shims that call the functions of the same name.

---

## Permissions

Application permissions on the Entra app (admin consent):

| Permission | Used by |
|------------|---------|
| `Mail.Read` | `huntEmail`, `timelineEmail` |
| `Mail.ReadWrite` | `purgeEmail` |
| `User.Read.All` | Tenant-wide mailbox enumeration |
| `AiEnterpriseInteraction.Read.All` | `huntCopilot`, `timelineEmail -IncludeCopilot` |

`AiEnterpriseInteraction.Read.All` is application-only. Delegated permissions are not supported for this API. If it is not consented, `huntCopilot` stops with HTTP 401/403 and names the permission. The target user also needs a Microsoft 365 Copilot license with the **Microsoft Copilot with Graph-grounded chat** service plan.

`auditCopilot` does not use Graph. The Exchange Online account needs the **Audit Logs** or **View-Only Audit Logs** role (included by default in Compliance Management and Organization Management).

---

## Quick start

```powershell
Import-Module .\huntEmail.psd1
Connect-HuntEmailGraph

huntEmail -Sender "phish@bad.example" -StartDate "2026-02-15" -EndDate "2026-02-18" |
  timelineEmail |
  purgeEmail -WhatIf
```

```powershell
hunt -InternetMessageId "<ABC123@mail.example>"
huntEmail -Sender "phish@bad.example" | timelineEmail -IncludeCopilot
```

Copilot, audit metadata first, then the prompt text:

```powershell
Connect-ExchangeOnline
auditCopilot -UserIds "user@example.com" -StartDate (Get-Date).AddDays(-7) |
  huntCopilot |
  Export-Csv .\copilot.csv -NoTypeInformation
```

```powershell
hcop -UserIds "user@example.com" -AppClass BizChat -StartDate (Get-Date).AddDays(-2)
acop -UserIds "user@example.com" -StartDate (Get-Date).AddDays(-2) | Export-Csv .\audit.csv -NoTypeInformation
```

---

## huntCopilot

`GET https://graph.microsoft.com/v1.0/copilot/users/{id}/interactionHistory/getAllEnterpriseInteractions`

Confirmed on Microsoft Learn (v1.0 and beta are both documented; this module calls **v1.0**). The normal `/users/{id}/messages` calls used by `huntEmail` do not return these hidden Copilot items (`IPM.SkypeTeams.Message.Copilot.*`).

Query parameters used, matching Learn:

- `$top=100` (Learn's recommended page size). Further pages follow `@odata.nextLink`.
- `$filter` always sends **both** `createdDateTime gt` and `createdDateTime lt`. Learn requires both bounds and those operators are exclusive, so the module widens each bound by one second to keep the caller's range inclusive.
- Optional `appClass eq 'IPM.SkypeTeams.Message.Copilot.<Name>'`.

Supported `-AppClass` values from the Learn interaction-history filter list:

| `-AppClass` | Item class sent to Graph |
|-------------|--------------------------|
| `Word` | `IPM.SkypeTeams.Message.Copilot.Word` |
| `Excel` | `IPM.SkypeTeams.Message.Copilot.Excel` |
| `Teams` | `IPM.SkypeTeams.Message.Copilot.Teams` |
| `BizChat` | `IPM.SkypeTeams.Message.Copilot.BizChat` |
| `WebChat` | `IPM.SkypeTeams.Message.Copilot.WebChat` |
| `CoworkChat` | `IPM.SkypeTeams.Message.Copilot.CoworkChat` |

Learn says other interaction-history filters are not supported. Purview eDiscovery documents additional item classes (`Outlook`, `PowerPoint`, `Loop`, `Whiteboard`, and others). Pass no `-AppClass` and filter the `AppClass` column if you need to see whether the service returns them. This was not executed against a live tenant.

Each result is one flat object (prompt or response) with `UserId`, `CreatedDateTime`, `AppClass`, `InteractionType`, `Body`, and related context, link, and attachment columns. `Export-Csv` works on that output.

`auditCopilot` output pipes in on `UserId`. If you do not pass `-StartDate` or `-EndDate`, the window is the min/max `CreationTime` of the piped audit rows.

`-IncludeCopilot` on `timelineEmail` asks the same API, one page per recipient, whether that mailbox has any interaction in the delivery window. A missing `AiEnterpriseInteraction.Read.All` consent does not fail the timeline; the Copilot columns are left empty and the error names the permission.

---

## auditCopilot

Wraps Exchange Online:

```powershell
Search-UnifiedAuditLog `
  -RecordType CopilotInteraction `
  -Operations CopilotInteraction `
  -SessionId <guid> `
  -SessionCommand ReturnLargeSet `
  -ResultSize 5000 `
  -Formatted `
  -StartDate <utc> `
  -EndDate <utc> `
  -UserIds <upns>
```

The same `SessionId` and `ReturnLargeSet` are repeated until `ResultIndex` reaches `ResultCount`, or a page is empty. ReturnLargeSet stops at 50,000 rows; the function warns if it hits that cap.

`AuditData` is parsed with the Copilot schema on Learn. Flat columns include:

| Column | Schema source |
|--------|----------------|
| `UserId`, `CreationTime`, `Operation`, `RecordType`, `Workload`, `ClientRegion` | Audit record |
| `AppHost`, `AppIdentity`, `ThreadId` | `CopilotEventData` |
| `PromptMessageIds`, `ResponseMessageIds` | `Messages[]` (`isPrompt`) |
| `AccessedResources`, `AccessedResourceUrls`, `AccessedResourceTypes`, `AccessedResourceActions` | `AccessedResources[]` `Name`, `SiteUrl`, `Type`, `Action` |
| `SensitivityLabelIds` | `AccessedResources[].SensitivityLabelId` |
| `ContextIds`, `ContextTypes` | `Contexts[]` |
| `AISystemPlugins` | `AISystemPlugin[]` |
| `ModelProviders`, `ModelNames` | `ModelTransparencyDetails[]` |

The audit record does not contain the prompt or response text. Pipe to `huntCopilot` for that.

If `Search-UnifiedAuditLog` or `Connect-ExchangeOnline` is missing, the function says so and does not pretend Graph auth is enough.

---

## Safety

- Hard delete is irreversible. Start with `purgeEmail -WhatIf`.
- A live purge asks you to type `PURGE` unless you pass `-Force`. `ConfirmImpact` is `High`.
- Copilot functions are read-only. Do not point `purgeEmail` at Copilot item classes.
- Use a dedicated app registration and certificate auth. Do not commit `.pfx`, `.cer`, thumbprints, tenant IDs, client IDs, or evidence exports.
- Copilot bodies and audit rows are sensitive. Treat CSV output like any other incident export.

---

## Setup

1. `.\cert\createCert.ps1` writes `C:\certs\huntEmail.cer` and `C:\certs\huntEmail.pfx`. See [documentation/Certificate_Setup.md](documentation/Certificate_Setup.md).
2. Upload `huntEmail.cer` to the Entra app and consent the permissions above. Keep real IDs in a local copy of [documentation/api_info.md](documentation/api_info.md), not in git.
3. Import the PFX once per workstation: `Import-HuntEmailCert` (or `.\Import-HuntEmailCert.ps1`).
4. Set `HUNTEMAIL_TENANT_ID`, `HUNTEMAIL_CLIENT_ID`, and `HUNTEMAIL_CERT_THUMBPRINT`.
5. `Connect-HuntEmailGraph` (or `.\Connect-HuntEmailGraph.ps1`).

---

## Project layout

| Path | Purpose |
|------|---------|
| `huntEmail.psm1` | All functions |
| `huntEmail.psd1` | Module manifest |
| `huntEmail.ps1` | Imports the module. Opens the menu when executed |
| `timelineEmail.ps1`, `purgeEmail.ps1` | Deprecated shims that import the module |
| `Connect-HuntEmailGraph.ps1`, `Import-HuntEmailCert.ps1` | Shims that call the module functions |
| `cert/createCert.ps1` | One-time certificate creation |
| `documentation/` | Setup, API template, threat model, hunt walkthrough |
| `tests/Verify-HuntEmailModule.ps1` | Offline parse, export, and mocked URI checks |
| `LICENSE` | MIT |
| `SECURITY.md` | Vulnerability reporting and secret hygiene |

---

## Documentation

- [Certificate setup](documentation/Certificate_Setup.md)
- [API / config template](documentation/api_info.md)
- [Threat model](documentation/Threat_Model.md)
- [huntEmail explained](documentation/huntEmail_Explained.md)

Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048

---

## License

Released under the [MIT License](LICENSE). Copyright (c) 2026 huntEmail contributors.
