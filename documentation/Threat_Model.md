# Email Threat Response Toolkit -- Threat Model

**Author:** Security Operations
**Date:** 02/20/2026
**Updated:** 03/23/2026 -- Scrubbed for external sharing and generalized configuration guidance
**Status:** Reference
**Classification:** Restricted to authorized administrators

---

## 1. Overview

This document describes a generic threat model for the Email Threat Response Toolkit. The toolkit uses a Microsoft Entra application registration and Microsoft Graph application permissions to search, investigate, and purge email across a tenant.

**Toolkit location:** project root  
**Scripts:** `huntEmail.ps1`, `purgeEmail.ps1`, `timelineEmail.ps1`

---

## 2. Architecture Diagram

```text
+------------------------------------------------------------------+
|                 TRUST BOUNDARY: Microsoft 365 Tenant             |
|                                                                  |
|   +-------------------------+   +-----------------------------+  |
|   | Dedicated Email App     |   | Exchange Online            |  |
|   | - Mail.Read             |-->| - Read mailbox content     |  |
|   | - Mail.ReadWrite        |   | - Delete email             |  |
|   | - User.Read.All         |   | - Enumerate users          |  |
|   +-------------------------+   +-----------------------------+  |
|                 ^                                                |
|                 |                                                |
|                 +--- Certificate auth or short-lived secret      |
|                                                                  |
|   +-------------------------+                                    |
|   | Other admin apps        |                                    |
|   | - Device management     |                                    |
|   | - Directory automation  |                                    |
|   +-------------------------+                                    |
+------------------------------------------------------------------+
|                 TRUST BOUNDARY: Admin Workstations               |
|                                                                  |
|   +-------------------------+                                    |
|   | Admin workstation       |                                    |
|   | - Local cert store      |                                    |
|   | - PowerShell tooling    |                                    |
|   +-------------------------+                                    |
+------------------------------------------------------------------+
```

---

## 3. Security Design

### Use a dedicated app registration

The email toolkit should use a dedicated app registration with only these application permissions:

| Permission | Function | What It Does |
|------------|----------|--------------|
| `Mail.Read` | `huntEmail`, `timelineEmail` | Read message content, folders, and MIME export data |
| `Mail.ReadWrite` | `purgeEmail` | Delete messages from target mailboxes |
| `User.Read.All` | `huntEmail` | Enumerate users for tenant-wide searches |

Do not reuse a broad operational app that also holds device, directory, or endpoint-management permissions.

### Preferred authentication

- Preferred: certificate authentication from `Cert:\CurrentUser\My`
- Acceptable fallback: short-lived client secret stored outside source control
- Avoid: plaintext passwords, shared PFX paths, and credentials embedded in scripts

---

## 4. Trust Boundaries and Data Flow

### Trust Boundary 1: Tenant resources

**What's inside:** Mailboxes, user objects, message content, and service principal configuration.

**Who can cross it:** Anyone with access to the toolkit app's credential material.

**Risk:** A compromised credential allows tenant-wide mail search and deletion within the scope of the granted permissions.

### Trust Boundary 2: Admin workstations

**What's inside:** The operator workstation, PowerShell profile, local certificate store, and toolkit scripts.

**Risk:** A compromised workstation can expose certificate material, message exports, and operator activity.

### Trust Boundary 3: Secret and certificate storage

**What's inside:** PFX files, passwords, secret-manager entries, and backup copies of certificates.

**Risk:** If these assets are stored on broad-access shares or in plaintext notes, the app identity can be reused by an attacker.

---

## 5. Threat Analysis

### T1: Certificate theft

| | |
|---|---|
| **Threat** | Attacker obtains the PFX certificate and password |
| **Attack Vector** | Shared storage, endpoint compromise, or copied backup files |
| **Impact** | Full access to the toolkit app's Graph permissions |
| **Likelihood** | Medium |
| **Recommendation** | Import the cert to the local certificate store, restrict export, and keep the PFX off shared paths |

### T2: Client secret exposure

| | |
|---|---|
| **Threat** | Attacker obtains the client secret value |
| **Attack Vector** | Terminal history, copied docs, screenshots, or committed scripts |
| **Impact** | Full access to the toolkit app's Graph permissions |
| **Likelihood** | Medium |
| **Recommendation** | Prefer certificates; if a secret is required, use a secret manager and short expiry |

### T3: Overprivileged app registration

| | |
|---|---|
| **Threat** | The toolkit app has broader permissions than required |
| **Attack Vector** | Design error or reuse of a shared admin app |
| **Impact** | Higher blast radius across tenant services |
| **Likelihood** | Medium |
| **Recommendation** | Use a dedicated app with only `Mail.Read`, `Mail.ReadWrite`, and `User.Read.All` |

### T4: Unauthorized purge activity

| | |
|---|---|
| **Threat** | An operator or attacker deletes legitimate mail |
| **Attack Vector** | Broad purge criteria, misuse of `-Force`, or compromised admin session |
| **Impact** | Irreversible data loss |
| **Likelihood** | Low to Medium |
| **Recommendation** | Require explicit confirmation, log purge actions, and consider second-person approval for large purges |

### T5: Audit trail gaps

| | |
|---|---|
| **Threat** | Searches and deletions cannot be reconstructed later |
| **Attack Vector** | Missing local or centralized logging |
| **Impact** | Weak forensics and weak operational accountability |
| **Likelihood** | High unless explicitly addressed |
| **Recommendation** | Log operator, timestamp, criteria, result count, and deletion count to an approved audit destination |

---

## 6. Hardening Recommendations

### Priority 1

| # | Action | Effort | Impact |
|---|--------|--------|--------|
| 1 | Use a dedicated Microsoft Entra app for this toolkit only | Medium | Limits blast radius |
| 2 | Use certificate thumbprint auth from the local cert store | Low | Removes plaintext credential handling |
| 3 | Add centralized logging to hunt, timeline, and purge workflows | Medium | Improves auditability |

### Priority 2

| # | Action | Effort | Impact |
|---|--------|--------|--------|
| 4 | Remove PFX files from shared storage after secure import | Low | Reduces credential exposure |
| 5 | Rotate any fallback client secret on a short schedule | Low | Limits secret lifetime |
| 6 | Restrict app ownership to authorized administrators only | Low | Reduces privilege misuse |
| 7 | Restrict service principal sign-in locations where licensing allows | Medium | Reduces unauthorized use |

### Priority 3

| # | Action | Effort | Impact |
|---|--------|--------|--------|
| 8 | Store fallback secrets in a secret manager or key vault | Medium | Reduces secret sprawl |
| 9 | Require approval workflow for large purge actions | High | Reduces accidental deletion risk |
| 10 | Alert on unusual service principal sign-ins | Medium | Improves detection |

---

## 7. Audit Coverage

Microsoft Graph activity is partially visible through:

- Microsoft Entra sign-in logs for the service principal
- Microsoft Entra audit logs for app and permission changes
- Exchange or unified audit logging, depending on configuration

Additional toolkit-level logging is still recommended because tenant logs may not show:

- Which specific mailboxes were searched
- Which exact messages were deleted
- Which human operator initiated the action

---

## 8. Next Steps

- [ ] Populate local configuration with your own tenant-specific values
- [ ] Validate certificate-based authentication end to end
- [ ] Add centralized logging before broad operational use
- [ ] Review the purge workflow and approval threshold
- [ ] Store certificates and secrets outside source control
