# Security Policy

## Reporting a vulnerability

Please report security issues **privately** (for example via a private GitHub security advisory or a direct message to a maintainer). Do not open a public issue that includes exploit details, tenant identifiers, or sample credentials.

## Secrets and certificates

- Never commit `.pfx`, `.cer`, `.pem`, `.env`, real thumbprints, client secrets, or tenant/client IDs.
- Prefer **certificate authentication** over client secrets.
- Keep PFX passwords in a secret manager; do not embed them in scripts or profiles.
- Treat evidence exports (`*.csv`, `.eml`) as sensitive.

## Operational risk

`purgeEmail` performs an **irreversible hard delete** of ordinary mail. Always validate with hunt/timeline output and `-WhatIf` before a live purge. Type `PURGE` only when you intend permanent removal.

`huntCopilot` and `auditCopilot` are read-only. Their output includes prompts, responses, and audit metadata. Treat it as incident evidence. Copilot content removal is a Purview eDiscovery task, not a function in this toolkit.

Docs: https://app.notion.com/p/40809ccbb4e38122852fdb2fb2cc8048
