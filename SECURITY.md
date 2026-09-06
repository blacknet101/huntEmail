# Security Policy

## Reporting a vulnerability

Please report security issues **privately** (for example via a private GitHub security advisory or a direct message to a maintainer). Do not open a public issue that includes exploit details, tenant identifiers, or sample credentials.

## Secrets and certificates

- Never commit `.pfx`, `.cer`, `.pem`, `.env`, real thumbprints, client secrets, or tenant/client IDs.
- Prefer **certificate authentication** over client secrets.
- Keep PFX passwords in a secret manager; do not embed them in scripts or profiles.
- Treat evidence exports (`*.csv`, `.eml`) as sensitive.

## Operational risk

`purgeEmail` performs an **irreversible hard delete**. Always validate with hunt/timeline output and `-WhatIf` before a live purge. Type `PURGE` only when you intend permanent removal.
