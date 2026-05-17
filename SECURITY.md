# Security Policy

## Reporting a vulnerability

If you've found a security issue in this project — for example a credential-handling bug, a path-traversal vulnerability in the export logic, or anything that could let an attacker exfiltrate mail content they shouldn't have access to — **please do not open a public issue**.

Instead, use GitHub's **[private vulnerability reporting](https://github.com/sethships/Microsoft365-email-export/security/advisories/new)** for this repository. That keeps the discussion private until a fix is in place.

> **Note**: private vulnerability reporting must be enabled by the repository owner under **Settings → Code security → Private vulnerability reporting**. If the link above returns a 404, please contact the owner through their GitHub profile (`@sethships`) and they'll enable it.

I aim to acknowledge reports within a few business days and to coordinate with the reporter on a fix and (if you'd like) public disclosure timing once a remediation is available.

## What's in scope

This project handles potentially sensitive material:

- **OAuth 2.0 access and refresh tokens** for Microsoft Graph — transient and held in memory only by design; should never touch disk in unencrypted form.
- **Mailbox content** in the form of `.eml` files under the user-configured `-OutputRoot`.
- **A `.export-state` file** containing Microsoft Graph message IDs (not credentials, but still identifying information).
- **A `manifest.jsonl` file** with metadata per exported message (subject, sender, dates, sizes, SHA-256 hashes).

Concerns worth reporting include but aren't limited to:

- Token leakage, improper handling, or transmission to anywhere other than `login.microsoftonline.com` / `graph.microsoft.com`.
- Insecure file permissions on output files (e.g., world-readable `.eml` files containing private mail).
- Path-traversal or command-injection paths via attacker-controlled mailbox content (folder names, message subjects, sender addresses, attachment filenames).
- Insufficient input validation on script parameters that could cause the script to write outside `-OutputRoot`.
- Bypass of the `-TenantId` / `organizations` audience constraint, allowing personal-MSA accounts to sign in when the user intended to scope to work/school only.

## Out of scope

- Vulnerabilities in **Microsoft Graph** itself — please report those to [Microsoft Security Response Center (MSRC)](https://msrc.microsoft.com/).
- Vulnerabilities in the **`Microsoft.Graph.Authentication` PowerShell module** — please report those upstream at <https://github.com/microsoftgraph/msgraph-sdk-powershell/security>.
- Self-XSS and similar issues that require the attacker already to have full control of the local machine on which the script runs.

## Supported versions

Only the **latest commit on `main`** is supported. There are no tagged releases at this time. If you encounter a security issue on an older commit, please verify it still reproduces on `main` before filing.
