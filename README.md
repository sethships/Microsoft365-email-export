# Outlook-to-eml

A small PowerShell 7 script that bulk-exports your Microsoft 365 (Outlook) mailbox to individual `.eml` files via Microsoft Graph.

Each message is saved as a standalone `.eml` (raw MIME) using `GET /me/messages/{id}/$value`. Folder hierarchy is preserved on disk. A `manifest.jsonl` and `.export-state` file are produced alongside the output for auditability, dedup, and resumability.

## Why this design

This script is the "custom Graph MIME exporter" approach: it returns the exact same bytes a mail server would deliver, in the file format every email client understands. No PST, no proprietary archive. The tradeoff is that it only exports message content — not calendar, contacts, or mailbox metadata. For full-fidelity Exchange-to-Exchange backup, use Purview eDiscovery export instead.

## Prerequisites

- **Windows 10/11**
- **PowerShell 7+** — verify with `pwsh --version`
- **Microsoft.Graph.Authentication** PowerShell module (one-time install):
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- An M365 / Outlook.com account you can sign in to interactively
- No Entra app registration required — `Connect-MgGraph` uses Microsoft's well-known public client and prompts for `Mail.Read` consent on first run

## Usage

From the project root:

```powershell
# Default: exports to ./export/ (default tenant audience: 'organizations',
# so only work/school M365 accounts can sign in)
pwsh -File .\Export-M365Mail.ps1

# Recommended for the most reliable sign-in UX on Windows
pwsh -File .\Export-M365Mail.ps1 -UseDeviceCode

# Export to an external drive
pwsh -File .\Export-M365Mail.ps1 -OutputRoot 'D:\MailArchive\me'

# Pin sign-in to a specific tenant by domain or GUID (helps when you have
# accounts in multiple tenants)
pwsh -File .\Export-M365Mail.ps1 -UseDeviceCode -TenantId 'contoso.onmicrosoft.com'

# Export a personal outlook.com / hotmail.com mailbox (overrides default
# 'organizations' tenant which blocks personal MSA accounts)
pwsh -File .\Export-M365Mail.ps1 -UseDeviceCode -TenantId 'common'

# Include hidden folders (rarely needed)
pwsh -File .\Export-M365Mail.ps1 -IncludeHiddenFolders

# Smaller pages if you hit throttling early
pwsh -File .\Export-M365Mail.ps1 -PageSize 25
```

### Parameters

| Parameter | Default | Description |
| --- | --- | --- |
| `-OutputRoot` | `./export/` | Directory to write `.eml` files, `manifest.jsonl`, `errors.jsonl`, and `.export-state` |
| `-IncludeHiddenFolders` | off | Walk hidden folders too (system folders) |
| `-PageSize` | 50 | Messages per Graph page (1–1000) — smaller pages reduce throttle risk |
| `-UseDeviceCode` | off | Use the OAuth device-code flow instead of the default interactive WAM/browser flow |
| `-TenantId` | `organizations` | Audience for sign-in: `organizations` (work/school), `common` (any), `consumers` (personal-only), a tenant GUID, or a domain |

### Sign-in flows

**Default (interactive, no flags)**: relies on `Connect-MgGraph -Scopes Mail.Read`. On Windows this opens the native Windows Account Manager (WAM) dialog (the same one Office uses) or a browser tab. WAM has a hard 120-second sign-in timeout, so if your 2FA prompt takes longer (e.g., waiting for an authenticator code on another device), it will fail. The dialog can also be hidden behind other windows.

**`-UseDeviceCode` (recommended)**: prints a URL and a short code straight into the PowerShell window. Open the URL in any browser, enter the code, sign in. No time pressure during 2FA, no hidden dialog. The script polls the OAuth endpoint until you complete sign-in (or the code expires after 15 minutes).

If your browser is already signed into a personal Microsoft account via SSO/cookies and is "hijacking" the sign-in away from your work account, **complete the device-code sign-in in an InPrivate / Incognito browser window** so the browser has no cached Microsoft cookies. The default `-TenantId organizations` setting also rejects personal accounts at the endpoint level, which forces the sign-in page to ask for a work email.

Tokens acquired by `-UseDeviceCode` are not cached across script runs (by design — keeps the script's footprint zero on disk for secrets). Each fresh launch with `-UseDeviceCode` requires a new sign-in. The default flow does cache via the Graph SDK's MSAL cache.

## Output structure

```
<OutputRoot>/
  Inbox/
    2026-05-15_081533_alice@example.com_Quarterly report_a1b2c3d4.eml
    2026-05-15_092011_billing@vendor.com_Invoice 4421_e5f6a7b8.eml
  Sent Items/
    ...
  Archive/
    2024/
      Q3/
        ...
  manifest.jsonl
  errors.jsonl
  .export-state
```

### Filename pattern

`yyyy-MM-dd_HHmmss_<sender>_<subject>_<idhash8>.eml` (UTC timestamp, sanitized for Windows, length-capped). The 8-char hash of the Graph message id guarantees uniqueness if two messages share a sender/subject/timestamp.

### `manifest.jsonl`

One JSON object per line (JSONL), appended as each message is exported successfully. Fields:

| Field | Description |
| --- | --- |
| `folder` | Folder display-name path (e.g., `Inbox/Receipts`) |
| `graphId` | Microsoft Graph message id |
| `internetMessageId` | RFC 5322 `Message-ID:` header |
| `subject` | Message subject (may be empty) |
| `from` | Sender SMTP address (may be `null`) |
| `receivedDateTime` | ISO 8601 timestamp from Graph |
| `hasAttachments` | Boolean |
| `relativePath` | Path of the `.eml` file relative to `OutputRoot` |
| `sizeBytes` | Size of the `.eml` file on disk |
| `sha256` | SHA-256 hash of the `.eml` file (lowercase hex) |
| `exportedAt` | UTC ISO 8601 of when this entry was written |

### `errors.jsonl`

One JSON object per failure (download or manifest stage). Successful retries are not logged; only terminal failures appear here. Use this to identify messages that need manual re-export.

### `.export-state`

Line-delimited Graph message ids that were successfully exported. Re-running the script reads this file and skips matching messages. Delete it to force a full re-export.

## Resumability

The export is fully resumable:

- On startup, the script reads `.export-state` into memory and skips any message id already present.
- Each successful `.eml` write is checkpointed to `.export-state` *before* the next message is processed.
- Crashes, network drops, and throttling backoffs don't lose progress beyond the in-flight message.

To force re-export of a single message, delete the matching line from `.export-state` and delete the `.eml` file; the next run will fetch it again.

## Throttling

Microsoft Graph throttles aggressive callers. The script:

- Honors `Retry-After` on `429`, `503`, and `504` responses
- Falls back to exponential backoff (capped at 120s) when `Retry-After` is missing
- Retries up to 8 times per request before giving up and writing to `errors.jsonl`

For large mailboxes, expect intermittent waits — that's normal. Don't run multiple instances against the same mailbox in parallel; you'll just rate-limit yourself faster.

## Limitations / known gaps

- **Delegated auth only.** Exports the signed-in user's mailbox. To export someone else's mailbox or run unattended, you'd need to register your own Entra app with app-only `Mail.Read` permission and an `ApplicationAccessPolicy`. Not in scope for v1.
- **No calendar/contacts/tasks.** Only mail messages.
- **No attachments-as-separate-files.** Attachments are inside the `.eml` (which is correct — that's the whole point of MIME). If you need them as loose files, parse the `.eml` afterward with any MIME library.
- **Search Folders are skipped** by display name. Works for English mailboxes; non-English mailboxes (where the folder might be e.g. "Suchordner") would need an extra entry in `$Script:SkipDisplayNames`. The proper `wellKnownName` field exists only on the Graph `/beta` endpoint, which we deliberately don't use.
- **Long paths.** The script sanitizes and length-caps each path segment, but very deep folder trees on Windows without long-path support could still hit the 260-char limit. Enable long-path support or use a shorter `OutputRoot`.
- **No parallelism.** Single-threaded by design — simpler, easier to debug, and avoids self-throttling.
- **No refresh-token caching.** Each `-UseDeviceCode` run requires a fresh sign-in. Could be added by encrypting the refresh token with DPAPI under `%LOCALAPPDATA%\Outlook-to-eml\` if frequent re-runs become a pain.

## License

[MIT](LICENSE) © 2026 Seth ([github.com/sethships](https://github.com/sethships)).

## Code Statistics

| Language | files | blank | comment | code |
| :--- | ---: | ---: | ---: | ---: |
| PowerShell | 1 | 95 | 177 | 511 |
| Markdown | 1 | 48 | 0 | 114 |
| **SUM** | **2** | **143** | **177** | **625** |

*Generated with `cloc --vcs=git --md .` against the working tree on 2026-05-16, after the auth + state-handling fixes that followed end-to-end testing. Rerun after material changes.*
