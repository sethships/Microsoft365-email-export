#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '2.0.0' }

<#
.SYNOPSIS
    Bulk-exports an M365 mailbox to individual .eml files via Microsoft Graph.

.DESCRIPTION
    Authenticates as the signed-in user (delegated Mail.Read), walks every mail
    folder recursively, and downloads each message's raw MIME content using
    GET /me/messages/{id}/$value. Each message is saved as a standalone .eml
    file under <OutputRoot>/<Folder Path>/, with a manifest.jsonl entry
    containing folder, graph id, internetMessageId, subject, from,
    receivedDateTime, file path, and SHA-256 hash.

    Resumable: a state file (.export-state) tracks completed graph message ids;
    re-running the script skips messages already exported. Honors Graph
    throttling (429/503/504) with Retry-After or exponential backoff.

.PARAMETER OutputRoot
    Directory where exported .eml files, manifest, and state are written.
    Defaults to ./export/ relative to the script location.

.PARAMETER IncludeHiddenFolders
    Include hidden folders (e.g., Conversation History settings) in the walk.
    Default: $false.

.PARAMETER PageSize
    Page size for the messages list call. Default 50. Max 1000 per Graph spec,
    but smaller pages reduce memory and throttle risk.

.PARAMETER UseDeviceCode
    Authenticate with the OAuth 2.0 device-code flow instead of the default
    interactive (WAM / browser) flow. The script prints a URL and a short code
    that you open on any device to complete sign-in. Useful when:
      - The interactive WAM dialog gets dismissed or times out during 2FA
      - You're running over SSH / from a non-interactive session
      - You'd rather complete sign-in on a phone or a different machine

.PARAMETER TenantId
    Which Entra ID audience the sign-in should target. Accepts:
      - 'organizations' (DEFAULT) — any work/school account; personal
        Microsoft accounts (outlook.com, hotmail.com etc.) are rejected.
        Pick this if your mailbox is M365 work/school.
      - 'common'                 — accepts both work/school and personal
        accounts. Use this only if you genuinely want to export a
        personal outlook.com mailbox.
      - 'consumers'              — personal accounts only.
      - A tenant GUID            — pins sign-in to one specific tenant.
        Example: 'a1b2c3d4-1111-2222-3333-444444444444'
      - A verified domain name   — same effect as a GUID for that tenant.
        Example: 'contoso.onmicrosoft.com'

    Setting this prevents browser-cached personal-account SSO from hijacking
    the sign-in: if 'organizations' is set, the auth endpoint forces the
    browser to prompt for a work/school account fresh.

.EXAMPLE
    pwsh -File .\Export-M365Mail.ps1
    Export the signed-in user's full M365 mailbox to ./export/ (default
    audience is 'organizations' — personal Microsoft accounts are blocked).

.EXAMPLE
    pwsh -File .\Export-M365Mail.ps1 -OutputRoot 'D:\MailArchive\me'
    Export to an external path.

.EXAMPLE
    pwsh -File .\Export-M365Mail.ps1 -UseDeviceCode
    Use device-code sign-in — no time pressure for 2FA, completes on any browser.

.EXAMPLE
    pwsh -File .\Export-M365Mail.ps1 -UseDeviceCode -TenantId 'contoso.onmicrosoft.com'
    Pin sign-in to one specific tenant (handy when you have accounts in
    several tenants and need to disambiguate).

.EXAMPLE
    pwsh -File .\Export-M365Mail.ps1 -UseDeviceCode -TenantId 'common'
    Export a personal outlook.com / hotmail.com mailbox.

.NOTES
    Requires the Microsoft.Graph.Authentication PowerShell module.
    Install once with:  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [string] $OutputRoot = (Join-Path $PSScriptRoot 'export'),
    [switch] $IncludeHiddenFolders,
    [ValidateRange(1, 1000)]
    [int]    $PageSize = 50,
    [switch] $UseDeviceCode,
    [string] $TenantId = 'organizations'
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Filename component caps. Total path length must stay well under 260 on
# Windows unless long-path support is enabled. Folder paths can already be
# deep; we budget conservatively for the leaf filename.
$Script:MaxSenderLen   = 40
$Script:MaxSubjectLen  = 80
$Script:MaxFolderSeg   = 60
$Script:MaxFileNameLen = 180

# Retry policy
$Script:MaxRetryAttempts  = 8
$Script:RetryCeilingSecs  = 120

# Search Folders are virtual references — exporting them would duplicate
# messages that live in real folders. The well-known name is 'searchfolders'
# in Graph; we additionally skip by display name as a belt-and-suspenders
# safeguard for localized mailboxes.
$Script:SkipWellKnown    = @('searchfolders')
$Script:SkipDisplayNames = @('Search Folders')

# ---------------------------------------------------------------------------
# Helpers — paths, hashing, sanitization
# ---------------------------------------------------------------------------

function ConvertTo-SafePathSegment {
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Text,
        [int] $MaxLength = 80
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '_unnamed' }

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = [System.Text.StringBuilder]::new($Text.Length)
    foreach ($c in $Text.ToCharArray()) {
        if ($invalid -contains $c -or [int]$c -lt 32) {
            [void] $sb.Append('_')
        } else {
            [void] $sb.Append($c)
        }
    }
    $clean = $sb.ToString().Trim().Trim('.').Trim('_').Trim()
    if (-not $clean) { return '_unnamed' }
    if ($clean.Length -gt $MaxLength) {
        $clean = $clean.Substring(0, $MaxLength).TrimEnd()
    }
    # Avoid reserved Windows device names by suffixing if matched.
    $reserved = 'CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9'
    if ($reserved -contains $clean.ToUpperInvariant()) { return "_$clean" }
    return $clean
}

function Get-ShortHash {
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] [string] $Text,
        [int] $Length = 8
    )
    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $hasher.ComputeHash($bytes)
        $hex  = [System.BitConverter]::ToString($hash).Replace('-', '').ToLowerInvariant()
        return $hex.Substring(0, [Math]::Min($Length, $hex.Length))
    } finally {
        $hasher.Dispose()
    }
}

function Get-FileSha256 {
    [OutputType([string])]
    param([Parameter(Mandatory)] [string] $Path)
    # Use -LiteralPath, not -Path. Subjects routinely contain bracket
    # characters ([Newsletter], [ACTION REQUIRED], etc.) which Path treats
    # as wildcard character classes — for a filename containing literal
    # brackets, that wildcard pattern matches zero files and Get-FileHash
    # returns nothing, then .Hash explodes under StrictMode 3.
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Build-MessageFileName {
    [OutputType([string])]
    param([Parameter(Mandatory)] [hashtable] $Msg)

    $received = $Msg['receivedDateTime']
    if ($received) {
        try {
            $stamp = ([datetime] $received).ToUniversalTime().ToString('yyyy-MM-dd_HHmmss')
        } catch {
            $stamp = 'unknown-date'
        }
    } else {
        $stamp = 'unknown-date'
    }

    $senderAddr = 'unknown-sender'
    if ($Msg.ContainsKey('from') -and $Msg['from']) {
        $emailAddr = $Msg['from']['emailAddress']
        if ($emailAddr -and $emailAddr['address']) {
            $senderAddr = [string] $emailAddr['address']
        }
    }

    $subject = if ($Msg['subject']) { [string] $Msg['subject'] } else { 'no-subject' }

    $senderClean  = ConvertTo-SafePathSegment -Text $senderAddr -MaxLength $Script:MaxSenderLen
    $subjectClean = ConvertTo-SafePathSegment -Text $subject    -MaxLength $Script:MaxSubjectLen
    $idHash       = Get-ShortHash -Text ([string] $Msg['id']) -Length 8

    $name = "${stamp}_${senderClean}_${subjectClean}_${idHash}.eml"
    if ($name.Length -gt $Script:MaxFileNameLen) {
        # Subject is the most likely culprit; aggressively truncate it.
        $overflow = $name.Length - $Script:MaxFileNameLen
        $newSubjLen = [Math]::Max(8, $subjectClean.Length - $overflow)
        $subjectClean = $subjectClean.Substring(0, $newSubjLen).TrimEnd()
        $name = "${stamp}_${senderClean}_${subjectClean}_${idHash}.eml"
    }
    return $name
}

# ---------------------------------------------------------------------------
# Helpers — Graph throttling wrapper
# ---------------------------------------------------------------------------

function Invoke-GraphWithRetry {
    <#
    .SYNOPSIS
        Wraps Invoke-MgGraphRequest with Retry-After / exponential backoff for
        429, 503, and 504 responses. Returns the parsed body for JSON calls or
        $null when -OutputFilePath is used.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [string] $Method = 'GET',
        [string] $OutputFilePath,
        [hashtable] $Headers
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $splat = @{
                Method = $Method
                Uri    = $Uri
            }
            if ($OutputFilePath) { $splat['OutputFilePath'] = $OutputFilePath }
            if ($Headers)        { $splat['Headers']        = $Headers }
            return Invoke-MgGraphRequest @splat
        } catch {
            # Microsoft.Graph SDK surfaces HTTP failures in a few shapes. Try
            # to dig out a status code and Retry-After header from whatever
            # we got.
            $err     = $_
            $status  = 0
            $retryAfter = 0

            $response = $null
            if ($err.Exception -and $err.Exception.PSObject.Properties['Response']) {
                $response = $err.Exception.Response
            }
            if ($response) {
                try { $status = [int] $response.StatusCode } catch { $status = 0 }
                if ($response.Headers -and $response.Headers.Contains('Retry-After')) {
                    $hv = ($response.Headers.GetValues('Retry-After') | Select-Object -First 1)
                    [int]::TryParse([string]$hv, [ref] $retryAfter) | Out-Null
                }
            } else {
                # Fall back to parsing the message text.
                $msg = [string] $err.Exception.Message
                if ($msg -match '\b(429|503|504)\b') {
                    $status = [int] $Matches[1]
                }
            }

            $retryable = ($status -in 429, 503, 504) -or ($status -eq 0 -and $err.Exception.Message -match 'timed out|timeout|temporarily')

            if (-not $retryable -or $attempt -ge $Script:MaxRetryAttempts) {
                throw
            }

            if ($retryAfter -le 0) {
                $retryAfter = [Math]::Min($Script:RetryCeilingSecs, [int][Math]::Pow(2, $attempt))
            } else {
                $retryAfter = [Math]::Min($Script:RetryCeilingSecs, $retryAfter)
            }
            Write-Warning ("  throttled (status={0}); waiting {1}s (attempt {2}/{3}) — {4}" -f `
                $status, $retryAfter, $attempt, $Script:MaxRetryAttempts, $Uri)
            Start-Sleep -Seconds $retryAfter
        }
    }
}

# ---------------------------------------------------------------------------
# Helpers — manual OAuth device code flow
#
# Why we don't just use `Connect-MgGraph -UseDeviceCode`: on PowerShell 7
# with Microsoft.Graph 2.x, the SDK's device-code prompt is delivered through
# a host UI mechanism that silently fails to print in several common console
# configurations (transcribed sessions, spawned consoles, terminals where
# Console.Out isn't the same handle as the PS host). The cmdlet then sits
# waiting for input that the user can't see, and eventually times out with
# "Authentication timed out after 120 seconds due to inactivity."
#
# We work around this by driving the v2.0 device-code endpoints ourselves,
# printing the prompt via our own Write-Host (which we know reaches the
# transcript and the console reliably), then handing the resulting access
# token to Connect-MgGraph via -AccessToken.
# ---------------------------------------------------------------------------

# Microsoft Graph PowerShell multi-tenant first-party app id. This is the
# same client id Connect-MgGraph uses by default, so consent and scopes
# behave identically to the SDK's own flow.
$Script:GraphPwshClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'

function Get-DeviceCodeAccessToken {
    [OutputType([string])]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Scope,
        [Parameter(Mandatory)] [string] $TenantSegment
    )

    $deviceUri = "https://login.microsoftonline.com/$TenantSegment/oauth2/v2.0/devicecode"
    $tokenUri  = "https://login.microsoftonline.com/$TenantSegment/oauth2/v2.0/token"

    # Step 1: request a device + user code from the identity platform.
    try {
        $dc = Invoke-RestMethod -Method POST -Uri $deviceUri -ContentType 'application/x-www-form-urlencoded' -Body @{
            client_id = $Script:GraphPwshClientId
            scope     = $Scope
        } -ErrorAction Stop
    } catch {
        throw "Device-code request failed: $($_.Exception.Message)"
    }

    # Step 2: print the prompt prominently so the user can see it both on
    # the console and via Start-Transcript.
    Write-Host ""
    Write-Host "===============================================================" -ForegroundColor Yellow
    Write-Host "  DEVICE CODE SIGN-IN" -ForegroundColor Yellow
    Write-Host "===============================================================" -ForegroundColor Yellow
    Write-Host "  1. Open this URL in any browser:" -ForegroundColor White
    Write-Host "       $($dc.verification_uri)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  2. Enter this code when prompted:" -ForegroundColor White
    Write-Host "       $($dc.user_code)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  3. Sign in with the M365 account you want exported." -ForegroundColor White
    Write-Host "     Approve the Mail.Read consent on first use." -ForegroundColor White
    Write-Host ""
    Write-Host ("  Code expires in {0} minutes. Waiting for sign-in..." -f [Math]::Round($dc.expires_in / 60, 0)) -ForegroundColor DarkGray
    Write-Host "===============================================================" -ForegroundColor Yellow
    Write-Host ""

    # Step 3: poll the token endpoint until success, denial, or expiry.
    $interval = [int] $dc.interval
    if ($interval -lt 1) { $interval = 5 }
    $deadline = (Get-Date).AddSeconds([int] $dc.expires_in)
    $pollBody = @{
        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
        client_id   = $Script:GraphPwshClientId
        device_code = $dc.device_code
    }

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method POST -Uri $tokenUri -ContentType 'application/x-www-form-urlencoded' -Body $pollBody -ErrorAction Stop
            Write-Host "Device code accepted." -ForegroundColor Green
            return [string] $tok.access_token
        } catch {
            # Parse the OAuth error response. PS exposes the body via
            # ErrorDetails.Message on most platforms; fall back to reading
            # the response stream where that property is empty.
            $errBody = $null
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                try { $errBody = $_.ErrorDetails.Message | ConvertFrom-Json } catch { $errBody = $null }
            }
            if (-not $errBody -and $_.Exception.Response) {
                try {
                    $stream = $_.Exception.Response.GetResponseStream()
                    $stream.Position = 0
                    $reader = [System.IO.StreamReader]::new($stream)
                    $errBody = $reader.ReadToEnd() | ConvertFrom-Json
                } catch { $errBody = $null }
            }
            if (-not $errBody) { throw }

            switch ($errBody.error) {
                'authorization_pending' { continue }
                'slow_down'             { $interval += 5; continue }
                'expired_token'         { throw "Device code expired before sign-in completed." }
                'access_denied'         { throw "Sign-in was denied by the user." }
                default                 { throw "Token poll failed: $($errBody.error) — $($errBody.error_description)" }
            }
        }
    }
    throw "Device code expired before sign-in completed."
}

# ---------------------------------------------------------------------------
# Helpers — folder enumeration
# ---------------------------------------------------------------------------

function Get-AllMailFolders {
    <#
    .SYNOPSIS
        Recursively enumerates every mail folder for the signed-in user.
        Yields objects with Id, Path (display-name path joined by '/'),
        WellKnownName, TotalItemCount.
    #>
    [CmdletBinding()]
    param(
        [string] $ParentId = $null,
        [string] $PathPrefix = '',
        [switch] $IncludeHidden
    )

    # NOTE: wellKnownName is only available on the /beta endpoint, not /v1.0.
    # We deliberately omit it here and rely on $Script:SkipDisplayNames to
    # filter out the "Search Folders" virtual node. This is fine for English
    # mailboxes; non-English mailboxes may surface the localized search-folder
    # name (e.g. "Suchordner") and would need an extra entry there.
    $select = '$select=id,displayName,childFolderCount,totalItemCount'
    $top    = '$top=100'
    $hidden = if ($IncludeHidden) { 'includeHiddenFolders=true' } else { $null }
    $query  = @($select, $top, $hidden) | Where-Object { $_ } | ForEach-Object { $_ } # filter nulls
    $qstr   = ($query -join '&')

    $uri = if ($ParentId) {
        "/v1.0/me/mailFolders/$ParentId/childFolders?$qstr"
    } else {
        "/v1.0/me/mailFolders?$qstr"
    }

    while ($uri) {
        $resp = Invoke-GraphWithRetry -Uri $uri
        foreach ($f in $resp.value) {
            $wkn  = if ($f.ContainsKey('wellKnownName')) { [string] $f['wellKnownName'] } else { '' }
            $name = [string] $f['displayName']

            $skip = $false
            if ($wkn -and ($Script:SkipWellKnown -contains $wkn.ToLowerInvariant())) { $skip = $true }
            if (-not $skip -and ($Script:SkipDisplayNames -contains $name))         { $skip = $true }

            if ($skip) {
                Write-Verbose "  skipping virtual folder: $name"
                continue
            }

            $safeName = ConvertTo-SafePathSegment -Text $name -MaxLength $Script:MaxFolderSeg
            $path = if ($PathPrefix) { "$PathPrefix/$safeName" } else { $safeName }

            [pscustomobject] @{
                Id              = [string] $f['id']
                DisplayName     = $name
                Path            = $path
                WellKnownName   = $wkn
                TotalItemCount  = [int] $f['totalItemCount']
                ChildFolderCount = [int] $f['childFolderCount']
            }

            if ([int] $f['childFolderCount'] -gt 0) {
                Get-AllMailFolders -ParentId ([string] $f['id']) -PathPrefix $path -IncludeHidden:$IncludeHidden
            }
        }
        # Invoke-MgGraphRequest returns a Hashtable. ContainsKey() is the
        # correct test — PSObject.Properties on a Hashtable enumerates the
        # CLR type's properties (Keys, Values, Count, …), not the dictionary
        # entries, so the @odata.nextLink key would never be detected and
        # only page 1 of every collection would be processed.
        $uri = if ($resp.ContainsKey('@odata.nextLink')) { [string] $resp['@odata.nextLink'] } else { $null }
    }
}

# ---------------------------------------------------------------------------
# State (resume) — line-delimited graph message ids
# ---------------------------------------------------------------------------

function Import-CompletedState {
    # Returns a Hashtable used as a string set (key = graph message id,
    # value = $true). Hashtable is chosen over HashSet[string] because
    # PowerShell's strict-mode-3 PSObject adapter does not expose .Count
    # on generic HashSet, which causes "property Count cannot be found"
    # at the first reference. Hashtable.Count is reliable.
    [OutputType([System.Collections.Hashtable])]
    param([Parameter(Mandatory)] [string] $StatePath)

    $ht = @{}
    if (-not (Test-Path -LiteralPath $StatePath -PathType Leaf)) { return $ht }
    foreach ($line in [System.IO.File]::ReadLines($StatePath)) {
        $trimmed = $line.Trim()
        if ($trimmed) { $ht[$trimmed] = $true }
    }
    return $ht
}

function Add-CompletedState {
    param(
        [Parameter(Mandatory)] [string] $StatePath,
        [Parameter(Mandatory)] [string] $MessageId
    )
    Add-Content -LiteralPath $StatePath -Value $MessageId -Encoding utf8
}

# ---------------------------------------------------------------------------
# Manifest + error log writers
# ---------------------------------------------------------------------------

function Write-ManifestEntry {
    param(
        [Parameter(Mandatory)] [string] $ManifestPath,
        [Parameter(Mandatory)] [hashtable] $Entry
    )
    $json = $Entry | ConvertTo-Json -Depth 6 -Compress
    Add-Content -LiteralPath $ManifestPath -Value $json -Encoding utf8
}

function Write-ErrorEntry {
    param(
        [Parameter(Mandatory)] [string] $ErrorPath,
        [Parameter(Mandatory)] [hashtable] $Entry
    )
    $json = $Entry | ConvertTo-Json -Depth 6 -Compress
    Add-Content -LiteralPath $ErrorPath -Value $json -Encoding utf8
}

# ---------------------------------------------------------------------------
# Per-folder export
# ---------------------------------------------------------------------------

function Export-FolderMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] [string] $OutRoot,
        [Parameter(Mandatory)] [System.Collections.Hashtable] $Completed,
        [Parameter(Mandatory)] [string] $StatePath,
        [Parameter(Mandatory)] [string] $ManifestPath,
        [Parameter(Mandatory)] [string] $ErrorPath,
        [Parameter(Mandatory)] [int]    $PageSize
    )

    $folderDir = Join-Path -Path $OutRoot -ChildPath $Folder.Path
    if (-not (Test-Path -LiteralPath $folderDir)) {
        New-Item -ItemType Directory -Path $folderDir -Force | Out-Null
    }

    $select = '$select=id,internetMessageId,subject,from,receivedDateTime,hasAttachments'
    $top    = "`$top=$PageSize"
    $uri    = "/v1.0/me/mailFolders/$($Folder.Id)/messages?$select&$top"

    $count       = 0
    $skippedDone = 0
    $written     = 0
    $errors      = 0
    $startTime   = Get-Date

    Write-Host ""
    Write-Host ("[{0}] {1}  (~{2} items)" -f (Get-Date -Format 'HH:mm:ss'), $Folder.Path, $Folder.TotalItemCount) -ForegroundColor Cyan

    while ($uri) {
        $resp = Invoke-GraphWithRetry -Uri $uri
        foreach ($msg in $resp.value) {
            $count++
            $msgId = [string] $msg['id']

            if ($Completed.ContainsKey($msgId)) {
                $skippedDone++
                continue
            }

            $fileName = Build-MessageFileName -Msg $msg
            $outPath  = Join-Path -Path $folderDir -ChildPath $fileName

            # Collision: same timestamp + sender + subject + id-hash is
            # vanishingly unlikely, but if the file exists from a partial
            # prior run that didn't checkpoint, overwrite is correct.
            try {
                $mimeUri = "/v1.0/me/messages/$msgId/`$value"
                Invoke-GraphWithRetry -Uri $mimeUri -OutputFilePath $outPath | Out-Null
            } catch {
                $errors++
                $errEntry = @{
                    timestamp   = (Get-Date).ToUniversalTime().ToString('o')
                    folder      = $Folder.Path
                    graphId     = $msgId
                    stage       = 'download'
                    error       = $_.Exception.Message
                }
                Write-ErrorEntry -ErrorPath $ErrorPath -Entry $errEntry
                Write-Warning ("    download failed: {0} — {1}" -f $msgId, $_.Exception.Message)
                continue
            }

            try {
                $sha = Get-FileSha256 -Path $outPath
                $size = (Get-Item -LiteralPath $outPath).Length

                $fromAddr = $null
                if ($msg.ContainsKey('from') -and $msg['from']) {
                    $ea = $msg['from']['emailAddress']
                    if ($ea -and $ea['address']) { $fromAddr = [string] $ea['address'] }
                }

                $relative = $outPath.Substring($OutRoot.Length).TrimStart('\','/')

                $entry = @{
                    folder            = $Folder.Path
                    graphId           = $msgId
                    internetMessageId = [string] $msg['internetMessageId']
                    subject           = [string] $msg['subject']
                    from              = $fromAddr
                    receivedDateTime  = [string] $msg['receivedDateTime']
                    hasAttachments    = [bool]   $msg['hasAttachments']
                    relativePath      = $relative
                    sizeBytes         = $size
                    sha256            = $sha
                    exportedAt        = (Get-Date).ToUniversalTime().ToString('o')
                }
                Write-ManifestEntry -ManifestPath $ManifestPath -Entry $entry
                Add-CompletedState -StatePath $StatePath -MessageId $msgId
                $Completed[$msgId] = $true
                $written++
            } catch {
                $errors++
                $errEntry = @{
                    timestamp = (Get-Date).ToUniversalTime().ToString('o')
                    folder    = $Folder.Path
                    graphId   = $msgId
                    stage     = 'manifest'
                    error     = $_.Exception.Message
                }
                Write-ErrorEntry -ErrorPath $ErrorPath -Entry $errEntry
                Write-Warning ("    manifest failed: {0} — {1}" -f $msgId, $_.Exception.Message)
            }

            if (($written % 25) -eq 0 -and $written -gt 0) {
                $rate = if ((Get-Date) -gt $startTime) {
                    [Math]::Round($written / ((Get-Date) - $startTime).TotalSeconds, 2)
                } else { 0 }
                Write-Host ("    {0} written / {1} skipped / {2} errors  ({3} msg/s)" -f `
                    $written, $skippedDone, $errors, $rate)
            }
        }
        # Invoke-MgGraphRequest returns a Hashtable. ContainsKey() is the
        # correct test — PSObject.Properties on a Hashtable enumerates the
        # CLR type's properties (Keys, Values, Count, …), not the dictionary
        # entries, so the @odata.nextLink key would never be detected and
        # only page 1 of every collection would be processed.
        $uri = if ($resp.ContainsKey('@odata.nextLink')) { [string] $resp['@odata.nextLink'] } else { $null }
    }

    [pscustomobject] @{
        Folder      = $Folder.Path
        Total       = $count
        Written     = $written
        Skipped     = $skippedDone
        Errors      = $errors
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

function Invoke-Main {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $OutputRoot,
        [Parameter(Mandatory)] [bool]   $IncludeHidden,
        [Parameter(Mandatory)] [int]    $PageSize,
        [Parameter(Mandatory)] [bool]   $UseDeviceCode,
        [Parameter(Mandatory)] [string] $TenantId
    )

    # Validate dependencies first — fail fast with a clear message.
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw "Microsoft.Graph.Authentication module is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $OutputRoot)) {
        New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
    }
    $OutputRoot = (Resolve-Path -LiteralPath $OutputRoot).Path

    $manifestPath = Join-Path -Path $OutputRoot -ChildPath 'manifest.jsonl'
    $errorPath    = Join-Path -Path $OutputRoot -ChildPath 'errors.jsonl'
    $statePath    = Join-Path -Path $OutputRoot -ChildPath '.export-state'

    Write-Host "Output root : $OutputRoot"
    Write-Host "Manifest    : $manifestPath"
    Write-Host "Errors      : $errorPath"
    Write-Host "State       : $statePath"
    Write-Host ""

    $authMode = if ($UseDeviceCode) { 'device code (custom)' } else { 'interactive (WAM/browser)' }
    Write-Host "Connecting to Microsoft Graph (delegated, scope: Mail.Read; tenant: $TenantId; auth: $authMode)..." -ForegroundColor Cyan
    if ($UseDeviceCode) {
        $token = Get-DeviceCodeAccessToken `
            -Scope         'https://graph.microsoft.com/Mail.Read offline_access' `
            -TenantSegment $TenantId
        $secure = ConvertTo-SecureString $token -AsPlainText -Force
        Connect-MgGraph -AccessToken $secure -NoWelcome | Out-Null
    } else {
        # Connect-MgGraph's -TenantId accepts a GUID or domain. For the
        # multi-tenant audience aliases ('common', 'organizations',
        # 'consumers') we omit -TenantId because the SDK will pick the
        # right authority itself based on its built-in defaults.
        $connectArgs = @{ Scopes = 'Mail.Read'; NoWelcome = $true }
        if ($TenantId -notin 'common','organizations','consumers') {
            $connectArgs['TenantId'] = $TenantId
        }
        Connect-MgGraph @connectArgs | Out-Null
    }

    $ctx = Get-MgContext
    if (-not $ctx) {
        throw "Connect-MgGraph did not return a session context."
    }
    Write-Host ("Signed in as: {0}  (tenant: {1})" -f $ctx.Account, $ctx.TenantId) -ForegroundColor Green

    $completed = Import-CompletedState -StatePath $statePath
    if ($completed.Count -gt 0) {
        Write-Host ("Resume mode : {0} messages already exported — will be skipped" -f $completed.Count) -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Enumerating mail folders..." -ForegroundColor Cyan
    $folders = @(Get-AllMailFolders -IncludeHidden:$IncludeHidden)
    Write-Host ("Found {0} folders." -f $folders.Count) -ForegroundColor Green

    $totals = @{
        Folders = $folders.Count
        Written = 0
        Skipped = 0
        Errors  = 0
    }

    $runStart = Get-Date
    foreach ($folder in $folders) {
        if ($folder.TotalItemCount -eq 0) {
            Write-Host ("[{0}] {1}  (empty — skipping)" -f (Get-Date -Format 'HH:mm:ss'), $folder.Path) -ForegroundColor DarkGray
            continue
        }
        $result = Export-FolderMessages `
            -Folder       $folder `
            -OutRoot      $OutputRoot `
            -Completed    $completed `
            -StatePath    $statePath `
            -ManifestPath $manifestPath `
            -ErrorPath    $errorPath `
            -PageSize     $PageSize

        $totals.Written += $result.Written
        $totals.Skipped += $result.Skipped
        $totals.Errors  += $result.Errors
    }

    $elapsed = (Get-Date) - $runStart
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Green
    Write-Host "Export complete."
    Write-Host ("Folders     : {0}" -f $totals.Folders)
    Write-Host ("Written     : {0}" -f $totals.Written)
    Write-Host ("Skipped     : {0}" -f $totals.Skipped)
    Write-Host ("Errors      : {0}" -f $totals.Errors)
    Write-Host ("Elapsed     : {0:hh\:mm\:ss}" -f $elapsed)
    Write-Host "=================================================" -ForegroundColor Green
}

try {
    Invoke-Main `
        -OutputRoot     $OutputRoot `
        -IncludeHidden  $IncludeHiddenFolders.IsPresent `
        -PageSize       $PageSize `
        -UseDeviceCode  $UseDeviceCode.IsPresent `
        -TenantId       $TenantId
} finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
