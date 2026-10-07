<# ============================================================================
  M365 & ENTRA ID GUEST GOVERNANCE DASHBOARD + PERIODIC CLEANUP (2-STAGE LIFECYCLE)
  ----------------------------------------------------------------------------
  Purpose : Inventory all Entra B2B guests with invited-by, invite status, full
            group/Teams membership, recency metrics, and trend charts; flag
            dormant/stale (>=90d inactive default), render CSV + HTML dashboard,
            email it, and optionally run 2-stage lifecycle: DISABLE stale
            guests, then DELETE 90 days after disable — with grace-period +
            state-file safety net.

  Context : SharePoint OTP (SPO OTP) retires mid-Oct 2026 -> end-Nov 2026.
            After retirement every external user needs a B2B guest account.
            This script is the periodic review/cleanup control for that
            guest population.

  Auth    : Same certificate app as OneDrive/SharePoint dashboards (app-only).
            Required app permissions (config.json):
              User.Read.All, User.ReadWrite.All (only for disable/delete),
              AuditLog.Read.All (accurate last-sign-in; else age proxy),
              Directory.Read.All, GroupMember.Read.All, Mail.Send

  Safety  : - DryRun enabled by default (config DryRun=true or parameter -DryRun).
            - In DryRun mode, no accounts are modified and state counters do NOT advance on disk.
            - To run live lifecycle: set config DryRun=false (or use -Live) AND specify
              -EnableCleanup (Stage 1 disable) and/or -EnableDelete (Stage 2 delete).
            - Stage 1 disable: -EnableCleanup  (or config AutoDisableStaleGuests).
            - Stage 2 delete:  -EnableDelete   (or config AutoDeleteDisabledGuests).
              Delete runs only on guests disabled for >= DeleteAfterDisabledDays
              (default 90), tracked in state via disabledOn/firstSeenDisabled.
            - A guest is disabled only after:
                1) inactive >= StaleThresholdDays (default 90), AND
                2) seen stale in >= RequireConsecutiveStaleRuns runs, AND
                3) GracePeriodDays elapsed since first seen stale.
            - Already-disabled guests start a 90-day delete clock on first
              observation (disabled date is not exposed by Graph).
            - State persisted in Reports\GuestCleanupState.json (skipped during DryRun).

  Usage   :
    .\M365-GuestGovernance-Dashboard.ps1                                  # report / dry run mode (default)
    .\M365-GuestGovernance-Dashboard.ps1 -DryRun                          # explicit dry run simulation
    .\M365-GuestGovernance-Dashboard.ps1 -Live -EnableCleanup             # live mode: disable eligible stale guests
    .\M365-GuestGovernance-Dashboard.ps1 -Live -EnableCleanup -EnableDelete # live mode: full 2-stage lifecycle
    .\M365-GuestGovernance-Dashboard.ps1 -EnableDelete -WhatIf            # preview what would be deleted
    .\M365-GuestGovernance-Dashboard.ps1 -StaleThresholdDays 60 -DormantThresholdDays 30
============================================================================ #>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$ConfigPath = "config.json",
  [string]$ReportPath,
  [int]$StaleThresholdDays,
  [int]$DormantThresholdDays,
  [int]$GracePeriodDays,
  [int]$DeleteAfterDisabledDays,
  [switch]$EnableCleanup,
  [switch]$DisableCleanup,
  [switch]$EnableDelete,
  [switch]$DryRun,
  [switch]$Live
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---------- Config ----------
$resolvedConfig = $ConfigPath
if (-not (Split-Path $ConfigPath -IsAbsolute)) { $resolvedConfig = Join-Path $scriptDir $ConfigPath }
if (-not (Test-Path $resolvedConfig)) { Write-Error "Config not found: $resolvedConfig"; exit 1 }
$cfg = Get-Content -Raw -Path $resolvedConfig | ConvertFrom-Json

$TenantId          = $cfg.TenantId
$ClientId          = $cfg.ClientId
$CertThumbprint    = $cfg.CertThumbprint
$Org               = if ($cfg.OrganizationName) { $cfg.OrganizationName } else { "Organization" }
if (-not $ReportPath) { $ReportPath = $cfg.ReportPath }
if (-not $StaleThresholdDays -or $StaleThresholdDays -le 0) { $StaleThresholdDays = if ($cfg.StaleThresholdDays) { [int]$cfg.StaleThresholdDays } else { 90 } }
if (-not $DormantThresholdDays -or $DormantThresholdDays -le 0) { $DormantThresholdDays = if ($cfg.DormantThresholdDays) { [int]$cfg.DormantThresholdDays } else { 30 } }
if (-not $GracePeriodDays) { $GracePeriodDays = if ($null -ne $cfg.GracePeriodDays) { [int]$cfg.GracePeriodDays } else { 7 } }
if (-not $DeleteAfterDisabledDays -or $DeleteAfterDisabledDays -le 0) { $DeleteAfterDisabledDays = if ($cfg.DeleteAfterDisabledDays) { [int]$cfg.DeleteAfterDisabledDays } else { 90 } }
$RequireRuns       = if ($cfg.RequireConsecutiveStaleRuns) { [int]$cfg.RequireConsecutiveStaleRuns } else { 2 }
$RetentionDays     = if ($cfg.ReportRetentionDays) { [int]$cfg.ReportRetentionDays } else { 90 }
$SenderEmail       = $cfg.SenderEmail
$RecipientEmail    = $cfg.RecipientEmail
$SubjectTemplate   = if ($cfg.EmailSubject) { $cfg.EmailSubject } else { "Guest User Review & Cleanup Dashboard - {Date}" }

# Determine DryRun mode:
# -Live switch forces live execution ($false)
# -DryRun switch forces simulation ($true)
# Else use config.json "DryRun" if present (default to $true for safety)
$IsDryRun = if ($Live) {
  $false
} elseif ($DryRun) {
  $true
} elseif ($null -ne $cfg.DryRun) {
  [bool]$cfg.DryRun
} else {
  $true
}

# Effective auto-disable / auto-delete:
if ($IsDryRun) {
  # In DryRun mode, hard-lock disable/delete to false to prevent any directory writes
  $AutoDisable = $false
  $AutoDelete  = $false
} else {
  $AutoDisable = if ($DisableCleanup) { $false } elseif ($EnableCleanup) { $true } elseif ($null -ne $cfg.AutoDisableStaleGuests) { [bool]$cfg.AutoDisableStaleGuests } else { $false }
  $AutoDelete  = if ($DisableCleanup) { $false } elseif ($EnableDelete)  { $true } elseif ($null -ne $cfg.AutoDeleteDisabledGuests) { [bool]$cfg.AutoDeleteDisabledGuests } else { $false }
}

if (-not (Split-Path $ReportPath -IsAbsolute)) { $ReportPath = Join-Path $scriptDir $ReportPath }
$ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null

$ScriptVersion      = "2.0"
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$csvName    = "${Org}_GuestUsers_$timestamp.csv"
$staleName  = "${Org}_GuestUsers_STALE_$timestamp.csv"
$htmlName   = "${Org}_GuestUsers_Dashboard_$timestamp.html"
$logName    = "${Org}_GuestUsers_Log_$timestamp.txt"
$csvPath    = Join-Path $ReportPath $csvName
$stalePath  = Join-Path $ReportPath $staleName
$htmlPath   = Join-Path $ReportPath $htmlName
$logPath    = Join-Path $ReportPath $logName
$statePath  = Join-Path $ReportPath "GuestCleanupState.json"

Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "$($Org.ToUpper()) GUEST USER DASHBOARD + CLEANUP" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
if ($IsDryRun) {
  Write-Host "MODE: DRY RUN (Simulation - No accounts will be disabled or deleted; state file will not be updated)" -ForegroundColor Yellow
} else {
  Write-Host "MODE: LIVE (Cleanup changes will be executed if armed)" -ForegroundColor Magenta
}
Write-Host "Stale>= $StaleThresholdDays d | Dormant>= $DormantThresholdDays d | Grace $GracePeriodDays d | RequireRuns $RequireRuns | DisableAfter $DeleteAfterDisabledDays d | DryRun: $IsDryRun | AutoDisable: $AutoDisable | AutoDelete: $AutoDelete" -ForegroundColor Gray

function Write-Log {
  param([string]$Message, [string]$Level = "INFO")
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $line = "[$ts] [$Level] $Message"
  switch ($Level) {
    "SUCCESS" { Write-Host $line -ForegroundColor Green }
    "ERROR"   { Write-Host $line -ForegroundColor Red }
    "WARNING" { Write-Host $line -ForegroundColor Yellow }
    default   { Write-Host $line }
  }
  try { $line | Out-File -FilePath $logPath -Append -Encoding UTF8 -WhatIf:$false } catch {}
}

function Get-GraphToken {
  Write-Log "Acquiring app-only Graph token (cert $CertThumbprint)..."
  $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$CertThumbprint" -ErrorAction SilentlyContinue
  if ($null -eq $cert) { $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction Stop }
  $b64 = { param([byte[]]$b) ([Convert]::ToBase64String($b)).TrimEnd('=').Replace('+','-').Replace('/','_') }
  $header  = @{ alg = "RS256"; typ = "JWT"; x5t = (& $b64 $cert.GetCertHash()) } | ConvertTo-Json -Compress
  $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
  $payload = @{ aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"; iss = $ClientId; sub = $ClientId; jti = [Guid]::NewGuid().ToString(); nbf = $now; exp = ($now + 300) } | ConvertTo-Json -Compress
  $enc = { param([string]$s) (& $b64 ([Text.Encoding]::UTF8.GetBytes($s))) }
  $h = & $enc $header; $p = & $enc $payload
  $toSign = "$h.$p"
  $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
  if ($null -eq $rsa) { throw "No RSA private key for cert $CertThumbprint." }
  $sig = (& $b64 ($rsa.SignData([Text.Encoding]::UTF8.GetBytes($toSign), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)))
  $assertion = "$h.$p.$sig"
  $body = @{ client_id = $ClientId; client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"; client_assertion = $assertion; scope = "https://graph.microsoft.com/.default"; grant_type = "client_credentials" }
  (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body).access_token
}

function Invoke-GraphPagedGet {
  param([string]$Uri, [hashtable]$Headers)
  $all = @(); $next = $Uri
  while ($next) {
    $r = Invoke-RestMethod -Uri $next -Method Get -Headers $Headers
    if ($r.value) { $all += $r.value } elseif ($r.id) { $all += $r }
    $next = $r.'@odata.nextLink'
  }
  return $all
}

function Get-AllGuestUsers {
  param([string]$Token)
  $headers = @{ Authorization = "Bearer $Token"; ConsistencyLevel = "eventual" }
  # Attempt 1: beta with signInActivity (needs AuditLog.Read.All).
  # NOTE: invited-by (sponsors API) is NOT expandable on user collections, so it is
  # resolved later via Graph $batch (see enrichment loop).
  try {
    Write-Log "Fetching all guest users (beta, incl. signInActivity)..."
    $select = "id,displayName,mail,userPrincipalName,createdDateTime,accountEnabled,externalUserState,creationType,signInActivity"
    $uri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Guest'&`$select=$select&`$top=999"
    $guests = Invoke-GraphPagedGet -Uri $uri -Headers $headers
    $script:SignInActivityAvailable = $true
    Write-Log "Found $($guests.Count) guest users (with sign-in activity)." -Level "SUCCESS"
    return $guests
  } catch {
    $msg = $_.Exception.Message
    $detail = ""
    try { $detail = ($_ | ConvertTo-Json -Depth 5) } catch {}
    if ($msg -like "*AuditLog.Read.All*" -or $detail -like "*AuditLog.Read.All*" -or $msg -like "*Authentication_MSGraphPermissionMissing*") {
      Write-Log "AuditLog.Read.All missing on app. Falling back to v1 without signInActivity (staleness by account age; invited-by via batch still attempted)." -Level "WARNING"
      $script:SignInActivityAvailable = $false
      $select2 = "id,displayName,mail,userPrincipalName,createdDateTime,accountEnabled,externalUserState,creationType"
      $uri2 = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'&`$select=$select2&`$top=999"
      $guests2 = Invoke-GraphPagedGet -Uri $uri2 -Headers @{ Authorization = "Bearer $Token" }
      Write-Log "Found $($guests2.Count) guest users (WITHOUT sign-in activity/invited-by - grant AuditLog.Read.All for accuracy)." -Level "SUCCESS"
      return $guests2
    }
    throw
  }
}

$PrivilegedGroupRegex = '\b(admin|admins|administrator|administrators|owner|owners|privileged|security|finance|hr|payroll|legal|executive|cfo|ceo)\b'

function Invoke-GraphBatch {
  # POST a set of sub-requests (max 20) to beta $batch; returns id -> response map.
  param([array]$SubRequests, [string]$Token)
  $body = @{ requests = $SubRequests } | ConvertTo-Json -Depth 6
  $attempt = 0
  while ($true) {
    try {
      $r = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/`$batch" -Method Post -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } -Body $body -ErrorAction Stop
      $map = @{}
      foreach ($resp in $r.responses) { $map[[string]$resp.id] = $resp }
      return $map
    } catch {
      $attempt++
      if ($attempt -ge 2) { throw }
      Start-Sleep -Seconds 5
    }
  }
  return @{}
}

function Load-State {
  if (Test-Path $statePath) {
    try { return (Get-Content -Raw -Path $statePath | ConvertFrom-Json -AsHashtable) } catch { Write-Log "State file corrupt, starting fresh." -Level "WARNING"; return @{} }
  }
  return @{}
}

# ---------- Main ----------
try {
  Write-Log "Script started. Report-only mode: $(-not $AutoDisable)"
  $token = Get-GraphToken
  Write-Log "Token acquired." -Level "SUCCESS"
  $headers = @{ Authorization = "Bearer $token" }

  $guests = Get-AllGuestUsers -Token $token
  $today = Get-Date
  $state = Load-State
  $seenIds = @()

  $rows = @()
  $signInOk = ($script:SignInActivityAvailable -ne $false)
  if (-not $signInOk) { Write-Log "Sign-in activity UNAVAILABLE - using CreatedDate as proxy for inactivity. Grant AuditLog.Read.All for true last-sign-in." -Level "WARNING" }
  foreach ($g in $guests) {
    $lastSignIn = $null
    try { $lastSignIn = $g.signInActivity.lastSignInDateTime } catch {}
    $created = $null
    try { if ($g.createdDateTime) { $created = [datetime]$g.createdDateTime } } catch {}
    $lastDt = $null
    try { if ($lastSignIn) { $lastDt = [datetime]$lastSignIn } } catch {}

    $ageDays = if ($created) { [math]::Round(($today - $created).TotalDays, 1) } else { $null }
    $inactiveDays = if ($lastDt) { [math]::Round(($today - $lastDt).TotalDays, 1) } else { $ageDays }
    $signInDays = if ($lastDt) { [math]::Round(($today - $lastDt).TotalDays, 1) } else { $null }
    if (-not $signInOk) { $neverSignedIn = $false } else { $neverSignedIn = ($null -eq $lastDt) }

    if (-not $signInOk) {
      # Fallback classification by age only; mark clearly
      if ($null -eq $ageDays) { $status = "Unknown" }
      elseif ($ageDays -ge $StaleThresholdDays) { $status = "Stale" }
      elseif ($ageDays -ge $DormantThresholdDays) { $status = "Dormant" }
      else { $status = "Active" }
    }
    elseif ($neverSignedIn) { $status = "NeverSignedIn" }
    elseif ($inactiveDays -ge $StaleThresholdDays) { $status = "Stale" }
    elseif ($inactiveDays -ge $DormantThresholdDays) { $status = "Dormant" }
    else { $status = "Active" }

    $enabled = [bool]$g.accountEnabled
    if (-not $enabled -and ($status -eq "Stale" -or $status -eq "NeverSignedIn")) { $status = "Stale-Disabled" }

    # Invited-by resolved later via Graph $batch (sponsors API; not expandable on collections)
    $sponName = ""; $sponEmail = ""
    # Invited domain: prefer mail, else UPN external part before #EXT#
    $invDomain = ""
    try {
      $src = if ($g.mail) { [string]$g.mail } else { ([string]$g.userPrincipalName -split "#EXT#")[0] -replace "_", "@" }
      if ($src -like "*@*") { $invDomain = ($src -split "@")[-1].ToLower() }
    } catch {}

    $rows += [PSCustomObject]@{
      DisplayName       = $g.displayName
      Email             = $g.mail
      UserPrincipalName = $g.userPrincipalName
      InvitedDomain     = $invDomain
      InvitedBy         = $sponName
      InvitedByEmail    = $sponEmail
      Id                = $g.id
      AccountEnabled    = $enabled
      ExternalUserState = $g.externalUserState
      InviteStatus      = if ($g.externalUserState) { [string]$g.externalUserState } else { "Unknown" }
      CreationType      = $g.creationType
      CreatedDate       = if ($created) { $created.ToString("yyyy-MM-dd") } else { "" }
      AgeDays           = $ageDays
      LastSignIn        = if ($lastDt) { $lastDt.ToString("yyyy-MM-dd") } elseif (-not $signInOk) { "Unknown (no AuditLog.Read.All)" } else { "Never" }
      SignInDays        = $signInDays   # $null = never signed in OR sign-in data unavailable
      InactiveDays      = $inactiveDays # sign-in recency, else account-age proxy (see Notes)
      Status            = $status
      Groups            = ""   # filled below for all guests
      GroupCount        = 0
      TeamCount         = 0
      PrivilegedGroups  = ""
      RecentlyInvited   = $false  # invited within last 7 days?
      BusinessOwner     = ""   # inviter where known, else Unassigned (see Notes)
      AccessSummary     = ""   # N groups / M Teams compact footprint
      FirstSeenStale    = ""
      Sightings         = ""
      DisableEligible   = ""
      DisabledOn        = ""
      DeleteDue         = ""
      RecommendedAction = ""
      CleanupAction     = "None"
      CleanupDetail     = ""
    }
  }

  # Enrich ALL guests with group/Teams membership + invited-by via Graph $batch
  # (10 guests = 20 sub-requests per batch: memberOf + sponsors each).
  Write-Log "Resolving groups + inviters for all $($rows.Count) guests via Graph batch..."
  $chunkSize = 10
  $batchFails = 0
  for ($c = 0; $c -lt $rows.Count; $c += $chunkSize) {
    $end = [Math]::Min($c + $chunkSize - 1, $rows.Count - 1)
    $chunk = @($rows[$c..$end])
    $subs = @()
    foreach ($u in $chunk) {
      # NOTE: batch sub-request urls inherit the version from the $batch endpoint - no /beta prefix.
      $subs += @{ id = "$($u.Id)|m"; method = "GET"; url = "/users/$($u.Id)/memberOf?`$select=displayName,id,groupTypes,resourceProvisioningOptions&`$top=100" }
      $subs += @{ id = "$($u.Id)|s"; method = "GET"; url = "/users/$($u.Id)/sponsors?`$select=id,displayName,mail,userPrincipalName" }
    }
    try {
      $map = Invoke-GraphBatch -SubRequests $subs -Token $token
      foreach ($u in $chunk) {
        $m = $map["$($u.Id)|m"]
        if ($m -and $m.status -eq 200 -and $m.body.value) {
          $grps = @($m.body.value | Where-Object { $_.'@odata.type' -like '*group*' })
          $names = @($grps | ForEach-Object { $_.displayName })
          $u.Groups = ($names -join "; ")
          $u.GroupCount = $names.Count
          $u.TeamCount = @($grps | Where-Object { $_.resourceProvisioningOptions -contains "Team" }).Count
          # Strip Org name case-insensitively before regex matching to prevent "OrgName" matching substrings
          $priv = @($names | Where-Object { ($clean = $_ -replace "(?i)$Org", "") -match $PrivilegedGroupRegex })
          $u.PrivilegedGroups = ($priv -join "; ")
        } elseif ($m -and $m.status -ne 200) {
          Write-Log "Sub-request memberOf for $($u.UserPrincipalName) returned HTTP $($m.status)" -Level "WARNING"
        }
        $s = $map["$($u.Id)|s"]
        if ($s -and $s.status -eq 200 -and $s.body.value -and @($s.body.value).Count -gt 0) {
          $sp0 = @($s.body.value)[0]
          $u.InvitedBy = [string]$sp0.displayName
          $u.InvitedByEmail = if ($sp0.mail) { [string]$sp0.mail } else { [string]$sp0.userPrincipalName }
        }
      }
    } catch {
      $batchFails++
      if ($batchFails -le 3) { Write-Log "Batch chunk failed ($($chunk.Count) guests): $_" -Level "WARNING" }
    }
    $done = [math]::Min($c + $chunkSize, $rows.Count)
    if ($done % 100 -eq 0 -or $done -eq $rows.Count) { Write-Progress -Activity "Resolving guest groups+inviters (batch)" -Status "$done of $($rows.Count)" -PercentComplete ([math]::Round($done / $rows.Count * 100, 1)) }
    Start-Sleep -Milliseconds 250
  }
  Write-Progress -Activity "Resolving guest groups+inviters (batch)" -Completed
  if ($batchFails -gt 0) { Write-Log "$batchFails batch chunks failed; affected guests keep blank groups/inviter." -Level "WARNING" }
  $invFound = @($rows | Where-Object { -not [string]::IsNullOrEmpty($_.InvitedByEmail) }).Count
  Write-Log "Enrichment done: inviters found for $invFound/$($rows.Count) guests." -Level "SUCCESS"

  # Flag recently invited guests (invited within the last 7 days)
  foreach ($r in $rows) {
    $r.RecentlyInvited = ($r.AgeDays -ne $null -and $r.AgeDays -le 7)
  }

  # ----- Periodic cleanup: STAGE 1 disable + STAGE 2 delete (90d after disable) -----
  $nowIso = (Get-Date).ToString("o")
  $now = Get-Date
  function Get-DisabledOnDate {
    param($Entry)
    foreach ($k in @("disabledOn", "firstSeenDisabled")) {
      if ($Entry.ContainsKey($k) -and $Entry[$k]) {
        try { return [datetime]$Entry[$k] } catch {}
      }
    }
    return $null
  }
  foreach ($r in $rows) {
    # --- Already-disabled guests: stage-2 delete clock ---
    if (-not $r.AccountEnabled) {
      if (-not $state.ContainsKey($r.Id)) {
        # First time we observe it disabled -> start delete clock now (Graph exposes no disabled date)
        $state[$r.Id] = @{ upn = $r.UserPrincipalName; firstSeenStale = $nowIso; staleRunCount = $RequireRuns; firstSeenDisabled = $nowIso }
        $r.CleanupAction = "DisabledHold"
        $r.CleanupDetail = "Already disabled; delete clock started $(Get-Date -Format 'yyyy-MM-dd'), eligible after $DeleteAfterDisabledDays d disabled."
      } else {
        $entry = $state[$r.Id]
        $entry.upn = $r.UserPrincipalName
        if (-not $entry.ContainsKey("firstSeenDisabled") -or -not $entry.firstSeenDisabled) {
          if ($entry.ContainsKey("disabledOn") -and $entry.disabledOn) { $entry.firstSeenDisabled = $entry.disabledOn }
          else { $entry.firstSeenDisabled = $nowIso }
        }
        $disSince = Get-DisabledOnDate $entry
        $daysDisabled = if ($disSince) { ($now - $disSince).TotalDays } else { 0 }
        $deleteDue = if ($disSince) { $disSince.AddDays($DeleteAfterDisabledDays) } else { $now.AddDays($DeleteAfterDisabledDays) }
        $r.CleanupDetail = "Disabled; delete due $($deleteDue.ToString('yyyy-MM-dd')) ($([math]::Round($daysDisabled,0))/$DeleteAfterDisabledDays d)."
        if ($daysDisabled -ge $DeleteAfterDisabledDays) {
          if ($IsDryRun) {
            $r.CleanupAction = "DryRun-Delete"
            $r.CleanupDetail = "Would delete (DryRun preview); disabled $([math]::Round($daysDisabled,0)) d (>= $DeleteAfterDisabledDays d threshold)."
          } elseif ($AutoDelete) {
            if ($PSCmdlet.ShouldProcess($r.UserPrincipalName, "DELETE guest (permanent, 90d+ disabled)")) {
              try {
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($r.Id)" -Method Delete -Headers @{ Authorization = "Bearer $token" } -ErrorAction Stop | Out-Null
                $r.CleanupAction = "Deleted"
                $r.CleanupDetail = "Deleted on $(Get-Date -Format 'yyyy-MM-dd') after $([math]::Round($daysDisabled,0)) d disabled."
                $state.Remove($r.Id)
                Write-Log "Deleted guest (90d+ disabled): $($r.DisplayName) ($($r.UserPrincipalName))" -Level "SUCCESS"
              } catch {
                $r.CleanupAction = "DeleteFailed"
                $r.CleanupDetail = $_.Exception.Message
                Write-Log "Failed to delete $($r.UserPrincipalName): $_" -Level "ERROR"
              }
            } else {
              $r.CleanupAction = "WhatIf-Delete"
              $r.CleanupDetail = "Would delete (WhatIf); disabled $([math]::Round($daysDisabled,0)) d."
            }
          } else {
            $r.CleanupAction = "DisabledHold"
            $r.CleanupDetail += " Auto-delete OFF (use -EnableDelete)."
          }
        } else {
          $r.CleanupAction = "DisabledHold"
        }
      }
      continue
    }
    # --- Enabled guests: stage-1 stale check ---
    $isStaleCandidate = ($r.Status -eq "Stale" -or $r.Status -eq "NeverSignedIn")
    if (-not $isStaleCandidate) {
      if ($state.ContainsKey($r.Id)) {
        # recovered -> drop stale watchlist entry (but never drop a disabled entry; handled above)
        $e = $state[$r.Id]
        $hasDis = ($e.ContainsKey("disabledOn") -and $e.disabledOn) -or ($e.ContainsKey("firstSeenDisabled") -and $e.firstSeenDisabled)
        if (($r.Status -eq "Active" -or $r.Status -eq "Dormant") -and -not $hasDis) { $state.Remove($r.Id) }
      }
      continue
    }
    $seenIds += $r.Id
    if (-not $state.ContainsKey($r.Id)) {
      $state[$r.Id] = @{ upn = $r.UserPrincipalName; firstSeenStale = $nowIso; staleRunCount = 1 }
      $r.CleanupAction = "PendingGrace"
      $r.CleanupDetail = "First seen stale; grace $GracePeriodDays d + $RequireRuns runs required."
    } else {
      $entry = $state[$r.Id]
      $entry.staleRunCount = [int]$entry.staleRunCount + 1
      $entry.upn = $r.UserPrincipalName
      $first = [datetime]$entry.firstSeenStale
      $graceElapsed = ((Get-Date) - $first).TotalDays -ge $GracePeriodDays
      $runsMet = [int]$entry.staleRunCount -ge $RequireRuns
      if ($graceElapsed -and $runsMet) {
        if ($IsDryRun) {
          $r.CleanupAction = "DryRun-Disable"
          $r.CleanupDetail = "Would disable (DryRun preview); runs=$($entry.staleRunCount)/$RequireRuns met, grace elapsed."
        } elseif ($AutoDisable) {
          if ($PSCmdlet.ShouldProcess($r.UserPrincipalName, "Disable guest (accountEnabled=false)")) {
            try {
              $patchUri = "https://graph.microsoft.com/v1.0/users/$($r.Id)"
              $patchBody = @{ accountEnabled = $false } | ConvertTo-Json
              Invoke-RestMethod -Uri $patchUri -Method Patch -Headers (@{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }) -Body $patchBody | Out-Null
              $r.CleanupAction = "Disabled"
              $r.CleanupDetail = "Disabled on $(Get-Date -Format 'yyyy-MM-dd'); delete due $((Get-Date).AddDays($DeleteAfterDisabledDays).ToString('yyyy-MM-dd')); runs=$($entry.staleRunCount)."
              $entry.disabledOn = $nowIso
              $entry.firstSeenDisabled = $nowIso
              Write-Log "Disabled stale guest: $($r.DisplayName) ($($r.UserPrincipalName)) — delete due in $DeleteAfterDisabledDays d" -Level "SUCCESS"
            } catch {
              $r.CleanupAction = "DisableFailed"
              $r.CleanupDetail = $_.Exception.Message
              Write-Log "Failed to disable $($r.UserPrincipalName): $_" -Level "ERROR"
            }
          } else {
            $r.CleanupAction = "WhatIf-Disable"
            $r.CleanupDetail = "Would disable (WhatIf); runs=$($entry.staleRunCount)."
          }
        } else {
          $reason = @()
          if (-not $AutoDisable) { $reason += "auto-disable OFF (use -EnableCleanup)" }
          $r.CleanupAction = "PendingGrace"
          $r.CleanupDetail = ($reason -join "; ")
        }
      } else {
        $reason = @()
        if (-not $runsMet) { $reason += "runs $($entry.staleRunCount)/$RequireRuns" }
        if (-not $graceElapsed) { $reason += "grace pending since $($first.ToString('yyyy-MM-dd'))" }
        $r.CleanupAction = "PendingGrace"
        $r.CleanupDetail = ($reason -join "; ")
      }
    }
    Start-Sleep -Milliseconds 50
  }
  # Prune state entries for guests no longer present
  $allIds = @($rows | ForEach-Object { $_.Id })
  foreach ($k in @($state.Keys)) { if ($k -notin $allIds) { $state.Remove($k) } }
  if ($WhatIfPreference -or $IsDryRun) {
    Write-Log "DryRun / WhatIf mode active: state file NOT saved (grace/delete counters unchanged on disk)." -Level "WARNING"
  } else {
    $state | ConvertTo-Json -Depth 5 | Out-File -FilePath $statePath -Encoding UTF8 -WhatIf:$false
  }

  # ----- Second pass: lifecycle audit fields + business owner + recommended action -----
  foreach ($r in $rows) {
    # Business owner = inviter where known; else Unassigned (see Notes)
    $r.BusinessOwner = if (-not [string]::IsNullOrEmpty($r.InvitedByEmail)) { $r.InvitedByEmail } `
      elseif (-not [string]::IsNullOrEmpty($r.InvitedDomain)) { "Unassigned (domain: $($r.InvitedDomain))" } `
      else { "Unassigned" }
    # Compact access footprint
    $acc = "$($r.GroupCount) groups / $($r.TeamCount) Teams"
    if (-not [string]::IsNullOrEmpty($r.PrivilegedGroups)) { $acc += " / PRIVILEGED" }
    $r.AccessSummary = $acc
    # Lifecycle audit fields from state
    if ($state.ContainsKey($r.Id)) {
      $e = $state[$r.Id]
      try { $r.FirstSeenStale = ([datetime]$e.firstSeenStale).ToString("yyyy-MM-dd") } catch {}
      $runs = try { [int]$e.staleRunCount } catch { 0 }
      $r.Sightings = "$runs/$RequireRuns"
      try { $r.DisableEligible = ([datetime]$e.firstSeenStale).AddDays($GracePeriodDays).ToString("yyyy-MM-dd") } catch {}
      $disRaw = if ($e.ContainsKey("disabledOn") -and $e.disabledOn) { $e.disabledOn } elseif ($e.ContainsKey("firstSeenDisabled") -and $e.firstSeenDisabled) { $e.firstSeenDisabled } else { $null }
      if ($disRaw) {
        try {
          $disDt = [datetime]$disRaw
          $r.DisabledOn = $disDt.ToString("yyyy-MM-dd")
          $r.DeleteDue = $disDt.AddDays($DeleteAfterDisabledDays).ToString("yyyy-MM-dd")
        } catch {}
      }
    } else {
      $r.Sightings = if ($r.Status -eq "Stale" -or $r.Status -eq "NeverSignedIn") { "0/$RequireRuns" } else { "-" }
    }
    # Recommended action (single governance decision per guest)
    $rec = switch ($r.CleanupAction) {
      "Deleted"        { "Deleted" }
      "Disabled"       { "Disabled this run — hold until $($r.DeleteDue)" }
      "WhatIf-Delete"  { "Would delete (preview)" }
      "WhatIf-Disable" { "Would disable (preview)" }
      "DryRun-Delete"  { "Would delete (DryRun preview)" }
      "DryRun-Disable" { "Would disable (DryRun preview)" }
      "DisableFailed"  { "Disable failed — retry" }
      "DeleteFailed"   { "Delete failed — retry" }
      default {
        if (-not $r.AccountEnabled) {
          if ($r.DeleteDue) { "Already disabled — delete due $($r.DeleteDue)" } else { "Already disabled — hold" }
        }
        elseif ($r.Status -eq "Active") { "Keep" }
        elseif ($r.Status -eq "Unknown") { "Investigate (no dates)" }
        elseif ($r.Status -eq "Dormant") { "Review (Dormant)" }
        elseif ($r.InviteStatus -eq "PendingAcceptance" -and $r.AgeDays -ne $null -and $r.AgeDays -lt $StaleThresholdDays) { "Awaiting acceptance" }
        elseif ($r.Status -eq "Stale" -or $r.Status -eq "NeverSignedIn") { "Disable after review" }
        else { "In lifecycle watch" }
      }
    }
    $r.RecommendedAction = $rec
  }

  # ----- Exports -----
  # NOTE: file writes bypass -WhatIf so preview runs still produce report+email;
  # -WhatIf only gates the Graph disable/delete calls via ShouldProcess.
  $rows | Sort-Object Status, InactiveDays -Descending | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
  $staleOnly = $rows | Where-Object { $_.Status -like "Stale*" -or $_.Status -eq "NeverSignedIn" }
  $staleOnly | Export-Csv -Path $stalePath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
  Write-Log "CSV: $csvName ($($rows.Count)) | STALE: $($staleOnly.Count)" -Level "SUCCESS"

  # Retention
  Get-ChildItem -Path $ReportPath -Filter "${Org}_GuestUsers_*.csv" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays) } | Remove-Item -Force -ErrorAction SilentlyContinue
  Get-ChildItem -Path $ReportPath -Filter "${Org}_GuestUsers_*.html" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays) } | Remove-Item -Force -ErrorAction SilentlyContinue

  # ----- Stats -----
  $total = $rows.Count
  $active = @($rows | Where-Object { $_.Status -eq "Active" }).Count
  $dormant = @($rows | Where-Object { $_.Status -eq "Dormant" }).Count
  $stale = @($rows | Where-Object { $_.Status -eq "Stale" }).Count
  $never = @($rows | Where-Object { $_.Status -eq "NeverSignedIn" }).Count
  $staleDis = @($rows | Where-Object { $_.Status -eq "Stale-Disabled" }).Count
  $enabledC = @($rows | Where-Object { $_.AccountEnabled }).Count
  $disabledC = $total - $enabledC
  $pending = @($rows | Where-Object { $_.CleanupAction -eq "PendingGrace" }).Count
  $justDisabled = @($rows | Where-Object { $_.CleanupAction -eq "Disabled" }).Count
  $whatif = @($rows | Where-Object { $_.CleanupAction -eq "WhatIf-Disable" }).Count
  $failed = @($rows | Where-Object { $_.CleanupAction -eq "DisableFailed" }).Count
  $onHold = @($rows | Where-Object { $_.CleanupAction -eq "DisabledHold" }).Count
  $justDeleted = @($rows | Where-Object { $_.CleanupAction -eq "Deleted" }).Count
  $whatifDel = @($rows | Where-Object { $_.CleanupAction -eq "WhatIf-Delete" }).Count
  $delFailed = @($rows | Where-Object { $_.CleanupAction -eq "DeleteFailed" }).Count
  $recentInvites = @($rows | Where-Object { $_.AgeDays -ne $null -and $_.AgeDays -le 7 })
  $recentInvitesC = $recentInvites.Count
  $recentPct = if ($total -gt 0) { [math]::Round($recentInvitesC / $total * 100, 1) } else { 0 }
  $pendingAccept = @($rows | Where-Object { $_.InviteStatus -eq "PendingAcceptance" }).Count
  $acceptedC = @($rows | Where-Object { $_.InviteStatus -eq "Accepted" }).Count
  $acceptRate = if ($total -gt 0) { [math]::Round($acceptedC / $total * 100, 1) } else { 0 }
  $noInviter = @($rows | Where-Object { [string]::IsNullOrEmpty($_.InvitedByEmail) }).Count
  $privGuests = @($rows | Where-Object { -not [string]::IsNullOrEmpty($_.PrivilegedGroups) }).Count
  $teamGuests = @($rows | Where-Object { $_.TeamCount -gt 0 }).Count
  $noGroups = @($rows | Where-Object { $_.GroupCount -eq 0 }).Count
  $unknownC = @($rows | Where-Object { $_.Status -eq "Unknown" }).Count
  $withSignInDate = @($rows | Where-Object { $_.SignInDays -ne $null }).Count
  $neverKnown = @($rows | Where-Object { $_.Status -eq "NeverSignedIn" }).Count
  # Data confidence: sign-in from Entra vs age proxy
  if ($signInOk) { $confidence = "High — last-sign-in from Entra signInActivity"; $proxyC = 0 }
  else { $confidence = "Medium — account-age proxy (grant AuditLog.Read.All for true sign-in)"; $proxyC = $total }
  $pct = { param($n) if ($total -gt 0) { [math]::Round($n / $total * 100, 1) } else { 0 } }
  # Top invited domains
  $topDomains = $rows | Where-Object { -not [string]::IsNullOrEmpty($_.InvitedDomain) } | Group-Object InvitedDomain | Sort-Object Count -Descending | Select-Object -First 8
  $maxDomain = if ($topDomains) { ($topDomains | Measure-Object Count -Maximum).Maximum } else { 1 }
  # Top invited domains table
  $ConsumerDomains = @("gmail.com","googlemail.com","outlook.com","hotmail.com","live.com","msn.com","yahoo.com","ymail.com","icloud.com","me.com","mac.com","aol.com","protonmail.com","proton.me","gmx.com","yandex.com","rediffmail.com")
  $domainTable = @()
  foreach ($dg in $topDomains) {
    $dm = @($dg.Group)
    $dActive = @($dm | Where-Object { $_.Status -eq "Active" -or $_.Status -eq "Dormant" }).Count
    $dStale = @($dm | Where-Object { $_.Status -eq "Stale" -or $_.Status -eq "NeverSignedIn" -or $_.Status -eq "Stale-Disabled" }).Count
    $dtype = if ($dg.Name -in $ConsumerDomains) { "Consumer" } else { "Corporate" }
    $domainTable += [PSCustomObject]@{ Domain = $dg.Name; Type = $dtype; Guests = $dg.Count; Active = $dActive; StaleNever = $dStale }
  }
  # LAST SIGN-IN recency buckets (true sign-in dates only; never/unknown as own rows)
  $siBuckets = @(
    @{ Label = "<30d"; Min = 0; Max = 30 },
    @{ Label = "30–90d"; Min = 31; Max = 90 },
    @{ Label = "91–180d"; Min = 91; Max = 180 },
    @{ Label = "181–365d"; Min = 181; Max = 365 },
    @{ Label = "365d+"; Min = 366; Max = 999999 }
  )
  $siRows = @()
  $maxSi = 1
  foreach ($b in $siBuckets) {
    $c = @($rows | Where-Object { $_.SignInDays -ne $null -and $_.SignInDays -ge $b.Min -and $_.SignInDays -le $b.Max }).Count
    if ($c -gt $maxSi) { $maxSi = $c }
    $siRows += [PSCustomObject]@{ Label = $b.Label; Count = $c }
  }
  $siRows += [PSCustomObject]@{ Label = "Never signed in"; Count = $neverKnown }
  $siRows += [PSCustomObject]@{ Label = "Unknown (no date)"; Count = @($rows | Where-Object { $_.SignInDays -eq $null -and $_.Status -ne "NeverSignedIn" }).Count }
  if ((($siRows | Measure-Object Count -Maximum).Maximum) -gt $maxSi) { $maxSi = ($siRows | Measure-Object Count -Maximum).Maximum }
  # ACCOUNT AGE buckets (creation recency; separate measurement from sign-in)
  $ageBuckets = @(
    @{ Label = "<30d"; Min = 0; Max = 30 },
    @{ Label = "30–90d"; Min = 31; Max = 90 },
    @{ Label = "91–180d"; Min = 91; Max = 180 },
    @{ Label = "181–365d"; Min = 181; Max = 365 },
    @{ Label = "365d+"; Min = 366; Max = 999999 }
  )
  $ageRows = @()
  $maxAge = 1
  foreach ($b in $ageBuckets) {
    $c = @($rows | Where-Object { $_.AgeDays -ne $null -and $_.AgeDays -ge $b.Min -and $_.AgeDays -le $b.Max }).Count
    if ($c -gt $maxAge) { $maxAge = $c }
    $ageRows += [PSCustomObject]@{ Label = $b.Label; Count = $c }
  }
  $ageRows += [PSCustomObject]@{ Label = "Unknown"; Count = @($rows | Where-Object { $_.AgeDays -eq $null }).Count }
  if ((($ageRows | Measure-Object Count -Maximum).Maximum) -gt $maxAge) { $maxAge = ($ageRows | Measure-Object Count -Maximum).Maximum }
  # Creation trend: last 12 months
  $trendRows = @()
  for ($m = 11; $m -ge 0; $m--) {
    $d = (Get-Date -Day 1).AddMonths(-$m)
    $key = $d.ToString("yyyy-MM")
    $c = @($rows | Where-Object { $_.CreatedDate -like "$key*" }).Count
    $mlabel = $d.ToString("MMM yy")
    if ($m -eq 0) { $mlabel += " MTD" }
    $trendRows += [PSCustomObject]@{ Month = $mlabel; Count = $c }
  }
  $maxTrend = [math]::Max(1, (($trendRows | Measure-Object Count -Maximum).Maximum))
  # Recommended-action counts (executive summary + actions table)
  $recDisable = @($rows | Where-Object { $_.RecommendedAction -eq "Disable after review" }).Count
  $recOwner = @($rows | Where-Object { $_.RecommendedAction -eq "Owner confirmation" }).Count
  $recApprove = @($rows | Where-Object { $_.RecommendedAction -like "Approve*" -or $_.RecommendedAction -like "Dormant*" }).Count
  $recKeep = @($rows | Where-Object { $_.RecommendedAction -eq "Keep" }).Count
  $recAwait = @($rows | Where-Object { $_.RecommendedAction -eq "Awaiting acceptance" }).Count
  # Trend rollups: Jun–Sep dominant period + Oct MTD
  $junSepSum = (($trendRows | Select-Object -Skip 7 -First 4 | Measure-Object Count -Sum).Sum)
  if ($null -eq $junSepSum) { $junSepSum = 0 }
  $octMTD = ($trendRows | Select-Object -Last 1).Count
  $peakRow = ($trendRows | Sort-Object Count -Descending | Select-Object -First 1)
  $peakMonth = $peakRow.Month
  $peakCount = $peakRow.Count

  # ----- HTML email (Outlook-compatible: tables + inline CSS only, no JS/forms) -----
  # Priority queue for action list (action needed and stale guests first)
  $actionRank = @{
    "DryRun-Delete"  = 0
    "DryRun-Disable" = 1
    "WhatIf-Delete"  = 0
    "WhatIf-Disable" = 1
    "Deleted"        = 2
    "Disabled"       = 2
    "PendingGrace"   = 3
    "DisabledHold"   = 4
    "None"           = 5
  }
  $queueList = $rows | Sort-Object @{ Expression = { if ($actionRank.ContainsKey($_.CleanupAction)) { $actionRank[$_.CleanupAction] } else { 9 } } }, @{ Expression = { if ($_.InactiveDays -eq $null) { -1 } else { $_.InactiveDays } }; Descending = $true }
  $esc = { param($s) if ($null -eq $s) { "" } else { [System.Net.WebUtility]::HtmlEncode([string]$s) } }
  $recColor = {
    param($rec)
    if ($rec -like "Disable*" -or $rec -like "Delete eligible*" -or $rec -like "Would delete*") { "#dc2626" }
    elseif ($rec -like "Owner confirmation*" -or $rec -like "Approve*" -or $rec -like "Awaiting*" -or $rec -like "Review*") { "#d97706" }
    elseif ($rec -like "Keep*") { "#16a34a" } else { "#475569" }
  }

  function Render-QueueRow {
    param($u, $n)
    $stBg = switch ($u.Status) { "Active" { "#dcfce7" } "Dormant" { "#fef3c7" } "Stale" { "#ffedd5" } "NeverSignedIn" { "#fee2e2" } default { "#f1f5f9" } }
    $stFg = switch ($u.Status) { "Active" { "#16a34a" } "Dormant" { "#d97706" } "Stale" { "#ea580c" } "NeverSignedIn" { "#dc2626" } default { "#475569" } }
    $lc = @()
    if ($u.Sightings -and $u.Sightings -ne "-") { $lc += "sight $($u.Sightings)" }
    if ($u.DisableEligible) { $lc += "dis-elig $($u.DisableEligible)" }
    if ($u.DeleteDue) { $lc += "del-due $($u.DeleteDue)" }
    $lcTxt = if ($lc) { ($lc -join " • ") } else { "-" }
    $inact = if ($u.InactiveDays -ne $null) { "$($u.InactiveDays)d inactive" } else { "-" }
    $ownerTxt = if ($u.BusinessOwner) { $u.BusinessOwner } else { "Unassigned" }
    return "<tr style='border-bottom:1px solid #f1f5f9;'><td style='padding:8px 6px;text-align:center;font-size:11px;color:#64748b;'>$n</td><td style='padding:8px 6px;'><strong>$(& $esc $u.DisplayName)</strong><br><span style='font-size:11px;color:#64748b;'>$(& $esc $u.UserPrincipalName)</span></td><td style='padding:8px 6px;text-align:center;'><span style='background:$stBg;color:$stFg;border-radius:12px;padding:2px 8px;font-size:10px;font-weight:600;'>$($u.Status)</span><br><span style='font-size:10px;color:#64748b;'>$inact</span></td><td style='padding:8px 6px;font-size:11px;'><strong>$(& $esc $u.AccessSummary)</strong><br><span style='color:#64748b;'>Owner: $(& $esc $ownerTxt)</span></td><td style='padding:8px 6px;font-size:11px;'><strong style='color:$(& $recColor $u.RecommendedAction);'>$(& $esc $u.RecommendedAction)</strong><br><span style='font-size:10px;color:#64748b;'>$lcTxt</span></td></tr>"
  }
  $queueHead = "<tr style='background:#f1f5f9;'><td style='padding:8px 6px; font-weight:600; text-align:center; font-size:11px; width:28px;'>#</td><td style='padding:8px 6px; font-weight:600; font-size:11px;'>Guest</td><td style='padding:8px 6px; font-weight:600; text-align:center; font-size:11px; width:110px;'>Status</td><td style='padding:8px 6px; font-weight:600; font-size:11px;'>Access & Business Owner</td><td style='padding:8px 6px; font-weight:600; font-size:11px;'>Action / Lifecycle</td></tr>"

  # Email table 1: Recently invited guests (last 7 days)
  $topRecentRowsHtml = ""
  $n = 1
  foreach ($u in (@($rows | Where-Object { $_.AgeDays -ne $null -and $_.AgeDays -le 7 } | Sort-Object AgeDays | Select-Object -First 10))) {
    $topRecentRowsHtml += Render-QueueRow $u $n
    $n++
  }
  if (-not $topRecentRowsHtml) { $topRecentRowsHtml = "<tr><td colspan='5' style='padding:20px;text-align:center;color:#64748b;'>No guests invited in the last 7 days.</td></tr>" }

  # Email table 2: top 10 requiring action (anything but Keep / Awaiting acceptance)
  $topActionRowsHtml = ""
  $n = 1
  foreach ($u in (@($queueList | Where-Object { $_.RecommendedAction -ne "Keep" -and $_.RecommendedAction -ne "Awaiting acceptance" } | Select-Object -First 10))) {
    $topActionRowsHtml += Render-QueueRow $u $n
    $n++
  }
  if (-not $topActionRowsHtml) { $topActionRowsHtml = "<tr><td colspan='5' style='padding:20px;text-align:center;color:#64748b;'>Nothing requiring action.</td></tr>" }

  # Invite-domain bars (fixed-pixel widths: % widths + &nbsp; don't scale in Outlook)
  $barMax = 280
  $domainRowsHtml = ""
  foreach ($d in $topDomains) {
    $px = if ($maxDomain -gt 0) { [math]::Round($d.Count / $maxDomain * $barMax) } else { 0 }
    if ($d.Count -gt 0 -and $px -lt 4) { $px = 4 }
    $rest = $barMax - $px
    if ($px -gt 0) { $bar = "<table width='$barMax' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:4px;'><tr><td width='$px' bgcolor='#2563eb' height='8' style='line-height:0;font-size:0;border-radius:4px;'>&nbsp;</td><td width='$rest' height='8' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>" }
    else { $bar = "<table width='$barMax' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:4px;'><tr><td height='8' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>" }
    $domainRowsHtml += "<tr><td style='padding:6px 8px;font-size:12px;'>$($d.Name)</td><td style='padding:6px 8px;'><table width='350' cellpadding='0' cellspacing='0'><tr><td width='280'>$bar</td><td width='10' style='line-height:0;font-size:0;'>&nbsp;</td><td width='60' style='font-weight:600;font-size:12px;'>$($d.Count)</td></tr></table></td></tr>"
  }
  if (-not $domainRowsHtml) { $domainRowsHtml = "<tr><td style='padding:10px;color:#64748b;font-size:12px;'>No domain data.</td></tr>" }

  # Creation trend bars (fixed-pixel widths, green)
  $trendRowsHtml = ""
  foreach ($t in $trendRows) {
    $px = if ($maxTrend -gt 0) { [math]::Round($t.Count / $maxTrend * $barMax) } else { 0 }
    if ($t.Count -gt 0 -and $px -lt 4) { $px = 4 }
    $rest = $barMax - $px
    if ($px -gt 0) { $bar = "<table width='$barMax' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:4px;'><tr><td width='$px' bgcolor='#16a34a' height='8' style='line-height:0;font-size:0;border-radius:4px;'>&nbsp;</td><td width='$rest' height='8' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>" }
    else { $bar = "<table width='$barMax' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:4px;'><tr><td height='8' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>" }
    $trendRowsHtml += "<tr><td style='padding:5px 8px;font-size:11px;' width='70'>$($t.Month)</td><td style='padding:5px 8px;'><table width='350' cellpadding='0' cellspacing='0'><tr><td width='280'>$bar</td><td width='10' style='line-height:0;font-size:0;'>&nbsp;</td><td width='60' style='font-size:11px;font-weight:600;'>$($t.Count)</td></tr></table></td></tr>"
  }
  # Compact Pixel-bar row helper for 2-column cards (fits 880px layout)
  function New-PixelBarRow {
    param([string]$Label, [int]$Count, [int]$Max, [string]$Color, [int]$LabelWidth = 70)
    $barMaxCompact = 180
    $px = if ($Max -gt 0) { [math]::Round($Count / $Max * $barMaxCompact) } else { 0 }
    if ($Count -gt 0 -and $px -lt 4) { $px = 4 }
    $rest = $barMaxCompact - $px
    if ($px -gt 0) { $bar = "<table width='$barMaxCompact' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:4px;'><tr><td width='$px' bgcolor='$Color' height='8' style='line-height:0;font-size:0;border-radius:4px;'>&nbsp;</td><td width='$rest' height='8' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>" }
    else { $bar = "<table width='$barMaxCompact' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:4px;'><tr><td height='8' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>" }
    return "<tr><td style='padding:5px 6px;font-size:11px;' width='$LabelWidth'>$Label</td><td style='padding:5px 6px;'><table width='240' cellpadding='0' cellspacing='0'><tr><td width='180'>$bar</td><td width='10' style='line-height:0;font-size:0;'>&nbsp;</td><td width='50' style='font-size:11px;font-weight:600;'>$Count</td></tr></table></td></tr>"
  }
  # Last-sign-in recency rows (true sign-in dates; never/unknown separate by design)
  $siRowsHtml = ""
  foreach ($b in $siRows) {
    $col = if ($b.Label -eq "Never signed in") { "#dc2626" } elseif ($b.Label -like "Unknown*") { "#94a3b8" } elseif ($b.Label -eq "<30d") { "#16a34a" } elseif ($b.Label -in @("30–90d")) { "#f59e0b" } else { "#dc2626" }
    $siRowsHtml += New-PixelBarRow -Label $b.Label -Count $b.Count -Max $maxSi -Color $col -LabelWidth 95
  }
  # Account-age rows (creation recency: separate measurement, never mixed with sign-in)
  $ageRowsHtml = ""
  foreach ($b in $ageRows) {
    $col = if ($b.Label -eq "Unknown") { "#94a3b8" } elseif ($b.Label -in @("<30d", "30–90d")) { "#2563eb" } else { "#7c3aed" }
    $ageRowsHtml += New-PixelBarRow -Label $b.Label -Count $b.Count -Max $maxAge -Color $col -LabelWidth 60
  }
  # Top invited domains table rows (no risk indicators)
  $domainTableRowsHtml = ""
  foreach ($d in $domainTable) {
    $tcol = if ($d.Type -eq "Consumer") { "#d97706" } else { "#2563eb" }
    $domainTableRowsHtml += "<tr style='border-bottom:1px solid #f1f5f9;'><td style='padding:10px 8px;'><strong>$($d.Domain)</strong></td><td style='padding:10px 8px;text-align:center;'><span style='background:#f1f5f9;color:$tcol;border-radius:12px;padding:3px 10px;font-size:11px;font-weight:600;'>$($d.Type)</span></td><td style='padding:10px 8px;text-align:center;font-weight:600;'>$($d.Guests)</td><td style='padding:10px 8px;text-align:center;'>$($d.Active)</td><td style='padding:10px 8px;text-align:center;'>$($d.StaleNever)</td></tr>"
  }
  if (-not $domainTableRowsHtml) { $domainTableRowsHtml = "<tr><td colspan='5' style='padding:20px;text-align:center;color:#64748b;'>No domain data.</td></tr>" }

  $modeLine = if ($IsDryRun) {
    if (($whatif + $whatifDel) -gt 0) { "DRY RUN (SIMULATION) — $whatif guests eligible for disable, $whatifDel eligible for delete. No accounts modified." }
    else { "DRY RUN (SIMULATION) — No accounts meet cleanup thresholds this run. No accounts modified." }
  } elseif ($AutoDelete -and ($justDeleted + $delFailed) -gt 0) {
    "LIFECYCLE ARMED — $justDisabled disabled, $justDeleted deleted this run."
  } elseif ($AutoDisable -and ($justDisabled + $failed) -gt 0) {
    "CLEANUP ARMED — $justDisabled disabled this run (deletes off)."
  } elseif ($AutoDisable -or $AutoDelete) {
    "ARMED — no guests met the thresholds this run; nothing changed."
  } else {
    "REPORT-ONLY — no guest accounts were changed."
  }

  $modeBadge = if ($IsDryRun) {
    '<span style="background:#fef3c7;color:#b45309;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600;">Dry Run (Simulation)</span>'
  } elseif ($AutoDelete) {
    '<span style="background:#fee2e2;color:#dc2626;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600;">Disable + Delete armed</span>'
  } elseif ($AutoDisable) {
    '<span style="background:#ffedd5;color:#ea580c;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600;">Cleanup armed (disable only)</span>'
  } else {
    '<span style="background:#dcfce7;color:#16a34a;border-radius:20px;padding:4px 12px;font-size:12px;font-weight:600;">Report-only</span>'
  }
  $activePct = (& $pct $active); $dormantPct = (& $pct $dormant); $stalePct = (& $pct ($stale + $never))
  $distTotal = [math]::Max(1, ($active + $dormant + $stale + $never))
  $distW = 580
  $distA = [math]::Round($active / $distTotal * $distW)
  $distD = [math]::Round($dormant / $distTotal * $distW)
  $distS = $distW - $distA - $distD
  $distBarHtml = "<table width='$distW' cellpadding='0' cellspacing='0' style='background:#e2e8f0;border-radius:7px;'><tr><td width='$distA' bgcolor='#16a34a' height='12' style='line-height:0;font-size:0;'>&nbsp;</td><td width='$distD' bgcolor='#f59e0b' height='12' style='line-height:0;font-size:0;'>&nbsp;</td><td width='$distS' bgcolor='#dc2626' height='12' style='line-height:0;font-size:0;'>&nbsp;</td></tr></table>"
  $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light dark">
<meta name="supported-color-schemes" content="light dark">
<title>$Org Guest User Dashboard</title>
</head>
<body style="font-family:'Segoe UI',Arial,sans-serif; background:#f8fafc; padding:16px; line-height:1.5; color:#0f172a; margin:0;">
<!--[if mso]>
<table align="center" width="880" cellpadding="0" cellspacing="0" border="0"><tr><td>
<![endif]-->
<div style="max-width:880px; margin:0 auto;">

<!-- Header -->
<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; margin-bottom:20px; border:1px solid #e2e8f0; border-top:4px solid #2563eb;">
    <tr><td style="padding:20px;"><h1 style="margin:0; font-size:22px; font-weight:600; color:#1e293b;">$Org Guest Governance Report</h1><p style="margin:4px 0 0 0; color:#64748b; font-size:12px;">Data as of: $(Get-Date -Format 'dd-MMM-yyyy HH:mm') IST • Mode: $modeLine</p></td></tr>
</table>

<!-- SPO OTP notice -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#fff7ed; border-left:4px solid #ea580c; border-radius:4px;">
    <tr><td style="padding:12px; font-size:12px; color:#7c2d12;"><strong>⚠️ SPO OTP retirement (mid-Oct 2026 – end-Nov 2026):</strong> external users without a B2B guest account get access denied on old specific-people links. Lifecycle: stale ≥ ${StaleThresholdDays}d → <strong>disable</strong>, then <strong>delete ${DeleteAfterDisabledDays}d after disable</strong>. Resharing one file re-creates the guest if access must be restored. CSVs attached: full inventory + stale-only list. State file tracks grace + delete clocks across runs.</td></tr>
</table>

<!-- Executive summary -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#eff6ff; border-left:4px solid #2563eb; border-radius:4px;">
    <tr><td style="padding:16px 20px;">
        <div style="font-size:15px; font-weight:700; color:#1e3a8a; margin-bottom:6px;">Executive Summary</div>
        <div style="font-size:13px; color:#1e40af; line-height:1.7;"><strong>$total</strong> Total Guests • <strong>$($stale + $never)</strong> ($(& $pct ($stale + $never))%) Stale / Never Signed In • <strong>$recentInvitesC</strong> ($recentPct%) Recently Invited (&le; 7d) • <strong>$pendingAccept</strong> Pending Invitations • <strong>$recDisable</strong> Disable Candidates • <strong>$recKeep</strong> Active / In Use. Recommended focus: review the <strong>$($stale + $never) stale & never-signed-in guests</strong> before enabling automated lifecycle cleanup.</div>
    </td></tr>
</table>

<!-- KPI cards: one per status, never mixed -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr>
        <td width="20%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;"><tr><td style="padding:20px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:8px;">TOTAL GUESTS</div>
                <div style="font-size:32px; font-weight:700; color:#0f172a;">$total</div>
                <div style="font-size:11px; color:#64748b;">enabled: $enabledC | disabled: $disabledC</div>
            </td></tr></table>
        </td>
        <td width="20%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;"><tr><td style="padding:20px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:8px;">ACTIVE</div>
                <div style="font-size:32px; font-weight:700; color:#16a34a;">$active</div>
                <div style="font-size:11px; color:#64748b;">$activePct% • signed in &lt; ${DormantThresholdDays}d</div>
            </td></tr></table>
        </td>
        <td width="20%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;"><tr><td style="padding:20px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:8px;">DORMANT</div>
                <div style="font-size:32px; font-weight:700; color:#d97706;">$dormant</div>
                <div style="font-size:11px; color:#64748b;">$(& $pct $dormant)% • ${DormantThresholdDays}–${StaleThresholdDays}d inactive</div>
            </td></tr></table>
        </td>
        <td width="20%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;"><tr><td style="padding:20px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:8px;">STALE</div>
                <div style="font-size:32px; font-weight:700; color:#ea580c;">$stale</div>
                <div style="font-size:11px; color:#64748b;">$(& $pct $stale)% • ≥ ${StaleThresholdDays}d inactive</div>
            </td></tr></table>
        </td>
        <td width="20%" style="vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;"><tr><td style="padding:20px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:8px;">NEVER SIGNED IN</div>
                <div style="font-size:32px; font-weight:700; color:#dc2626;">$never</div>
                <div style="font-size:11px; color:#64748b;">$(& $pct $never)% • invited, zero sign-ins</div>
            </td></tr></table>
        </td>
    </tr>
</table>

<!-- Governance KPI row: Recently Invited, Pending, Stale/Never, Disabled -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr>
        <td width="25%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #2563eb;"><tr><td style="padding:16px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">RECENTLY INVITED (&le;7d)</div>
                <div style="font-size:28px; font-weight:700; color:#2563eb;">$recentInvitesC</div>
                <div style="font-size:11px; color:#64748b;">$recentPct% of guests</div>
            </td></tr></table>
        </td>
        <td width="25%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #7c3aed;"><tr><td style="padding:16px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">PENDING INVITES</div>
                <div style="font-size:28px; font-weight:700; color:#7c3aed;">$pendingAccept</div>
                <div style="font-size:11px; color:#64748b;">never accepted • acceptance $acceptRate%</div>
            </td></tr></table>
        </td>
        <td width="25%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #ea580c;"><tr><td style="padding:16px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">STALE & NEVER SIGNED IN</div>
                <div style="font-size:28px; font-weight:700; color:#ea580c;">$($stale + $never)</div>
                <div style="font-size:11px; color:#64748b;">$stalePct% • cleanup candidates</div>
            </td></tr></table>
        </td>
        <td width="25%" style="vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #dc2626;"><tr><td style="padding:16px; text-align:center;">
                <div style="font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">DISABLED ACCOUNTS</div>
                <div style="font-size:28px; font-weight:700; color:#dc2626;">$disabledC</div>
                <div style="font-size:11px; color:#64748b;">held in Stage 1 / Stage 2 delete clock</div>
            </td></tr></table>
        </td>
    </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
    <tr><td style="padding:24px;">
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr><td style="font-size:16px; font-weight:600; color:#0f172a; padding-bottom:4px;" colspan="2">Guest Hygiene $modeBadge</td></tr>
            <tr><td style="font-size:12px; color:#64748b; padding-bottom:12px;" colspan="2">Stale = no sign-in ≥ ${StaleThresholdDays}d • Lifecycle this run: disabled $justDisabled, deleted $justDeleted, pending grace $pending, on hold $onHold</td></tr>
            <tr><td colspan="2">
                $distBarHtml
            </td></tr>
            <tr><td style="font-size:12px; color:#475569; padding-top:8px;" colspan="2">Active $active ($activePct%) &nbsp;&nbsp; Dormant $dormant ($dormantPct%) &nbsp;&nbsp; Stale+Never $($stale + $never) ($stalePct%)</td></tr>
            <tr><td style="font-size:12px; color:#475569; padding-top:8px; border-top:1px solid #f1f5f9;" colspan="2">Data confidence: <strong>$confidence</strong> • Sign-in date present: $withSignInDate • Never signed in: $neverKnown • No dates (Unknown): $unknownC • Privileged group membership detected: $privGuests (direct memberships only — not Entra roles, PIM, or SharePoint/Teams permissions; see Notes) • In Teams: $teamGuests • No inviter record: $noInviter</td></tr>
        </table>
    </td></tr>
</table>

<!-- Last sign-in recency + Account age: two separate measurements, never mixed -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr>
        <td width="50%" style="padding-right:12px; vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
                <tr><td style="padding:20px 20px 8px 20px;">
                    <div style="font-size:15px; font-weight:600; color:#0f172a;">Last Sign-in Activity</div>
                    <div style="font-size:12px; color:#64748b; margin-bottom:8px;">True sign-in recency from Entra (never/unknown shown separately — not folded into buckets)</div>
                    <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px; border-collapse:collapse;">$siRowsHtml</table>
                </td></tr>
            </table>
        </td>
        <td width="50%" style="vertical-align:top;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
                <tr><td style="padding:20px 20px 8px 20px;">
                    <div style="font-size:15px; font-weight:600; color:#0f172a;">Account Age</div>
                    <div style="font-size:12px; color:#64748b; margin-bottom:8px;">Days since guest account creation (lifecycle input, not activity)</div>
                    <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px; border-collapse:collapse;">$ageRowsHtml</table>
                </td></tr>
            </table>
        </td>
    </tr>
</table>

<!-- Creation trend + trend KPIs fill the row -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
    <tr>
        <td width="62%" style="padding:20px 10px 8px 20px; vertical-align:top;">
            <div style="font-size:15px; font-weight:600; color:#0f172a;">Guest Creation Trend — Last 12 Months</div>
            <div style="font-size:12px; color:#64748b; margin-bottom:8px;">New B2B guest accounts per month • Invitation acceptance overall: $acceptedC accepted ($acceptRate%) / $pendingAccept pending • Oct 26 is month-to-date (report of $(Get-Date -Format 'dd MMM'))</div>
            <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px; border-collapse:collapse;">$trendRowsHtml</table>
            <div style="font-size:11px; color:#92400e; background:#fffbeb; border:1px solid #fde68a; border-radius:6px; padding:8px 12px; margin-top:10px;">Jun–Sep 2026: <strong>$junSepSum</strong> guest accounts created — the dominant creation period in the last 12 months, ahead of the mid-Oct 2026 SPO OTP retirement. Oct 2026 MTD: <strong>$octMTD</strong>.</div>
        </td>
        <td width="38%" style="padding:20px 20px 8px 10px; vertical-align:top;">
            <div style="font-size:13px; font-weight:600; color:#0f172a; margin-bottom:10px;">Trend KPIs</div>
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="50%" style="padding:0px 6px 12px 0px; vertical-align:top;">
                        <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc; border-radius:10px; border:1px solid #e2e8f0;"><tr><td style="padding:14px; text-align:center;">
                            <div style="font-size:10px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Peak Month</div>
                            <div style="font-size:24px; font-weight:700; color:#16a34a;">$peakCount</div>
                            <div style="font-size:11px; color:#64748b;">$peakMonth</div>
                        </td></tr></table>
                    </td>
                    <td width="50%" style="padding:0px 0px 12px 6px; vertical-align:top;">
                        <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc; border-radius:10px; border:1px solid #e2e8f0;"><tr><td style="padding:14px; text-align:center;">
                            <div style="font-size:10px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Jun–Sep Created</div>
                            <div style="font-size:24px; font-weight:700; color:#2563eb;">$junSepSum</div>
                            <div style="font-size:11px; color:#64748b;">pre-retirement bulk</div>
                        </td></tr></table>
                    </td>
                </tr>
                <tr>
                    <td width="50%" style="padding:0px 6px 0px 0px; vertical-align:top;">
                        <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc; border-radius:10px; border:1px solid #e2e8f0;"><tr><td style="padding:14px; text-align:center;">
                            <div style="font-size:10px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Oct MTD</div>
                            <div style="font-size:24px; font-weight:700; color:#d97706;">$octMTD</div>
                            <div style="font-size:11px; color:#64748b;">month to date</div>
                        </td></tr></table>
                    </td>
                    <td width="50%" style="padding:0px 0px 0px 6px; vertical-align:top;">
                        <table width="100%" cellpadding="0" cellspacing="0" style="background:#f8fafc; border-radius:10px; border:1px solid #e2e8f0;"><tr><td style="padding:14px; text-align:center;">
                            <div style="font-size:10px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Acceptance</div>
                            <div style="font-size:24px; font-weight:700; color:#7c3aed;">$acceptRate%</div>
                            <div style="font-size:11px; color:#64748b;">$acceptedC / $total accepted</div>
                        </td></tr></table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>

<!-- Top invited domains -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
    <tr><td style="padding:24px;">
        <div style="font-size:16px; font-weight:600; color:#0f172a;">Top External Domains — Where Guest Accounts Originate</div>
        <div style="font-size:12px; color:#64748b; margin-bottom:12px;">Top sending domains by guest count. Consumer domains (Gmail etc.) deserve extra scrutiny for business use.</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px; border-collapse:collapse; background:#fafafa; border-radius:8px;">
            <tr style="background:#f1f5f9;">
                <td style="padding:10px 8px; font-weight:600;">Domain</td>
                <td style="padding:10px 8px; font-weight:600; text-align:center;">Type</td>
                <td style="padding:10px 8px; font-weight:600; text-align:center;">Guests</td>
                <td style="padding:10px 8px; font-weight:600; text-align:center;">Active/Dormant</td>
                <td style="padding:10px 8px; font-weight:600; text-align:center;">Stale/Never</td>
            </tr>
            $domainTableRowsHtml
        </table>
    </td></tr>
</table>

<!-- Recommended actions -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
    <tr><td style="padding:24px;">
        <div style="font-size:16px; font-weight:600; color:#0f172a;">Recommended Actions</div>
        <div style="font-size:12px; color:#64748b; margin-bottom:12px;">What should happen to each guest, per the lifecycle rules. Full per-guest list is in the attached CSV.</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px; border-collapse:collapse; background:#fafafa; border-radius:8px;">
            <tr style="background:#f1f5f9;"><td style="padding:10px 12px; font-weight:600;">Action</td><td style="padding:10px 12px; font-weight:600; text-align:right;">Guests</td></tr>
            <tr style="border-bottom:1px solid #f1f5f9;"><td style="padding:10px 12px;">Disable after review</td><td style="padding:10px 12px; text-align:right; font-weight:700; color:#dc2626;">$recDisable</td></tr>
            <tr style="border-bottom:1px solid #f1f5f9;"><td style="padding:10px 12px;">Owner confirmation</td><td style="padding:10px 12px; text-align:right; font-weight:700; color:#d97706;">$recOwner</td></tr>
            <tr style="border-bottom:1px solid #f1f5f9;"><td style="padding:10px 12px;">Approve (add to group)</td><td style="padding:10px 12px; text-align:right; font-weight:700; color:#d97706;">$recApprove</td></tr>
            <tr style="border-bottom:1px solid #f1f5f9;"><td style="padding:10px 12px;">Keep</td><td style="padding:10px 12px; text-align:right; font-weight:700; color:#16a34a;">$recKeep</td></tr>
            <tr><td style="padding:10px 12px;">Awaiting acceptance</td><td style="padding:10px 12px; text-align:right; font-weight:700; color:#64748b;">$recAwait</td></tr>
        </table>
    </td></tr>
</table>

<!-- Recently invited guests (last 7 days) -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
    <tr><td style="padding:24px;">
        <div style="font-size:16px; font-weight:600; color:#0f172a;">Recently Invited Guests (Last 7 Days)</div>
        <div style="font-size:12px; color:#64748b; margin-bottom:12px;">Guests created or invited in the last 7 days ($recentInvitesC total). Full inventory is in the attached CSV.</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="font-size:11px; border-collapse:collapse; background:#fafafa; border-radius:8px;">
            $queueHead
            $topRecentRowsHtml
        </table>
    </td></tr>
</table>

<!-- Top 10 guests requiring action -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0;">
    <tr><td style="padding:24px;">
        <div style="font-size:16px; font-weight:600; color:#0f172a;">Top 10 Guests Requiring Action</div>
        <div style="font-size:12px; color:#64748b; margin-bottom:12px;">Highest-priority decision items requiring administrator or owner review. View the full guest inventory in the attached CSV.</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="font-size:11px; border-collapse:collapse; background:#fafafa; border-radius:8px;">
            $queueHead
            $topActionRowsHtml
        </table>
    </td></tr>
</table>

<!-- Notes & Governance Reference -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px; background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;">
    <tr><td style="padding:20px;">
        <div style="font-size:15px; font-weight:600; color:#0f172a; margin-bottom:4px;">📝 Governance & Lifecycle Policy</div>
        <div style="font-size:12px; color:#64748b; margin-bottom:12px;">Thresholds: stale ≥ ${StaleThresholdDays}d inactive • dormant ≥ ${DormantThresholdDays}d • delete ${DeleteAfterDisabledDays}d after disable • grace ${GracePeriodDays}d with $RequireRuns sightings.</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px; border-collapse:collapse; color:#334155;">
            <tr style="border-bottom:1px solid #f1f5f9;"><td style="padding:8px 4px; width:28px; vertical-align:top;">1️⃣</td><td style="padding:8px 4px;"><strong>SPO OTP Context:</strong> External users without a B2B guest account lose access to shared links post-retirement. This periodic review governs that external guest population.</td></tr>
            <tr style="border-bottom:1px solid #f1f5f9;"><td style="padding:8px 4px; vertical-align:top;">2️⃣</td><td style="padding:8px 4px;"><strong>2-Stage Safety Lifecycle:</strong> Stale guests are disabled first (requiring $RequireRuns consecutive sightings + ${GracePeriodDays}d grace). Disabled accounts are soft-deleted after ${DeleteAfterDisabledDays}d (holding a 30-day recovery window in Entra recycle bin).</td></tr>
            <tr><td style="padding:8px 4px; vertical-align:top;">3️⃣</td><td style="padding:8px 4px;"><strong>Methodology:</strong> Sign-in source: $(if ($signInOk) { 'Entra signInActivity' } else { 'account-age proxy' }). Direct privileged group detection identifies elevated access. Detailed per-guest analysis is available in the attached CSV and standalone HTML files.</td></tr>
        </table>
    </td></tr>
</table>

<!-- Administrator Quick-Guide -->
<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px; background:#f8fafc; border-radius:12px; border:1px solid #e2e8f0;">
    <tr><td style="padding:16px 20px;">
        <div style="font-size:13px; font-weight:600; color:#0f172a; margin-bottom:4px;">Administrator Quick-Guide</div>
        <div style="font-size:11px; color:#475569; line-height:1.6;">Current mode: <strong>$modeLine</strong>. Attached files: <strong>Full Guest Inventory CSV</strong>, <strong>Stale-Only CSV</strong>, and standalone <strong>HTML Dashboard</strong>. To inspect full multi-column tables with interactive browser sorting, open the attached HTML dashboard file.</div>
    </td></tr>
</table>

<!-- Footer -->
<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:12px; border:1px solid #e2e8f0;">
    <tr><td style="padding:14px; text-align:center; font-size:11px; color:#94a3b8;">$Org Guest Governance Report v$ScriptVersion • Generated $(Get-Date -Format 'dd-MMM-yyyy HH:mm') IST • Mode: $modeLine</td></tr>
</table>

</div>
<!--[if mso]>
</td></tr></table>
<![endif]-->
</body>
</html>
"@
  $html | Out-File -FilePath $htmlPath -Encoding UTF8 -WhatIf:$false
  Write-Log "Dashboard HTML: $htmlName" -Level "SUCCESS"

  # ----- Email via Graph -----
  if ($SenderEmail -and $RecipientEmail) {
    Write-Log "Sending email via Graph ($SenderEmail -> $RecipientEmail)..."
    $subject = $SubjectTemplate.Replace("{Date}", (Get-Date -Format "yyyy-MM-dd"))
    if ($IsDryRun) { $subject = "[DRY RUN] $subject" }
    $subject += " | Total: $total | Stale+Never: $($stale + $never) | Recent(<7d): $recentInvitesC | DisabledNow: $justDisabled | DeletedNow: $justDeleted | Hold: $onHold | Pending: $pending"
    if ($IsDryRun -and ($whatif + $whatifDel) -gt 0) {
      $subject += " | DryRunEligible: (Disable=$whatif, Delete=$whatifDel)"
    }
    $toList = @($RecipientEmail -split "[;,]" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $recips = @($toList | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
    $atts = @()
    foreach ($f in @($csvPath, $stalePath, $htmlPath)) {
      $bytes = [System.IO.File]::ReadAllBytes($f)
      $ct = if ($f -like "*.html") { "text/html" } else { "text/csv" }
      $atts += @{ "@odata.type" = "#microsoft.graph.fileAttachment"; name = (Split-Path $f -Leaf); contentType = $ct; contentBytes = [Convert]::ToBase64String($bytes) }
    }
    $msg = @{ message = @{ subject = $subject; body = @{ contentType = "HTML"; content = $html }; toRecipients = $recips; attachments = $atts } }
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$SenderEmail/sendMail" -Method Post -Headers (@{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }) -Body ($msg | ConvertTo-Json -Depth 10) | Out-Null
    Write-Log "Email sent." -Level "SUCCESS"
  } else { Write-Log "Sender/Recipient email missing — skipping email." -Level "WARNING" }

  Write-Host ""
  if ($IsDryRun) {
    Write-Host "[DRY RUN SUMMARY] Simulation only - no changes made to any accounts or state file." -ForegroundColor Yellow
    Write-Host "DryRun Eligibility: would-disable=$whatif | would-delete=$whatifDel | pending-grace=$pending | on-hold=$onHold" -ForegroundColor Yellow
  }
  Write-Host "Total: $total | Active: $active | Dormant: $dormant | Stale: $stale | Never: $never | Disabled-already: $staleDis" -ForegroundColor White
  Write-Host "Cleanup: disabled-now=$justDisabled deleted-now=$justDeleted hold=$onHold pending=$pending whatif=$($whatif + $whatifDel) failed=$($failed + $delFailed) (DryRun=$IsDryRun AutoDisable=$AutoDisable AutoDelete=$AutoDelete DeleteAfter=${DeleteAfterDisabledDays}d)" -ForegroundColor Cyan
  Write-Host "Invites: recent(<7d)=$recentInvitesC pending=$pendingAccept accepted=$acceptedC | NoInviter=$noInviter | PrivGrp=$privGuests | InTeams=$teamGuests" -ForegroundColor Yellow
  Write-Log "Done. CSV=$csvName STALE=$staleName HTML=$htmlName" -Level "SUCCESS"
}
catch {
  Write-Log "FATAL: $_" -Level "ERROR"
  Write-Log ($_.ScriptStackTrace) -Level "ERROR"
  exit 1
}

