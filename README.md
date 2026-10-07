# M365 & Entra ID Guest Governance Dashboard & 2-Stage Lifecycle

An automated, enterprise-grade PowerShell governance solution for Microsoft Entra ID (Azure AD) and Microsoft 365 B2B guest accounts.

This tool inventories all external guest users, tracks inviter and group access relationships, monitors true sign-in recency via Microsoft Graph, generates an **Outlook-optimized responsive HTML dashboard** with attached CSV reports, and safely manages account deprovisioning through a **2-stage lifecycle cleanup** engine.

---

## Key Features

- **Executive Governance Reporting:**
  - Automated inventory of all Entra ID B2B guest users.
  - Tracking of invited-by relationships, invitation acceptance rates, and external domain categorization (Corporate vs Consumer).
  - Last sign-in activity analysis via Graph `signInActivity` (with account-age fallback).
  - Identification of recently invited guests ($\le 7$ days) and pending invitations.
  - Segregation of formally approved vs non-approved guest accounts.
- **Outlook-Optimized HTML Dashboard:**
  - Responsive, centered 880px design with pixel-bar charts that render reliably in Microsoft Outlook desktop and web clients.
  - Interactive standalone HTML and CSV inventory attachments sent automatically via Microsoft Graph Mail API.
- **Enterprise 2-Stage Lifecycle Engine:**
  - **Stage 1 (Disable):** Inactive guests ($\ge 90$ days inactive default) are automatically disabled only after meeting multi-run observation criteria and grace period thresholds.
  - **Stage 2 (Soft-Delete):** Disabled accounts are queued for soft deletion after a configurable retention window (e.g. 90 days), preserving a 30-day restore period in the Entra ID recycle bin.
- **Fail-Safe Protections:**
  - **DryRun Simulation:** Enabled by default in configuration. Evaluates criteria without making changes or mutating state files.
  - **State Tracking:** Tracks sightings, grace periods, and deletion clocks in `GuestCleanupState.json`.
  - **Protected Scope:** Built-in safeguards prevent touching internal member accounts or privileged directories.

---

## Architecture & Lifecycle Workflow

```
[ Active Guest User ]
       │
       │ (Inactive >= 90 days)
       ▼
[ Stale Observation ] ──► Sighted across >= 2 runs + Grace Period (7 days)
       │
       ▼
[ Stage 1: Account Disabled ]
       │
       │ (Remains disabled for >= 90 days)
       ▼
[ Stage 2: Soft-Deleted ] ──► 30-Day Entra ID Recycle Bin Recovery Window
```

---

## Prerequisites

1. **PowerShell 7+** (Recommended) or **Windows PowerShell 5.1**.
2. **Entra ID App Registration** (App-only certificate authentication):
   - Client Certificate installed in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`.
   - Microsoft Graph Application permissions:
     - `User.Read.All`
     - `User.ReadWrite.All` *(Required only if automated disable/delete is enabled)*
     - `AuditLog.Read.All` *(Required for accurate `signInActivity` recency)*
     - `Directory.Read.All`
     - `GroupMember.Read.All`
     - `Mail.Send` *(To deliver email dashboards)*

---

## Setup & Configuration

1. Clone the repository:
   ```bash
   git clone https://github.com/tusharind007/M365-GuestGovernance-Dashboard_Lifecycle.git
   cd M365-GuestGovernance-Dashboard_Lifecycle
   ```

2. Copy the configuration template:
   ```powershell
   Copy-Item config.template.json config.json
   ```

3. Update `config.json` with your tenant and application parameters:
   ```json
   {
     "OrganizationName": "YourOrganization",
     "TenantId": "YOUR-TENANT-GUID",
     "ClientId": "YOUR-APP-CLIENT-GUID",
     "CertThumbprint": "YOUR-CERTIFICATE-THUMBPRINT-HEX",
     "ReportPath": ".\\Reports",
     "DryRun": true,
     "StaleThresholdDays": 90,
     "DormantThresholdDays": 30,
     "GracePeriodDays": 7,
     "AutoDisableStaleGuests": false,
     "AutoDeleteDisabledGuests": false,
     "DeleteAfterDisabledDays": 90,
     "ApprovedGuestGroupId": "OPTIONAL-APPROVED-GROUP-GUID",
     "ApprovedGuestGroupName": "Approved_External_Guests",
     "RequireConsecutiveStaleRuns": 2,
     "SenderEmail": "admin@yourdomain.com",
     "RecipientEmail": "governance-alerts@yourdomain.com",
     "EmailSubject": "Guest User Review & Cleanup Dashboard - {Date}"
   }
   ```

---

## Usage

### 1. Report-Only / Simulation (Default)
Runs a full audit, generates CSVs + HTML dashboard, sends the email report, and previews eligible cleanup actions without modifying any accounts:
```powershell
.\M365-GuestGovernance-Dashboard.ps1
```

Or explicitly specify `-DryRun`:
```powershell
.\M365-GuestGovernance-Dashboard.ps1 -DryRun
```

### 2. Previewing Changes (-WhatIf)
Preview what actions would execute against Graph API:
```powershell
.\M365-GuestGovernance-Dashboard.ps1 -EnableCleanup -WhatIf
```

### 3. Live Stage 1 (Disable Inactive Guests)
Requires `-Live` and `-EnableCleanup` (or setting `"AutoDisableStaleGuests": true` and `"DryRun": false` in `config.json`):
```powershell
.\M365-GuestGovernance-Dashboard.ps1 -Live -EnableCleanup
```

### 4. Live Stage 2 (Full Lifecycle: Disable + Delete)
Enables both Stage 1 disable and Stage 2 soft-deletion of accounts disabled $\ge 90$ days:
```powershell
.\M365-GuestGovernance-Dashboard.ps1 -Live -EnableCleanup -EnableDelete
```

---

## License

This project is licensed under the [MIT License](LICENSE).
