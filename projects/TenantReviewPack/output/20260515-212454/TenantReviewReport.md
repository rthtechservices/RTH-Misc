# Tenant Review Report

Tenant: rthtechservices

Review period: 2026

Generated: 2026-05-15T21:25:40.5516872-07:00

## Executive Summary

### Tenant configured with verified domains

RTH Tech Services Inc. has 5 verified domains configured, with the default domain set to rthtechservices.com. The tenant was created in January 2021 and has 69 assigned service plans.

Business impact: The tenant is properly established and operational with multiple domains available for business use.

Recommended action: Periodically review domains and technical notification contacts to ensure they remain current and aligned to business needs.

### Very high license quantity with minimal assignments

There are 1,050,005 total licenses recorded, but only 12 are assigned to users. Estimated annual cost is 480 CAD, with several SKUs missing pricing data.

Business impact: Large discrepancies between purchased and assigned counts may indicate reporting anomalies or misaligned license records.

Recommended action: Validate license quantities and SKU records with your licensing provider to ensure counts and billing are accurate.

### Most users unlicensed and no sign-in data collected

Out of 15 total users (10 members and 5 guests), only 2 are licensed. Sign-in activity data has not been requested or collected.

Business impact: Lack of licensing and sign-in visibility limits insight into actual usage and potential security risks.

Recommended action: Enable sign-in activity reporting and review which users require licenses based on business roles.

### Significant unused license count identified

There are 1,049,993 unused licenses recorded, with one item flagged for attention. No licensed users are disabled, stale, or guests.

Business impact: Excess unused licenses may create administrative complexity and indicate inaccurate license tracking.

Recommended action: Reconcile license inventory to confirm actual entitlements and remove or correct unnecessary allocations.

### Mailbox forwarding enabled and one mailbox dominates storage

There are 9 mailboxes, including 4 shared mailboxes and 1 user mailbox. Two mailboxes have forwarding enabled, and the shared mailbox 'Escala' accounts for 186,719,050.60 GB of the reported 186,719,055.74 GB total storage.

Business impact: Mailbox forwarding can increase data exposure risk, and extreme mailbox size values may indicate reporting or configuration issues.

Recommended action: Review forwarding configurations for business justification and validate mailbox size reporting for accuracy.

### SharePoint lightly used with no external sharing

There are 38 SharePoint sites using 22.12 GB of storage out of a 972,800 GB quota. No sites have external sharing enabled.

Business impact: Storage utilization is low and external sharing risk is minimized under current settings.

Recommended action: Continue monitoring storage growth and enable external sharing only where a defined business need exists.

### Teams deployed but no recent activity

There are 3 private Teams, with 2 marked inactive and no recorded channel messages or meetings in the last 90 days.

Business impact: Low Teams activity may indicate underutilization of collaboration capabilities.

Recommended action: Assess whether Teams should be further adopted for collaboration or if unused Teams should be cleaned up.

### Multiple stale devices and no compliance visibility

Nine devices are registered, with 5 marked as stale and 2 disabled. No devices are managed by Intune, and compliance status is unknown for all devices.

Business impact: Unmanaged and stale devices reduce visibility and may increase security risk.

Recommended action: Enable Intune management and review stale or disabled devices for removal or remediation.

### Single Copilot license fully assigned

One Copilot license has been purchased and assigned, with an estimated monthly cost of 40 CAD. Usage reporting has not been requested.

Business impact: Without usage reporting, it is unclear whether the Copilot investment is delivering value.

Recommended action: Enable Copilot usage reporting to evaluate adoption and return on investment.

## Big Numbers

| Area | Metric | Value |
| --- | --- | ---: |
| Licenses | Purchased | 1050005 |
| Licenses | Unused | 1049993 |
| Users | Total users | 15 |
| Users | Licensed disabled users | 0 |
| Mail | External forwarding suspected | 0 |
| SharePoint | Total sites | 38 |
| Teams | Inactive teams | 2 |
| Devices | Stale devices | 5 |
| Copilot | Unused Copilot licenses | 0 |

## Attention Items

- **Medium**: 1049993 purchased licenses appear unused - Review upcoming hiring, project, and retention needs before reducing or reallocating licenses.

## Dataset Notes

### TenantOverview

- No collector warnings.
- JSON: `TenantOverview.json`

### LicenseInventory

- No collector warnings.
- JSON: `LicenseInventory.json`

### UserInventory

- No collector warnings.
- JSON: `UserInventory.json`

### LicenseUserAnalysis

- No collector warnings.
- JSON: `LicenseUserAnalysis.json`

### MailboxInventory

- No collector warnings.
- JSON: `MailboxInventory.json`

### SharePoint

- No collector warnings.
- JSON: `SharePoint.json`

### Teams

- No collector warnings.
- JSON: `Teams.json`

### Devices

- No collector warnings.
- JSON: `Devices.json`

### Copilot

- No collector warnings.
- JSON: `Copilot.json`
