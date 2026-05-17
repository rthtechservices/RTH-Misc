# Microsoft 365 Tenant Review

**Organisation:** RTH Tech Services Inc.
**Review period:** Sample review period
**Prepared by:** RTH Tech Services
**Generated:** May 16, 2026

---

## Executive Summary

> The tenant is in good standing. No major issues were identified during this review.

- This review covers the RTH Tech Services Inc. Microsoft 365 tenant. The tenant has 15 user accounts — 10 members and 5 guests.
- 12 Microsoft 365 licenses are currently assigned with an estimated monthly spend of CAD $40 (partial — 9 SKUs are missing price data). 1049993 licenses appear unassigned.
- Mail forwarding is active on 2 mailboxes, and 17 inbox forwarding rules were found. These should be confirmed as intentional.
- 2 of 3 Teams workspaces had no recorded activity in the last 90 days and may be candidates for archiving.
- 5 of 9 registered devices appear stale. Cleaning up outdated device records supports better compliance and security reporting.
- 1 Microsoft 365 Copilot license is assigned at an estimated CAD $40/month. Enable usage reporting to confirm active adoption.

## Key Metrics at a Glance

| Area | Metric | Value | Status |
| --- | --- | ---: | --- |
| Users | Total Users | 15 | Info |
| Users | Guest Users | 5 | Watch |
| Users | Stale Users | N/A | Not Available |
| Licensing | Licensed Users | 12 | Info |
| Licensing | Unused Licenses | 1,049,993 | Action Recommended |
| Licensing | Est. Monthly Spend | CAD $40.00 | Info |
| Licensing | Disabled Licensed Users | 0 | Healthy |
| Licensing | Stale Licensed Users | 0 | Healthy |
| Exchange | Total Mailboxes | 9 | Info |
| Exchange | Mailboxes with Forwarding | 2 | Watch |
| Exchange | External Forwarding | 0 | Healthy |
| Exchange | Inbox Forwarding Rules | 17 | Action Recommended |
| SharePoint | SharePoint Sites | 38 | Info |
| SharePoint | SharePoint Storage | 22.1 GB | Info |
| SharePoint | External Sharing Sites | 0 | Healthy |
| SharePoint | OneDrive Accounts | 0 | Info |
| Teams | Total Teams | 3 | Info |
| Teams | Inactive Teams | 2 | Watch |
| Teams | Teams Without Owners | 0 | Healthy |
| Devices | Total Devices | 9 | Info |
| Devices | Stale Devices | 5 | Watch |
| Devices | Intune Managed | 0 | Watch |
| Copilot | Copilot Licenses | 1 | Info |
| Copilot | Unused Copilot Licenses | 0 | Healthy |
| Copilot | Active Copilot Users | N/A | Not Available |
| Copilot | Copilot Monthly Spend | CAD $40.00 | Info |

## Tenant Overview

**Tenant configuration is stable and fully verified** · *Healthy*

RTH Tech Services Inc. has a properly configured Microsoft 365 tenant with all five domains verified and no federation in place. Core setup appears stable and centrally managed.

**Why it matters:** A properly verified and non-federated tenant reduces identity complexity and supports predictable authentication behavior.

**Recommended action:** Confirm password expiration policy and validate that the technical notification address is monitored and backed up.

## Cost & Licensing Review

**License inventory shows significant SKU count anomaly** · *Watch*

The tenant reports over one million purchased licenses but only 12 assigned, with an estimated total cost of $40 CAD per month. This discrepancy suggests reporting or SKU configuration issues rather than actual spend.

**Why it matters:** Inaccurate license counts and missing pricing data can obscure true spend and complicate budgeting or compliance reviews.

**Recommended action:** Review license subscriptions in the Microsoft 365 admin center to validate actual quantities and ensure pricing data is complete for accurate cost reporting.

## License Optimisation

**High reported unused license count requires validation** · *Watch*

Nearly all reported licenses appear unused, but no direct cost impact is calculated. This likely reflects a reporting inconsistency rather than actual financial waste.

**Why it matters:** If accurate, unused licenses could represent financial waste; if inaccurate, reporting gaps may undermine governance decisions.

**Recommended action:** Reconcile subscription counts against Microsoft billing to confirm true license quantities and eliminate reporting anomalies.

## Identity & User Review

**Most users unlicensed and no sign-in tracking enabled** · *Action Recommended*

The tenant contains 15 users, but only 2 are licensed and sign-in activity reporting has not been enabled. This limits visibility into actual usage and risk.

**Why it matters:** Limited licensing and lack of sign-in data reduce oversight of account activity and potential security exposure.

**Recommended action:** Enable Azure AD sign-in activity reporting and review whether unlicensed accounts are required or should be removed.

## Exchange & Mail Flow

**Mailbox storage heavily concentrated in one shared mailbox** · *Action Recommended*

Email usage is modest overall, but one shared mailbox holds nearly all reported storage. Two mailboxes have forwarding enabled, both routing internally.

**Why it matters:** Mailbox size anomalies and forwarding rules can mask data governance or security risks if not validated.

**Recommended action:** Validate mailbox statistics for accuracy, review the Escala shared mailbox, and assess the necessity of existing inbox forwarding rules.

## SharePoint & OneDrive

**Low SharePoint utilization with no external sharing** · *Healthy*

SharePoint usage is light, with 38 sites consuming a small portion of the available storage and no external sharing enabled. OneDrive reporting was not collected.

**Why it matters:** Low storage utilization reduces risk, but missing OneDrive visibility limits full assessment of data exposure.

**Recommended action:** Enable OneDrive reporting and periodically review site ownership and lifecycle to prevent sprawl.

## Teams Collaboration

**Minimal Teams activity in the last 90 days** · *Watch*

Microsoft Teams is deployed but shows no measurable activity in the past 90 days. All teams are private and none are archived.

**Why it matters:** Underutilized collaboration tools may indicate missed productivity opportunities or redundant configuration.

**Recommended action:** Confirm whether Teams is intended for active collaboration and archive or remove unused teams if not required.

## Devices & Endpoints

**Devices unmanaged and several marked stale** · *Action Recommended*

Nine devices are registered, but none are managed through Intune and over half are considered stale. Compliance status is unknown for all devices.

**Why it matters:** Unmanaged and stale devices increase the risk of unauthorized access and reduce control over corporate data.

**Recommended action:** Implement Intune device management and review or remove stale device registrations.

## Copilot Review

**Single Copilot license fully assigned** · *Info*

One Microsoft Copilot license is purchased and assigned, with no unused capacity. Usage reporting has not been enabled.

**Why it matters:** Without usage insights, it is unclear whether the Copilot investment is delivering value.

**Recommended action:** Enable Copilot usage reporting to measure adoption and return on investment.

## Top Recommendations

| # | Category | Recommendation | Effort | Impact | Status |
| --- | --- | --- | --- | --- | --- |
| 1 | Licensing | Review 1049993 unassigned licenses before next renewal | Low | Medium | Watch |
| 2 | Exchange | Review 17 inbox forwarding rules | Medium | Medium | Watch |
| 3 | Teams | Archive or clean up 2 inactive Teams | Low | Low | Watch |
| 4 | Devices | Clean up 5 stale device records | Low | Medium | Watch |
| 5 | Devices | Consider enrolling devices in Intune for management and compliance coverage | High | High | Watch |
| 6 | Copilot | Enable Copilot usage reporting to measure adoption | Low | Medium | Watch |
| 7 | Licensing | Add price data for 9 SKUs missing cost information | Low | Low | Watch |

## Data Coverage

This report was generated from the following data sources. Where data was unavailable due to licensing or permissions, sections are marked accordingly.

| Dataset | Status |
| --- | --- |
| TenantOverview | Collected |
| LicenseInventory | Collected |
| UserInventory | Collected |
| LicenseUserAnalysis | Collected |
| MailboxInventory | Collected |
| SharePoint | Collected |
| Teams | Collected |
| Devices | Collected |
| Copilot | Collected |

## Appendix

### License SKUs

| SKU | Purchased | Assigned | Unused | Monthly Cost |
| --- | ---: | ---: | ---: | ---: |
| POWER_BI_STANDARD | 1000001 | 1 | 1000000 | — |
| FLOW_FREE | 10000 | 2 | 9998 | — |
| CCIBOTS_PRIVPREV_VIRAL | 10000 | 1 | 9999 | — |
| POWERAPPS_VIRAL | 10000 | 1 | 9999 | — |
| Power_Pages_vTrial_for_Makers | 10000 | 1 | 9999 | — |
| POWERAPPS_DEV | 10000 | 2 | 9998 | — |
| Microsoft 365 Copilot | 1 | 1 | 0 | CAD $40 |
| O365_BUSINESS_PREMIUM | 1 | 1 | 0 | — |
| POWERAUTOMATE_ATTENDED_RPA | 1 | 1 | 0 | — |
| Microsoft_Teams_Audio_Conferencing_select_dial_out | 1 | 1 | 0 | — |

### Mailboxes with Forwarding

| Display Name | Email | Type | External? |
| --- | --- | --- | --- |
| Make a Booking with RTH Tech Services Inc. | MakeaBookingwithRTHTechServicesInc@rthtechservices.com | SchedulingMailbox | No |
| RTH Tech Services Inc. | RTHTechServicesInc1@rthtechservices.com | SchedulingMailbox | No |

### Largest SharePoint Sites

| Site | Storage Used |
| --- | ---: |
| General RTH Tech Owners | 17.46 GB |
| 98e686d9-5ca0-4802-aab4-eae6c6083b83 | 3.49 GB |
| HGLLP Owners | 0.27 GB |
| RBS LLP Owners | 0.26 GB |
| Rohan Hare | 0.22 GB |
| SharePoint Administrator | 0.18 GB |
| Global Administrator | 0.13 GB |
| Rohan Hare | 0.05 GB |
| Rohan Hare | 0.05 GB |
| NT Service\sptimerv4 | 0.01 GB |

### Teams Inventory

| Team | Visibility | Last Activity | Inactive? |
| --- | --- | --- | --- |
| General RTH Tech | Private | 2025-11-11 | Yes |
| HGLLP | Private | 2025-03-31 | Yes |
| RBS LLP | Private | — | — |

### Stale Devices

| Device | OS | Last Sign-In (days ago) | Enabled |
| --- | --- | ---: | --- |
| samsungSM-S928W | Android | 690 | No |
| samsungSM-S908W | Android | 213 | Yes |
| RTH-DELLXPS | Windows | 96 | Yes |
| RTH-DELL-01 | Windows | 371 | No |
| RTH-PREDATOR | Windows | 110 | Yes |

### Copilot Licensed Users

| User | Stale? |
| --- | --- |
| rohan@rthtechservices.com | No |

