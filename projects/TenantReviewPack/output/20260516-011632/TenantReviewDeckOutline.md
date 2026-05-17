# Microsoft 365 Tenant Review — Deck Outline

| | |
| --- | --- |
| **Organisation** | RTH Tech Services Inc. |
| **Review Period** | 2026 |
| **Prepared by** | RTH Tech Services |
| **Generated** | May 16, 2026 |

> This file is a structured slide-by-slide deck outline. Use it to build a presentation in PowerPoint, Google Slides, or any compatible tool.
> Each slide includes a suggested visual treatment, a key message, supporting bullets, a business takeaway, and speaker notes.

---

## Slide 1 — Cover

**Visual treatment:** Full-bleed gradient background (#5a2f61 → #4fa7db). Client logo (placeholder) centred. Review period and "Microsoft 365 Tenant Review" in white text. RTH Tech Services branding bottom right.

> Microsoft 365 Tenant Review — RTH Tech Services Inc. — 2026

**Business takeaway:**
This is a professional, data-driven view of the RTH Tech Services Inc. Microsoft 365 environment.

**Speaker notes:**
Welcome the client. Introduce this as their Microsoft 365 Quarterly Business Review.
Review period: 2026.
Prepared by RTH Tech Services. This deck is based on data collected from live tenant systems.
All findings are as at the collection date shown on the report.

---

## Slide 2 — Executive Snapshot

**Visual treatment:** Single-column card layout on a light background. Status pill at top right (green/amber/red). Bullet list with icon markers.

> The tenant is in good standing. No major issues were identified during this review.

- This review covers the RTH Tech Services Inc. Microsoft 365 tenant. The tenant has 15 user accounts — 10 members and 5 guests.
- 12 Microsoft 365 licenses are currently assigned with an estimated monthly spend of CAD $40 (partial — 9 SKUs are missing price data). 1049993 licenses appear unassigned.
- Mail forwarding is active on 2 mailboxes, and 17 inbox forwarding rules were found. These should be confirmed as intentional.
- 2 of 3 Teams workspaces had no recorded activity in the last 90 days and may be candidates for archiving.
- 5 of 9 registered devices appear stale. Cleaning up outdated device records supports better compliance and security reporting.
- 1 Microsoft 365 Copilot license is assigned at an estimated CAD $40/month. Enable usage reporting to confirm active adoption.

**Business takeaway:**
The tenant is broadly operational. Priority items are highlighted in this review.

**Speaker notes:**
High-level tenant health statement: The tenant is in good standing. No major issues were identified during this review.
Walk the client through each bullet, pausing for questions.
Position this as the 'short story' — the rest of the deck is the evidence.
Reinforce that findings are based on live data, not estimates.

---

## Slide 3 — Big Numbers Dashboard

**Visual treatment:** KPI card grid (3×3 or 3×4). Each card: large number, short label, status badge (green/amber/red). Cards grouped by area.

> Key numbers at a glance across all review areas

- **15** — Total Users [Informational]
- **5** — Guest Users [Watch]
- **N/A** — Stale Users [NotAvailable]
- **12** — Licensed Users [Informational]
- **1,049,993** — Unused Licenses [ActionRecommended]
- **CAD $40.00** — Est. Monthly Spend [Informational]
- **0** — Disabled Licensed Users [Good]
- **0** — Stale Licensed Users [Good]
- **9** — Total Mailboxes [Informational]

**Business takeaway:**
These headline numbers frame the detailed findings that follow.

**Speaker notes:**
These numbers come directly from the tenant data collected during this review.
Highlight any red or amber cards as items to discuss in later slides.
Green items are healthy — acknowledge them to balance the conversation.
Remind the client that some metrics (especially Copilot) require Entra P1/P2 or specific permissions to collect fully.

---

## Slide 4 — Licensing & Cost Opportunities

**Visual treatment:** Split layout: left side big numbers (SKUs, total purchased, assigned, unused). Right side: mini bar representing utilisation. Amber/red callouts for disabled and stale licensed users.

> N/A unused license units across 10 SKUs

- Total purchased units: **N/A** | Assigned: **N/A** | Unused: **N/A**
- Active SKUs: 10 | Estimated monthly cost: **40**
- Disabled users with active licenses: **0** → unnecessary spend
- Stale (inactive) licensed users: **0** → potential cleanup
- Guest users with licenses: **0** → confirm if intentional
- Action: Review license subscriptions in the Microsoft 365 admin center to validate actual quantities and ensure pricing data is complete for accurate cost reporting.

**Business takeaway:**
License optimisation is one of the highest-ROI improvements available. Even modest cleanup can reduce monthly M365 spend.

**Speaker notes:**
Walk through the license SKU breakdown from the report appendix.
0 disabled users with active licenses represent avoidable cost.
0 stale licensed users should be reviewed before the next renewal or monthly true-up.
Not all SKUs have price data — work with the client to confirm exact pricing from their agreement.
Action: Reconcile subscription counts against Microsoft billing to confirm true license quantities and eliminate reporting anomalies.

---

## Slide 5 — Users & Access Review

**Visual treatment:** Donut chart: Members vs Guests. Three highlight cards below: Total users, Active, Guests. Status badges.

> 15 users in directory — 10 members, 5 guests

- Total directory users: **15** (10 members, 5 guests)
- Disabled licensed users: **0** (confirm these are intentionally retained)
- Stale licensed users: **0** (no sign-in activity in review period)
- Guest users with licenses: **0** (confirm business justification)
- Action: Enable Azure AD sign-in activity reporting and review whether unlicensed accounts are required or should be removed.

**Business takeaway:**
Clean user directories reduce security risk and eliminate unnecessary license spend.

**Speaker notes:**
User counts include all Entra ID users regardless of license status.
Guests are external users. Review each to confirm they still need access.
Disabled users consuming licenses are a cost and compliance concern.
Stale accounts can represent a security risk — anyone with credentials to an inactive account can access any services still assigned.
Action: Enable Azure AD sign-in activity reporting and review whether unlicensed accounts are required or should be removed.

---

## Slide 6 — Exchange & Mail Flow

**Visual treatment:** Icon-based stat row across top (mailboxes / forwarding / inbox rules / transport rules). Below: table snippet of forwarding mailboxes if any.

> 0 mailboxes with external forwarding — review required

- Total mailboxes: **9**
- Mailboxes with forwarding enabled: **0** (potential data exfiltration risk)
- Inbox forwarding rules: **N/A** (user-level)
- Transport rules (tenant-level): **N/A** (admin-defined mail flow rules)
- Action: Validate mailbox statistics for accuracy, review the Escala shared mailbox, and assess the necessity of existing inbox forwarding rules.

**Business takeaway:**
Unreviewed forwarding and transport rules are a compliance and data security risk.

**Speaker notes:**
External mail forwarding is one of the most common data exfiltration vectors.
Each of the 0 forwarding mailboxes should have a documented business justification.
Transport rules can silently redirect mail — review for any unexpected destinations.
Inbox rules (N/A) are user-created and harder to audit centrally; consider running a sweep with a script.
Action: Validate mailbox statistics for accuracy, review the Escala shared mailbox, and assess the necessity of existing inbox forwarding rules.

---

## Slide 7 — SharePoint & OneDrive

**Visual treatment:** KPI tiles (sites / storage / external sharing). Bar or treemap of top sites by storage. External sharing flagged in amber/red.

> 38 sites using 22.12 GB — 0 with external sharing enabled

- Total SharePoint sites: **38**
- Total storage used: **22.12 GB**
- Sites with external sharing enabled: **0** (review required)
- Largest sites are listed in the report appendix
- Action: Enable OneDrive reporting and periodically review site ownership and lifecycle to prevent sprawl.

**Business takeaway:**
External sharing controls and storage growth are the two most actionable SharePoint findings.

**Speaker notes:**
SharePoint is often the most data-rich area of a tenant review.
0 sites allow external access — each should have a documented reason.
Storage growth is worth tracking trend over time. Consider alerting at 75% of allocated quota.
Orphaned sites (no owner) are a governance risk. If ownership data is available, review and assign.
Action: Enable OneDrive reporting and periodically review site ownership and lifecycle to prevent sprawl.

---

## Slide 8 — Teams Collaboration Health

**Visual treatment:** Stat cards row (total / active / inactive). Table of inactive teams. Activity heatmap if available. Amber callout for inactive count.

> 2 inactive Teams out of 3 total — governance review required

- Total Teams: **3**
- Inactive Teams: **2** (no recent activity)
- Active Teams: **1** (estimate)
- Private Teams may have unmanaged guest access
- Action: Confirm whether Teams is intended for active collaboration and archive or remove unused teams if not required.

**Business takeaway:**
Teams sprawl increases storage and compliance overhead. A lifecycle policy prevents the problem from growing.

**Speaker notes:**
Inactive Teams accumulate files, channels, and guest access over time.
Microsoft defines inactivity as no messages, meetings, or file activity in 30 days (or configured threshold).
Archive rather than delete where project history should be preserved.
Review guest membership in all teams — especially inactive ones.
Action: Confirm whether Teams is intended for active collaboration and archive or remove unused teams if not required.

---

## Slide 9 — Devices & Endpoint Hygiene

**Visual treatment:** Donut chart: Windows vs Android vs other. Three KPI cards: Total / Stale / Intune Managed. Red callout if Intune = 0.

> 5 stale devices out of 9 — N/A managed by Intune

- Total registered devices: **9**
- Stale devices (no recent sign-in): **5**
- Intune-managed devices: **N/A**
- Without Intune: no compliance policy, no conditional access baseline, no remote wipe
- Action: Implement Intune device management and review or remove stale device registrations.

**Business takeaway:**
Device management is foundational to Zero Trust. Even a basic Intune deployment significantly improves security posture.

**Speaker notes:**
Devices registered to Entra ID but not in Intune have no compliance enforcement.
Stale devices may represent hardware that is decommissioned but still has active credentials.
Removing stale devices cleans up the device list and prevents token reuse on old hardware.
If Intune is at 0, this is a significant security gap — recommend a phased deployment starting with Windows endpoints.
Action: Implement Intune device management and review or remove stale device registrations.

---

## Slide 10 — Copilot Adoption & Value

**Visual treatment:** Hero KPI: Copilot licenses purchased and assigned. Usage bar (assigned vs active). Right: ROI callout — time saved per user, cost per seat.

> 1 of 1 Copilot licenses assigned — 0 unused

- Copilot for M365 licenses purchased: **1**
- Licenses assigned to users: **1**
- Potentially unused: **0**
- Copilot requires M365 E3/E5 or Business Premium as a base license
- Action: Enable Copilot usage reporting to measure adoption and return on investment.

**Business takeaway:**
Copilot ROI is maximised through structured adoption. Unassigned or unactivated licenses are an immediate cost saving opportunity.

**Speaker notes:**
Copilot for Microsoft 365 is one of the highest per-seat costs in the M365 family.
Unused Copilot licenses represent significant avoidable spend.
True usage data (prompts, active users) requires the Microsoft 365 admin usage reports — available in the admin centre.
If the client has not yet run a Copilot adoption programme, this is an opportunity to position RTH's managed adoption service.
Action: Enable Copilot usage reporting to measure adoption and return on investment.

---

## Slide 11 — Top Recommendations

**Visual treatment:** Numbered list with priority colour coding (red = P1, amber = P2, blue = P3+). Each row: priority, category, action, effort tag, impact tag.

> 7 actionable recommendations identified

- **P1** | Review 1049993 unassigned licenses before next renewal _(Effort: Low / Impact: Medium)_
- **P2** | Review 17 inbox forwarding rules _(Effort: Medium / Impact: Medium)_
- **P3** | Archive or clean up 2 inactive Teams _(Effort: Low / Impact: Low)_
- **P4** | Clean up 5 stale device records _(Effort: Low / Impact: Medium)_
- **P5** | Consider enrolling devices in Intune for management and compliance coverage _(Effort: High / Impact: High)_
- **P6** | Enable Copilot usage reporting to measure adoption _(Effort: Low / Impact: Medium)_

**Business takeaway:**
These recommendations are data-driven and ranked by priority. Completing P1 and P2 items will deliver the most value in the shortest time.

**Speaker notes:**
Walk the client through each recommendation, pausing to agree on priority and ownership.
Effort Low = days, Medium = weeks, High = weeks to months.
Impact High = material cost saving or significant risk reduction.
Assign an owner for each item before leaving the meeting.
Follow up with a written action plan within 48 hours.

---

## Slide 12 — Next Quarter Action Plan

**Visual treatment:** Kanban-style columns: Immediate (P1) / This Quarter (P2) / Roadmap (P3+). Each item is a card with owner and due date placeholder.

> Commit to actions before this meeting closes

- [ ] **Review 1049993 unassigned licenses before next renewal** — Owner: License Manager / Finance
- [ ] **Review 17 inbox forwarding rules** — Owner: Exchange Admin

**Business takeaway:**
The value of a review is in the actions it drives. Leave today with clear commitments and owners.

**Speaker notes:**
Use this slide to close the meeting with clear commitments.
Each action should have a named owner and an agreed delivery date.
RTH to send a written summary of agreed actions within 48 hours.
Schedule the next quarterly review now while everyone is in the room.
Offer to track action completion progress as a managed service engagement.

---

## Slide 13 — Appendix — Data Sources & Coverage

**Visual treatment:** Two-column layout. Left: data source coverage table. Right: key definitions and notes.

> All data collected live from the RTH Tech Services Inc. Microsoft 365 tenant

- **TenantOverview**: Collected
- **LicenseInventory**: Collected
- **UserInventory**: Collected
- **LicenseUserAnalysis**: Collected
- **MailboxInventory**: Collected
- **SharePoint**: Collected
- **Teams**: Collected
- **Devices**: Collected
- **Copilot**: Collected

**Business takeaway:**
Full detail tables are in the accompanying HTML report.

**Speaker notes:**
This slide documents where data came from and any gaps.
Skipped datasets were not available due to licensing, permissions, or configuration at the time of collection.
For a complete review, ensure the service account has at minimum: Global Reader, Exchange View-Only Admin, SharePoint Admin, Teams Admin.
The full technical data tables are in the TenantReviewReport.html file provided alongside this deck.
Report generated: May 16, 2026

