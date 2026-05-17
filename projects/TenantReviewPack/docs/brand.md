# Tenant Review Report Brand Guide

## Colour Palette

- Deep purple `#5a2f61`: cover, section header bars, recommendation priority blocks, primary headings.
- Accent purple `#6f28db`: secondary brand accent and light callout support.
- Blue `#4fa7db`: cover accent band, information status, recommended-action callouts.
- Neutral grey `#767779`: metadata, secondary labels, not-available status.
- Green `#27ae60`: healthy status.
- Amber `#f39c12`: watch status.
- Red `#e74c3c`: action-recommended status.

## Typography

- Use Aptos or Segoe UI for Word output.
- Cover title should be large, bold, and centred.
- Section headers should be strong purple bars with white text.
- Body copy should stay readable at normal business-report sizes.
- Appendix tables may use a smaller font, but only after columns have been curated.

## Section Hierarchy

- Cover page comes first and must show report title, tenant/client name, review period, prepared-by line, and generated date.
- Executive Summary follows the cover and uses a health statement callout plus short executive bullets.
- Key Metrics at a Glance follows the executive summary and uses KPI cards or Word-friendly card tables.
- Main narrative sections follow the HTML report order:
  Tenant Overview, Cost & Licensing Review, License Optimisation, Identity & User Review, Exchange & Mail Flow, SharePoint & OneDrive, Teams Collaboration, Devices & Endpoints, Copilot Review.
- Top Recommendations, Data Coverage, and Appendix close the report.

## KPI Card Rules

- Each KPI block includes a large value, a short label, a plain-English description, and a status badge.
- Use three KPI cards per row in Word.
- Do not invent metrics. Reuse the narrative `kpis` collection generated for the HTML report.
- Keep descriptions short enough to scan in a card.

## Status Badge Rules

- `Good` renders as `Healthy` using green.
- `Watch` uses amber.
- `ActionRecommended` renders as `Action Recommended` using red.
- `Informational` and unknown statuses render as `Info` using blue.
- `NotAvailable` renders as `Not Available` using grey.

## Appendix Table Rules

- Appendix tables must be curated, not raw object dumps.
- Use only useful client-facing columns.
- Avoid GUID-heavy fields unless they are the business identifier.
- Convert arrays to readable comma-separated labels.
- Convert complex objects to selected display properties or omit them.
- Never output `System.Object[]`, raw hashtable syntax, Graph model type names, or raw JSON.
- Use landscape orientation for the appendix when tables are wide.
- Repeat header rows where possible and use purple table headers with alternating light row shading.

## Tone And Wording

- Use plain-English business language from the generated narrative.
- Avoid developer phrasing in the client-facing report.
- Put warnings and collector limitations in Data Coverage as simple status labels, not raw warning arrays.
- Recommendations should explain why the action matters and who should own it where available.
