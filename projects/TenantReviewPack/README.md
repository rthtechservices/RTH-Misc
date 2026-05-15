# TenantReviewPack

TenantReviewPack is a repeatable Microsoft 365 tenant review package generator.

The goal is to collect tenant data, normalize it into clean JSON datasets, enrich the summaries with plain-language AI narratives, and generate a polished client-facing report and PowerPoint deck for quarterly or annual reviews.

## Intended output

- Executive report
- PowerPoint slide deck
- Raw JSON/CSV exports
- Appendix-style technical evidence
- Cost, security, usage, storage, device, and collaboration highlights

## Current project status

Initial scaffold only. The first implementation milestone should focus on high-value collectors:

1. Tenant overview
2. License inventory
3. User inventory
4. Mailbox inventory
5. SharePoint / OneDrive usage

## Proposed command

```powershell
.\Invoke-TenantReviewPack.ps1 -TenantName "Client Name" -ReviewPeriod "Q4 2026" -OutputPath ".\output"
```

## Design principle

The scripts own the facts and layout. AI only writes short, client-friendly explanations based on supplied JSON. No free-range report goblinry.
