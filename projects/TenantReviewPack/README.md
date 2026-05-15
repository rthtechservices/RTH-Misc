# TenantReviewPack

TenantReviewPack is a repeatable Microsoft 365 tenant review package generator.

The goal is to collect tenant data, normalize it into clean JSON datasets, enrich the summaries with plain-language AI narratives, and generate a polished client-facing report and PowerPoint deck for quarterly or annual reviews.

## Intended Output

- Executive report
- PowerPoint slide deck
- Raw JSON/CSV exports
- Appendix-style technical evidence
- Cost, security, usage, storage, device, and collaboration highlights

## Current Project Status

The scaffold now includes Microsoft Graph authentication plus the first production collectors:

1. Tenant overview placeholder
2. License inventory
3. User inventory
4. License and user analysis

Mailbox, SharePoint, Teams, device, and Copilot collectors are still placeholders.

## Prerequisites

Install Microsoft Graph PowerShell before collecting data:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Exchange Online support is prepared for later mailbox collectors. Install it only when those collectors are implemented or explicitly required:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Authentication

Authentication is read from `ConnectConfig.json`. Do not copy tenant IDs, client IDs, certificate thumbprints, private keys, PFX files, or exported certificate material into other repo files.

App-only certificate authentication uses this shape:

```json
{
  "auth": {
    "mode": "AppCertificate",
    "tenantId": "...",
    "clientId": "...",
    "certificateThumbprint": "...",
    "exchangeOrganization": "contoso.onmicrosoft.com"
  }
}
```

The certificate thumbprint must exist in either `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`. The script reports whether the values were loaded and whether the certificate was found, but it does not print the tenant ID, client ID, or thumbprint.

For development, `auth.mode` can be set to `Interactive`. Interactive mode requests delegated Graph scopes for user, directory, organization, license assignment, reports, and audit log reads.

## Example Command

```powershell
.\Invoke-TenantReviewPack.ps1 `
  -TenantName "RTH Tech Services" `
  -ReviewPeriod "Q2 2026" `
  -ConnectConfigPath ".\ConnectConfig.json" `
  -SkipAI `
  -SkipRender
```

## Validation

Parser and function-load validation does not require Microsoft 365 login:

```powershell
.\Test-TenantReviewPack.ps1
```

## Expected Outputs

Each run creates a timestamped folder under `output` with JSON datasets including:

- `LicenseInventory.json`
- `UserInventory.json`
- `LicenseUserAnalysis.json`

## Known Limitations

- Microsoft Graph does not provide exact client license pricing. License costs come from a local price map such as `config/license-prices.sample.json`.
- `signInActivity` may be unavailable depending on permissions, licensing, and Graph API behavior. User collection falls back to a no-sign-in query and records a warning.
- Exchange collectors are not implemented yet. Exchange Online authentication readiness is optional and does not fail the run unless configured as required.

## Design Principle

The scripts own the facts and layout. AI only writes short, client-friendly explanations based on supplied JSON.
