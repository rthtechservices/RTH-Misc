# TenantReviewPack

TenantReviewPack is a repeatable Microsoft 365 tenant review package generator.

It connects to Microsoft 365, collects review-ready JSON datasets, produces deterministic narrative summaries, and renders simple Markdown/HTML report artifacts plus a deck outline.

## Implemented Collectors

- Tenant overview: organization and verified domain data from Microsoft Graph.
- License inventory: subscribed SKUs, assigned/unused units, service plans, and local price-map estimates.
- User inventory: users, guests, license state, enabled/disabled state, and sign-in staleness where Graph returns sign-in activity.
- License/user analysis: unused license cost, disabled licensed users, stale licensed users, and licensed guests.
- Mailbox inventory: Exchange mailbox forwarding, transport rules, optional inbox rules, and optional mailbox statistics when Exchange Online PowerShell is connected.
- SharePoint and OneDrive: SharePoint sites from PnP/SPO when available, with Graph report fallback, plus optional OneDrive usage.
- Teams: Teams-backed Microsoft 365 groups plus Teams usage report data where available.
- Devices: Entra devices plus optional Intune managed-device data.
- Copilot: Copilot-related SKU inventory, licensed users, and usage report data where available.

## Modules

Required for Graph authentication and most collectors:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Optional modules:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

Optional module absence is reported as a warning and the affected collector returns limited or empty data.

## Authentication

Authentication is read from `ConnectConfig.json`:

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

The script does not print tenant IDs, client IDs, certificate thumbprints, secrets, private keys, or tokens. The certificate thumbprint must exist in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`.

Interactive Graph mode is also supported for development.

## Run Validation

```powershell
.\Test-TenantReviewPack.ps1
```

The test parses scripts, dot-sources safe functions, runs analyzer/narrative/renderer mock checks, and fails if source files still contain old marker text.

## Run Data Collection Only

```powershell
.\Invoke-TenantReviewPack.ps1 `
  -TenantName "RTH Tech Services" `
  -ReviewPeriod "Q2 2026" `
  -ConnectConfigPath ".\ConnectConfig.json" `
  -SkipAI `
  -SkipRender
```

## Run Collection With Local Narrative and Rendering

```powershell
.\Invoke-TenantReviewPack.ps1 `
  -TenantName "RTH Tech Services" `
  -ReviewPeriod "Q2 2026" `
  -ConnectConfigPath ".\ConnectConfig.json" `
  -SkipAI
```

`-SkipAI` skips external AI calls, but the script still creates a local rule-based narrative for renderers.

## Outputs

Each run creates a timestamped folder under `output` with:

- `TenantOverview.json`
- `LicenseInventory.json`
- `UserInventory.json`
- `LicenseUserAnalysis.json`
- `MailboxInventory.json`
- `SharePoint.json`
- `Teams.json`
- `Devices.json`
- `Copilot.json`
- `Narrative.json` when narrative is not skipped
- `TenantReviewReport.md` and `TenantReviewReport.html` when rendering is enabled
- `TenantReviewDeckOutline.md` when rendering is enabled

## AI Narrative

AI is optional. Configure `config/sample.settings.json` under `ai` and set the API key only in the environment variable named by `apiKeyEnvironmentVariable`. If AI is disabled, missing configuration, or returns invalid JSON, TenantReviewPack falls back to deterministic local narrative sections.

## Known Limitations

- Microsoft Graph does not provide exact client license pricing. Pricing comes from `config/license-prices.sample.json` or another local price map.
- `signInActivity` may be unavailable depending on tenant licensing, permissions, and Graph API behavior.
- Exchange app-only auth requires `Exchange.ManageAsApp` and a supported Exchange/Entra role assignment for the app service principal.
- SharePoint tenant settings require PnP/SPO module connectivity; otherwise the collector uses Graph usage reports where available.
- Graph report endpoints may return no data when reports are disabled, delayed, or not licensed.
