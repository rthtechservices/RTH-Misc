# Roadmap

## Completed Foundation

- App certificate and interactive Graph auth
- Optional Exchange and SharePoint connection attempts with warnings
- Stable JSON dataset shape across collectors
- Local price-map based license costing
- Local deterministic narrative fallback
- Markdown/HTML report output
- Deck outline output
- Offline validation script

## Implemented Collectors

- Tenant overview
- License inventory
- User inventory
- License/user analysis
- Mailbox inventory
- SharePoint and OneDrive inventory
- Teams inventory
- Device inventory
- Copilot inventory

## Next Improvements

- Add richer Exchange RBAC detection and friendly remediation messages.
- Add direct SharePoint admin settings export when PnP app-only permissions are confirmed.
- Add service health and message center collector.
- Add admin role and privileged account collector.
- Add richer PowerPoint generation when a dependency is intentionally chosen.
- Add CSV exports for selected appendix tables.

## Useful Commands

```powershell
.\Test-TenantReviewPack.ps1
```

```powershell
.\Invoke-TenantReviewPack.ps1 `
  -TenantName "RTH Tech Services" `
  -ReviewPeriod "Q2 2026" `
  -ConnectConfigPath ".\ConnectConfig.json" `
  -SkipAI
```

```powershell
.\Invoke-TenantReviewPack.ps1 `
  -TenantName "RTH Tech Services" `
  -ReviewPeriod "Q2 2026" `
  -ConnectConfigPath ".\ConnectConfig.json" `
  -SkipAI `
  -SkipRender
```
