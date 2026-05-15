# Roadmap

## Milestone 1 - Scaffold

- Project structure
- Entry script
- Config samples
- Collector stubs
- Renderer stubs
- Documentation placeholders

## Milestone 2 - Minimum Useful Data Pack

- Tenant overview collector
- License inventory collector using `Get-MgSubscribedSku`
- User inventory collector using `Get-MgUser`
- License/user analyzer for unused licenses and risky assignments
- Mailbox/shared mailbox collector
- JSON output validation with `Test-TenantReviewPack.ps1`

## Milestone 3 - Storage and Collaboration

- SharePoint site inventory
- OneDrive usage
- Teams usage
- Inactive Teams/sites detection

## Milestone 4 - Risk and Cost Highlights

- Forwarding and transport rule flags
- Disabled users with licenses - first analyzer implemented
- Stale users/devices - stale licensed user analysis implemented, devices pending
- Unused license summary - first analyzer implemented using local price map
- External sharing summary

## Milestone 5 - AI Narrative

- Prompt contract
- Chunked dataset submission
- Strict JSON response validation
- Retry/fallback handling

## Milestone 6 - Client Package Rendering

- Branded Word/PDF report
- Branded PowerPoint deck
- Appendix exports
- Executive summary slide/page pattern

## Current Limitations

- Microsoft Graph does not provide exact client license pricing, so TenantReviewPack uses a local price map.
- `signInActivity` may be unavailable depending on permissions, licensing, or Graph API behavior.
- Exchange collectors are not implemented yet. Exchange authentication readiness is optional for now.

## Useful Commands

```powershell
.\Test-TenantReviewPack.ps1
```

```powershell
.\Invoke-TenantReviewPack.ps1 `
  -TenantName "RTH Tech Services" `
  -ReviewPeriod "Q2 2026" `
  -ConnectConfigPath ".\ConnectConfig.json" `
  -SkipAI `
  -SkipRender
```
