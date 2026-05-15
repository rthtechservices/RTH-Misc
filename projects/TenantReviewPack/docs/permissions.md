# Permissions

TenantReviewPack needs different Microsoft 365 permissions depending on which collectors are enabled.

## Module Prerequisites

- Required: `Microsoft.Graph`
- Optional for later Exchange collectors: `ExchangeOnlineManagement`

Install commands:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Authentication Modes

TenantReviewPack reads authentication from `ConnectConfig.json`.

### AppCertificate

Use app-only certificate authentication for repeatable tenant reviews:

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

The certificate must be installed in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`. Do not commit PFX files, private keys, PEM files, exported certificates, or secret-bearing JSON files.

### Interactive

Use interactive authentication for development and troubleshooting. Interactive mode requests these delegated scopes:

- `User.Read.All`
- `Directory.Read.All`
- `Organization.Read.All`
- `LicenseAssignment.Read.All`
- `Reports.Read.All`
- `AuditLog.Read.All`

## Required Microsoft Graph Permission Areas

For the first real collectors, grant admin consent for these application permissions when using AppCertificate mode:

- `User.Read.All`
- `Directory.Read.All`
- `Organization.Read.All`
- `LicenseAssignment.Read.All`
- `Reports.Read.All`
- `AuditLog.Read.All`

Additional collectors may later need permissions for groups, Teams, devices, directory roles, service health, SharePoint usage, and reporting endpoints.

## Exchange Online

Exchange Online PowerShell will be required for the most reliable mailbox forwarding, inbox rule, shared mailbox, and transport rule data. Exchange connection is currently optional and does not fail the whole run unless explicitly configured as required.

## SharePoint Online

SharePoint Online Management Shell or PnP PowerShell is recommended for site inventory, storage usage, and sharing posture.
