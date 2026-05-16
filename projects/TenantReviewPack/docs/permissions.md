# Permissions

TenantReviewPack is read-only reporting. Grant only the permissions needed for the collectors you intend to run.

## Required Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Optional:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

## Graph Application Permissions

Likely application permissions for the full data pack:

- `User.Read.All`
- `Directory.Read.All`
- `Organization.Read.All`
- `LicenseAssignment.Read.All`
- `Reports.Read.All`
- `AuditLog.Read.All`
- `Device.Read.All`
- `DeviceManagementManagedDevices.Read.All`
- `Group.Read.All`
- `Team.ReadBasic.All`
- `TeamSettings.Read.All`
- `Channel.ReadBasic.All` if channel-level expansion is added later
- `Sites.Read.All` for SharePoint/OneDrive report and site reads where Graph is used

Admin consent is required for app-only certificate authentication.

## Exchange Online

For Exchange app-only PowerShell:

- API permission: `Office 365 Exchange Online` > `Exchange.ManageAsApp`
- Admin consent granted
- Supported role assignment for the app service principal, commonly `Exchange Administrator` for broad review access or a narrower supported Exchange role where sufficient

If this is not configured, interactive runs prompt for a retry/continue/abort decision. Non-interactive runs fail on the warning.

## SharePoint

The SharePoint collector prefers read-only posture:

- PnP/SPO tenant site enumeration when already connected through supported app certificate auth
- Graph report-only collection when `sharePoint.siteSource` is `GraphReports`
- `Sites.Read.All` is generally preferred for read-only Graph access
- Some tenant administration reads may require broader SharePoint app permissions depending on module and tenant policy

If PnP/SPO tenant-site enumeration is unauthorized, interactive runs prompt before continuing. Choose the Graph-only retry when usage-report data is sufficient, or fix the app authorization and retry.

## Sensitive Values

Do not commit certificate files, private keys, exported certs, tokens, or secret-bearing JSON. The script reports only whether auth values were loaded and whether the certificate was found.
