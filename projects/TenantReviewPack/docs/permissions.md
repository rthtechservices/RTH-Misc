# Permissions

TenantReviewPack will need different Microsoft 365 permissions depending on which collectors are enabled.

## Recommended approach

Start interactive while developing. Move to app-only authentication once the collector set stabilizes.

## Likely Microsoft Graph permission areas

- Organization and domain details
- User inventory
- Group and Teams inventory
- Reports and usage data
- Device inventory
- Directory roles
- Service health

## Exchange Online

Exchange Online PowerShell is required for the most reliable mailbox forwarding, inbox rule, shared mailbox, and transport rule data.

## SharePoint Online

SharePoint Online Management Shell or PnP PowerShell is recommended for site inventory, storage usage, and sharing posture.

## Notes

Document the final permission set here once collectors are implemented and tested against a real tenant.
