# Data Inventory

## Tenant Overview

- Organization display name
- Tenant ID presence flag by default
- Default and initial domains
- Verified, federated, and total domain counts
- Domain authentication type and supported services
- Password validity period, assigned plan count, and technical notification mails where available

## Licensing

- Subscribed SKUs from `Get-MgSubscribedSku`
- Purchased, assigned, suspended, warning, and unused units
- SKU display names and monthly costs from local price map
- Estimated monthly and annual costs
- Estimated unused monthly and annual costs
- Service plan names and provisioning status

## Users

- Members and guests from `Get-MgUser`
- Enabled/disabled account state
- Department, job title, company, mail, and UPN
- Assigned SKU IDs and license counts
- Sign-in activity when available
- Stale, licensed-stale, and licensed-disabled flags
- Guest users with assigned licenses

## License and User Analysis

- Unused license count and estimated unused spend
- Disabled licensed users
- Stale licensed users
- Guest licensed users
- Plain-English attention items

## Mailbox

- User, shared, room, and equipment mailbox counts
- Mailbox-level forwarding settings
- External forwarding suspicion
- Transport rules and enabled rule count
- Optional inbox forwarding rule scan
- Optional mailbox sizes and largest mailboxes

## SharePoint and OneDrive

- Tenant sites from PnP/SPO where available
- Graph SharePoint usage report fallback
- Site URL, title, template, storage usage, quota, modified date, and sharing capability
- OneDrive owner, URL, storage, active file, and file counts when enabled

## Teams

- Teams-backed Microsoft 365 groups
- Team visibility, created date, archive state where available
- Optional owner and member counts
- Teams activity report metrics such as last activity, channel messages, replies, and meetings

## Devices

- Entra device ID, display name, enabled state, OS, OS version, trust type, and approximate last sign-in
- Stale device flag
- Optional Intune managed status, compliance state, and primary user

## Copilot

- Copilot-related SKUs matched from license inventory
- Purchased, assigned, unused, and estimated spend
- Licensed Copilot users from user inventory
- Copilot usage report when the tenant and Graph module expose it
