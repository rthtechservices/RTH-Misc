# Data Inventory

This document tracks the datasets planned for TenantReviewPack.

## Tenant Overview

- Organization name
- Tenant ID
- Domains
- Verified domains
- Default domain
- Service health summary

## Licensing

- Subscribed SKUs from `Get-MgSubscribedSku`
- SKU part number and mapped display name
- Capability status
- Purchased/enabled units
- Assigned/consumed units
- Suspended units
- Warning units
- Unused units
- Local price-map monthly unit cost
- Estimated monthly and annual license cost
- Estimated unused monthly and annual license cost
- Service plan count and simple service plan status

## Users

- Member and guest users from `Get-MgUser`
- Account enabled/disabled state
- User profile fields including department, job title, and company
- Assigned license count and assigned SKU IDs
- Last successful, interactive, and non-interactive sign-in timestamps when Graph returns `signInActivity`
- Stale users based on last successful sign-in and the configured threshold
- Licensed disabled users
- Licensed stale users
- Guest users with licenses
- Users without sign-in data

`signInActivity` may be unavailable depending on permissions, tenant licensing, and Graph API behavior. The collector retries without sign-in activity and records a warning instead of failing the whole run.

## License and User Analysis

- Unused license count
- Estimated unused monthly and annual cost
- Disabled users with licenses
- Stale licensed users
- Guest users with licenses
- Plain-English attention items for review follow-up

## Exchange Online

- User mailboxes
- Shared mailboxes
- Mailbox forwarding
- Inbox rules with forwarding
- Transport rules
- Mailbox sizes

## SharePoint and OneDrive

- Sites
- Storage usage
- External sharing posture
- OneDrive usage
- Largest sites and OneDrives

## Teams

- Teams inventory
- Activity reports
- Inactive teams
- Meetings and messages where available

## Devices

- Entra devices
- Intune managed devices where available
- Compliance status
- Operating system versions
- Stale devices
- Office activation reports

## Copilot

- Assigned licenses
- Active users
- Usage trend
- Inactive licensed users
