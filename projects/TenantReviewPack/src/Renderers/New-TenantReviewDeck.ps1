function New-TenantReviewDeck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [string]$ReviewPeriod,

        [Parameter(Mandatory = $true)]
        [object]$Datasets,

        [Parameter(Mandatory = $true)]
        [object]$Narrative,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Set-StrictMode -Off

    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $outlinePath = Join-Path $OutputPath 'TenantReviewDeckOutline.md'
    $generatedAt = (Get-Date).ToString('MMMM d, yyyy')

    # ── helpers ──────────────────────────────────────────────────────────────

    function Nvl { param([object]$Value, [string]$Default = 'N/A') if ($null -eq $Value -or $Value.ToString() -eq '') { return $Default } return $Value.ToString() }

    function Get-Section { param([string]$Dataset) @(Get-TenantReviewProperty -InputObject $Narrative -Name 'sections') | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'dataset') -eq $Dataset } | Select-Object -First 1 }

    function Get-KpisForArea { param([string]$Area) @(Get-TenantReviewProperty -InputObject $Narrative -Name 'kpis') | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'area') -eq $Area } }

    function Get-TopKpi { param([string]$Area) (Get-KpisForArea -Area $Area | Select-Object -First 1) }

    function Format-KpiLine { param([object]$Kpi) if (-not $Kpi) { return '' } $val = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'value'); $lbl = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'label'); "**$val** — $lbl" }

    function Add-Slide { param([System.Collections.Generic.List[string]]$L, [int]$Num, [string]$Title, [string]$Visual = '', [string]$BigMessage = '', [string[]]$Bullets = @(), [string[]]$SpeakerNotes = @(), [string]$Takeaway = '') $L.Add("---"); $L.Add(""); $L.Add("## Slide $Num — $Title"); $L.Add(""); if ($Visual)     { $L.Add("**Visual treatment:** $Visual"); $L.Add('') } if ($BigMessage) { $L.Add("> $BigMessage"); $L.Add('') } if ($Bullets.Count) { foreach ($b in $Bullets) { $L.Add("- $b") }; $L.Add('') } $L.Add("**Business takeaway:**"); if ($Takeaway) { $L.Add($Takeaway) } else { $L.Add('See speaker notes.') }; $L.Add(""); $L.Add("**Speaker notes:**"); foreach ($n in $SpeakerNotes) { $L.Add($n) }; $L.Add('') }

    # ── pull data ─────────────────────────────────────────────────────────────

    $kpis            = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'kpis')
    $recommendations = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'recommendations')
    $execBullets     = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'executiveSummaryBullets')
    $healthStatement = Nvl (Get-TenantReviewProperty -InputObject $Narrative -Name 'tenantHealthStatement') 'Review the findings in this report and act on the highlighted recommendations.'

    $tenantSummary  = Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'summary'
    $licSummary     = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
    $anaSummary     = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
    $userSummary    = Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'summary'
    $mbxSummary     = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
    $spSummary      = Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'summary'
    $teamsSummary   = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
    $deviceSummary  = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
    $copilotSummary = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'

    $orgName = Nvl (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'organizationName') $TenantName

    # Friendly metric extracts
    $totalUsers    = Nvl (Get-TenantReviewProperty -InputObject $userSummary   -Name 'totalUsers')
    $memberUsers   = Nvl (Get-TenantReviewProperty -InputObject $userSummary   -Name 'memberUsers')
    $guestUsers    = Nvl (Get-TenantReviewProperty -InputObject $userSummary   -Name 'guestUsers')
    $totalLic      = Nvl (Get-TenantReviewProperty -InputObject $licSummary    -Name 'totalPurchasedUnits')
    $assignedLic   = Nvl (Get-TenantReviewProperty -InputObject $licSummary    -Name 'totalAssignedUnits')
    $unusedLic     = Nvl (Get-TenantReviewProperty -InputObject $licSummary    -Name 'totalUnusedUnits')
    $totalSkus     = Nvl (Get-TenantReviewProperty -InputObject $licSummary    -Name 'totalSkus')
    $monthlyCost   = Nvl (Get-TenantReviewProperty -InputObject $licSummary    -Name 'estimatedMonthlyCost')
    $disabledLic   = Nvl (Get-TenantReviewProperty -InputObject $anaSummary   -Name 'disabledLicensedUserCount')
    $staleLic      = Nvl (Get-TenantReviewProperty -InputObject $anaSummary   -Name 'staleLicensedUserCount')
    $guestLic      = Nvl (Get-TenantReviewProperty -InputObject $anaSummary   -Name 'guestLicensedUserCount')
    $totalMbx      = Nvl (Get-TenantReviewProperty -InputObject $mbxSummary   -Name 'totalMailboxes')
    $extFwd        = Nvl (Get-TenantReviewProperty -InputObject $mbxSummary   -Name 'mailboxesForwardingExternally')
    $inboxRules    = Nvl (Get-TenantReviewProperty -InputObject $mbxSummary   -Name 'totalInboxForwardingRules')
    $transportRules = Nvl (Get-TenantReviewProperty -InputObject $mbxSummary  -Name 'totalTransportRules')
    $totalSites    = Nvl (Get-TenantReviewProperty -InputObject $spSummary    -Name 'totalSites')
    $spStorage     = Nvl (Get-TenantReviewProperty -InputObject $spSummary    -Name 'totalStorageGB')
    $extSharing    = Nvl (Get-TenantReviewProperty -InputObject $spSummary    -Name 'externalSharingEnabledSites')
    $totalTeams    = Nvl (Get-TenantReviewProperty -InputObject $teamsSummary -Name 'totalTeams')
    $inactiveTeams = Nvl (Get-TenantReviewProperty -InputObject $teamsSummary -Name 'inactiveTeams')
    $totalDevices  = Nvl (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'totalDevices')
    $staleDevices  = Nvl (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'staleDevices')
    $intuneDevices = Nvl (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'intuneManaged')
    $copPurchased  = Nvl (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotPurchased')
    $copAssigned   = Nvl (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotAssigned')
    $copUnused     = Nvl (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotUnused')

    # ── top rec lines ─────────────────────────────────────────────────────────

    $topRecLines = if ($recommendations.Count -gt 0) {
        $recommendations | Select-Object -First 6 | ForEach-Object {
            $pri   = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'priority')
            $title = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'title')
            $eff   = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'effort')
            $imp   = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'impact')
            "**P$pri** | $title _(Effort: $eff / Impact: $imp)_"
        }
    } else {
        @('No specific recommendations were generated from the available data.')
    }

    # ── build lines ──────────────────────────────────────────────────────────

    $lines = [System.Collections.Generic.List[string]]::new()

    $lines.Add("# Microsoft 365 Tenant Review — Deck Outline")
    $lines.Add("")
    $lines.Add("| | |")
    $lines.Add("| --- | --- |")
    $lines.Add("| **Organisation** | $orgName |")
    $lines.Add("| **Review Period** | $ReviewPeriod |")
    $lines.Add("| **Prepared by** | RTH Tech Services |")
    $lines.Add("| **Generated** | $generatedAt |")
    $lines.Add("")
    $lines.Add("> This file is a structured slide-by-slide deck outline. Use it to build a presentation in PowerPoint, Google Slides, or any compatible tool.")
    $lines.Add("> Each slide includes a suggested visual treatment, a key message, supporting bullets, a business takeaway, and speaker notes.")
    $lines.Add("")

    # ── Slide 1: Cover ───────────────────────────────────────────────────────

    Add-Slide -L $lines -Num 1 -Title 'Cover' `
        -Visual 'Full-bleed gradient background (#5a2f61 -> #4fa7db). Client logo area centred when available. Review period and "Microsoft 365 Tenant Review" in white text. RTH Tech Services branding bottom right.' `
        -BigMessage "Microsoft 365 Tenant Review — $orgName — $ReviewPeriod" `
        -SpeakerNotes @(
            "Welcome the client. Introduce this as their Microsoft 365 Quarterly Business Review.",
            "Review period: $ReviewPeriod.",
            "Prepared by RTH Tech Services. This deck is based on data collected from live tenant systems.",
            "All findings are as at the collection date shown on the report."
        ) `
        -Takeaway "This is a professional, data-driven view of the $orgName Microsoft 365 environment."

    # ── Slide 2: Executive Snapshot ──────────────────────────────────────────

    $execBulletLines = if ($execBullets.Count) { $execBullets } else { @('Tenant review completed. See detailed findings in following slides.') }

    Add-Slide -L $lines -Num 2 -Title 'Executive Snapshot' `
        -Visual 'Single-column card layout on a light background. Status pill at top right (green/amber/red). Bullet list with icon markers.' `
        -BigMessage $healthStatement `
        -Bullets $execBulletLines `
        -SpeakerNotes @(
            "High-level tenant health statement: $healthStatement",
            "Walk the client through each bullet, pausing for questions.",
            "Position this as the 'short story' — the rest of the deck is the evidence.",
            "Reinforce that findings are based on live data, not estimates."
        ) `
        -Takeaway "The tenant is broadly operational. Priority items are highlighted in this review."

    # ── Slide 3: Big Numbers Dashboard ───────────────────────────────────────

    $allKpiLines = $kpis | Select-Object -First 9 | ForEach-Object {
        $v = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'value')
        $l = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'label')
        $s = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'status') 'Informational'
        "**$v** — $l [$s]"
    }

    Add-Slide -L $lines -Num 3 -Title 'Big Numbers Dashboard' `
        -Visual 'KPI card grid (3×3 or 3×4). Each card: large number, short label, status badge (green/amber/red). Cards grouped by area.' `
        -BigMessage "Key numbers at a glance across all review areas" `
        -Bullets $allKpiLines `
        -SpeakerNotes @(
            "These numbers come directly from the tenant data collected during this review.",
            "Highlight any red or amber cards as items to discuss in later slides.",
            "Green items are healthy — acknowledge them to balance the conversation.",
            "Remind the client that some metrics (especially Copilot) require Entra P1/P2 or specific permissions to collect fully."
        ) `
        -Takeaway "These headline numbers frame the detailed findings that follow."

    # ── Slide 4: Licensing & Cost Opportunities ───────────────────────────────

    $licSec = Get-Section -Dataset 'LicenseInventory'
    $licHeadline = Nvl (Get-TenantReviewProperty -InputObject $licSec -Name 'headline') 'License inventory reviewed.'
    $licAction   = Nvl (Get-TenantReviewProperty -InputObject $licSec -Name 'recommendedAction') 'Review license assignments before next renewal.'
    $anaSec = Get-Section -Dataset 'LicenseUserAnalysis'
    $anaAction = Nvl (Get-TenantReviewProperty -InputObject $anaSec -Name 'recommendedAction') 'Remove licenses from disabled and stale accounts.'

    Add-Slide -L $lines -Num 4 -Title 'Licensing & Cost Opportunities' `
        -Visual 'Split layout: left side big numbers (SKUs, total purchased, assigned, unused). Right side: mini bar representing utilisation. Amber/red callouts for disabled and stale licensed users.' `
        -BigMessage "$unusedLic unused license units across $totalSkus SKUs" `
        -Bullets @(
            "Total purchased units: **$totalLic** | Assigned: **$assignedLic** | Unused: **$unusedLic**"
            "Active SKUs: $totalSkus | Estimated monthly cost: **$monthlyCost**"
            "Disabled users with active licenses: **$disabledLic** → unnecessary spend"
            "Stale (inactive) licensed users: **$staleLic** → potential cleanup"
            "Guest users with licenses: **$guestLic** → confirm if intentional"
            "Action: $licAction"
        ) `
        -SpeakerNotes @(
            "Walk through the license SKU breakdown from the report appendix.",
            "$disabledLic disabled users with active licenses represent avoidable cost.",
            "$staleLic stale licensed users should be reviewed before the next renewal or monthly true-up.",
            "Not all SKUs have price data — work with the client to confirm exact pricing from their agreement.",
            "Action: $anaAction"
        ) `
        -Takeaway "License optimisation is one of the highest-ROI improvements available. Even modest cleanup can reduce monthly M365 spend."

    # ── Slide 5: Users & Access Review ───────────────────────────────────────

    $userSec    = Get-Section -Dataset 'UserInventory'
    $userAction = Nvl (Get-TenantReviewProperty -InputObject $userSec -Name 'recommendedAction') 'Review guest and stale user accounts.'

    Add-Slide -L $lines -Num 5 -Title 'Users & Access Review' `
        -Visual 'Donut chart: Members vs Guests. Three highlight cards below: Total users, Active, Guests. Status badges.' `
        -BigMessage "$totalUsers users in directory — $memberUsers members, $guestUsers guests" `
        -Bullets @(
            "Total directory users: **$totalUsers** ($memberUsers members, $guestUsers guests)"
            "Disabled licensed users: **$disabledLic** (confirm these are intentionally retained)"
            "Stale licensed users: **$staleLic** (no sign-in activity in review period)"
            "Guest users with licenses: **$guestLic** (confirm business justification)"
            "Action: $userAction"
        ) `
        -SpeakerNotes @(
            "User counts include all Entra ID users regardless of license status.",
            "Guests are external users. Review each to confirm they still need access.",
            "Disabled users consuming licenses are a cost and compliance concern.",
            "Stale accounts can represent a security risk — anyone with credentials to an inactive account can access any services still assigned.",
            "Action: $userAction"
        ) `
        -Takeaway "Clean user directories reduce security risk and eliminate unnecessary license spend."

    # ── Slide 6: Exchange & Mail Flow ─────────────────────────────────────────

    $mbxSec    = Get-Section -Dataset 'MailboxInventory'
    $mbxHline  = Nvl (Get-TenantReviewProperty -InputObject $mbxSec -Name 'headline')    'Mailbox inventory reviewed.'
    $mbxAction = Nvl (Get-TenantReviewProperty -InputObject $mbxSec -Name 'recommendedAction') 'Review forwarding rules and transport rules for business justification.'

    Add-Slide -L $lines -Num 6 -Title 'Exchange & Mail Flow' `
        -Visual 'Icon-based stat row across top (mailboxes / forwarding / inbox rules / transport rules). Below: table snippet of forwarding mailboxes if any.' `
        -BigMessage "$extFwd mailboxes with external forwarding — review required" `
        -Bullets @(
            "Total mailboxes: **$totalMbx**"
            "Mailboxes with forwarding enabled: **$extFwd** (potential data exfiltration risk)"
            "Inbox forwarding rules: **$inboxRules** (user-level)"
            "Transport rules (tenant-level): **$transportRules** (admin-defined mail flow rules)"
            "Action: $mbxAction"
        ) `
        -SpeakerNotes @(
            "External mail forwarding is one of the most common data exfiltration vectors.",
            "Each of the $extFwd forwarding mailboxes should have a documented business justification.",
            "Transport rules can silently redirect mail — review for any unexpected destinations.",
            "Inbox rules ($inboxRules) are user-created and harder to audit centrally; consider running a sweep with a script.",
            "Action: $mbxAction"
        ) `
        -Takeaway "Unreviewed forwarding and transport rules are a compliance and data security risk."

    # ── Slide 7: SharePoint & OneDrive ────────────────────────────────────────

    $spSec    = Get-Section -Dataset 'SharePoint'
    $spAction = Nvl (Get-TenantReviewProperty -InputObject $spSec -Name 'recommendedAction') 'Review external sharing settings for all sites.'

    Add-Slide -L $lines -Num 7 -Title 'SharePoint & OneDrive' `
        -Visual 'KPI tiles (sites / storage / external sharing). Bar or treemap of top sites by storage. External sharing flagged in amber/red.' `
        -BigMessage "$totalSites sites using $spStorage GB — $extSharing with external sharing enabled" `
        -Bullets @(
            "Total SharePoint sites: **$totalSites**"
            "Total storage used: **$spStorage GB**"
            "Sites with external sharing enabled: **$extSharing** (review required)"
            "Largest sites are listed in the report appendix"
            "Action: $spAction"
        ) `
        -SpeakerNotes @(
            "SharePoint is often the most data-rich area of a tenant review.",
            "$extSharing sites allow external access — each should have a documented reason.",
            "Storage growth is worth tracking trend over time. Consider alerting at 75% of allocated quota.",
            "Orphaned sites (no owner) are a governance risk. If ownership data is available, review and assign.",
            "Action: $spAction"
        ) `
        -Takeaway "External sharing controls and storage growth are the two most actionable SharePoint findings."

    # ── Slide 8: Teams Collaboration Health ──────────────────────────────────

    $teamsSec    = Get-Section -Dataset 'Teams'
    $teamsAction = Nvl (Get-TenantReviewProperty -InputObject $teamsSec -Name 'recommendedAction') 'Archive or delete inactive teams.'

    Add-Slide -L $lines -Num 8 -Title 'Teams Collaboration Health' `
        -Visual 'Stat cards row (total / active / inactive). Table of inactive teams. Activity heatmap if available. Amber callout for inactive count.' `
        -BigMessage "$inactiveTeams inactive Teams out of $totalTeams total — governance review required" `
        -Bullets @(
            "Total Teams: **$totalTeams**"
            "Inactive Teams: **$inactiveTeams** (no recent activity)"
            "Active Teams: **$([int]$totalTeams - [int]$inactiveTeams)** (estimate)"
            "Private Teams may have unmanaged guest access"
            "Action: $teamsAction"
        ) `
        -SpeakerNotes @(
            "Inactive Teams accumulate files, channels, and guest access over time.",
            "Microsoft defines inactivity as no messages, meetings, or file activity in 30 days (or configured threshold).",
            "Archive rather than delete where project history should be preserved.",
            "Review guest membership in all teams — especially inactive ones.",
            "Action: $teamsAction"
        ) `
        -Takeaway "Teams sprawl increases storage and compliance overhead. A lifecycle policy prevents the problem from growing."

    # ── Slide 9: Devices & Endpoint Hygiene ──────────────────────────────────

    $devSec    = Get-Section -Dataset 'Devices'
    $devAction = Nvl (Get-TenantReviewProperty -InputObject $devSec -Name 'recommendedAction') 'Review stale devices and consider deploying Intune.'

    Add-Slide -L $lines -Num 9 -Title 'Devices & Endpoint Hygiene' `
        -Visual 'Donut chart: Windows vs Android vs other. Three KPI cards: Total / Stale / Intune Managed. Red callout if Intune = 0.' `
        -BigMessage "$staleDevices stale devices out of $totalDevices — $intuneDevices managed by Intune" `
        -Bullets @(
            "Total registered devices: **$totalDevices**"
            "Stale devices (no recent sign-in): **$staleDevices**"
            "Intune-managed devices: **$intuneDevices**"
            "Without Intune: no compliance policy, no conditional access baseline, no remote wipe"
            "Action: $devAction"
        ) `
        -SpeakerNotes @(
            "Devices registered to Entra ID but not in Intune have no compliance enforcement.",
            "Stale devices may represent hardware that is decommissioned but still has active credentials.",
            "Removing stale devices cleans up the device list and prevents token reuse on old hardware.",
            "If Intune is at 0, this is a significant security gap — recommend a phased deployment starting with Windows endpoints.",
            "Action: $devAction"
        ) `
        -Takeaway "Device management is foundational to Zero Trust. Even a basic Intune deployment significantly improves security posture."

    # ── Slide 10: Copilot Adoption & Value ───────────────────────────────────

    $copSec    = Get-Section -Dataset 'Copilot'
    $copAction = Nvl (Get-TenantReviewProperty -InputObject $copSec -Name 'recommendedAction') 'Review Copilot license usage and build an adoption plan.'

    Add-Slide -L $lines -Num 10 -Title 'Copilot Adoption & Value' `
        -Visual 'Hero KPI: Copilot licenses purchased and assigned. Usage bar (assigned vs active). Right: ROI callout — time saved per user, cost per seat.' `
        -BigMessage "$copAssigned of $copPurchased Copilot licenses assigned — $copUnused unused" `
        -Bullets @(
            "Copilot for M365 licenses purchased: **$copPurchased**"
            "Licenses assigned to users: **$copAssigned**"
            "Potentially unused: **$copUnused**"
            "Copilot requires M365 E3/E5 or Business Premium as a base license"
            "Action: $copAction"
        ) `
        -SpeakerNotes @(
            "Copilot for Microsoft 365 is one of the highest per-seat costs in the M365 family.",
            "Unused Copilot licenses represent significant avoidable spend.",
            "True usage data (prompts, active users) requires the Microsoft 365 admin usage reports — available in the admin centre.",
            "If the client has not yet run a Copilot adoption programme, this is an opportunity to position RTH's managed adoption service.",
            "Action: $copAction"
        ) `
        -Takeaway "Copilot ROI is maximised through structured adoption. Unassigned or unactivated licenses are an immediate cost saving opportunity."

    # ── Slide 11: Top Recommendations ────────────────────────────────────────

    Add-Slide -L $lines -Num 11 -Title 'Top Recommendations' `
        -Visual 'Numbered list with priority colour coding (red = P1, amber = P2, blue = P3+). Each row: priority, category, action, effort tag, impact tag.' `
        -BigMessage "$($recommendations.Count) actionable recommendations identified" `
        -Bullets $topRecLines `
        -SpeakerNotes @(
            "Walk the client through each recommendation, pausing to agree on priority and ownership.",
            "Effort Low = days, Medium = weeks, High = weeks to months.",
            "Impact High = material cost saving or significant risk reduction.",
            "Assign an owner for each item before leaving the meeting.",
            "Follow up with a written action plan within 48 hours."
        ) `
        -Takeaway "These recommendations are data-driven and ranked by priority. Completing P1 and P2 items will deliver the most value in the shortest time."

    # ── Slide 12: Next Quarter Action Plan ───────────────────────────────────

    $p1Recs = @($recommendations | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'priority') -le 2 })
    $actionLines = if ($p1Recs.Count -gt 0) {
        $p1Recs | ForEach-Object {
            $t = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'title')
            $o = Nvl (Get-TenantReviewProperty -InputObject $_ -Name 'owner')
            "[ ] **$t** — Owner: $o"
        }
    } else {
        @('[ ] Review all findings and assign owners', '[ ] Schedule follow-up session', '[ ] Monitor tenant health metrics monthly')
    }

    Add-Slide -L $lines -Num 12 -Title 'Next Quarter Action Plan' `
        -Visual 'Kanban-style columns: Immediate (P1) / This Quarter (P2) / Roadmap (P3+). Each item is a card with owner and target due date.' `
        -BigMessage "Commit to actions before this meeting closes" `
        -Bullets $actionLines `
        -SpeakerNotes @(
            "Use this slide to close the meeting with clear commitments.",
            "Each action should have a named owner and an agreed delivery date.",
            "RTH to send a written summary of agreed actions within 48 hours.",
            "Schedule the next quarterly review now while everyone is in the room.",
            "Offer to track action completion progress as a managed service engagement."
        ) `
        -Takeaway "The value of a review is in the actions it drives. Leave today with clear commitments and owners."

    # ── Slide 13: Appendix / Data Sources ────────────────────────────────────

    $coverageLines = $Datasets.Keys | ForEach-Object {
        $ds     = $Datasets[$_]
        $sum    = Get-TenantReviewProperty -InputObject $ds -Name 'summary'
        $skip   = Get-TenantReviewProperty -InputObject $sum -Name 'skipped'
        $warns  = @(Get-TenantReviewProperty -InputObject $ds -Name 'warnings')
        $status = if ($skip) { 'Skipped' } elseif ($warns.Count) { "Collected ($($warns.Count) warning$(if($warns.Count -ne 1){'s'}))" } else { 'Collected' }
        "**$_**: $status"
    }

    Add-Slide -L $lines -Num 13 -Title 'Appendix — Data Sources & Coverage' `
        -Visual 'Two-column layout. Left: data source coverage table. Right: key definitions and notes.' `
        -BigMessage "All data collected live from the $orgName Microsoft 365 tenant" `
        -Bullets $coverageLines `
        -SpeakerNotes @(
            "This slide documents where data came from and any gaps.",
            "Skipped datasets were not available due to licensing, permissions, or configuration at the time of collection.",
            "For a complete review, ensure the service account has at minimum: Global Reader, Exchange View-Only Admin, SharePoint Admin, Teams Admin.",
            "The full technical data tables are in the TenantReviewReport.html file provided alongside this deck.",
            "Report generated: $generatedAt"
        ) `
        -Takeaway "Full detail tables are in the accompanying HTML report."

    # ── write output ─────────────────────────────────────────────────────────

    $newline = [Environment]::NewLine
    Set-Content -Path $outlinePath -Value ($lines -join $newline) -Encoding UTF8
    Write-Host "Deck outline created: $outlinePath"
}

