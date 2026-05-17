function New-TenantReviewReport {
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

    $reportPath = Join-Path $OutputPath 'TenantReviewReport.md'
    $htmlPath   = Join-Path $OutputPath 'TenantReviewReport.html'
    $generatedAt = (Get-Date).ToString('MMMM d, yyyy')

    # ── data helpers ─────────────────────────────────────────────────────────

    function Escape-Html {
        param([object]$Value)
        if ($null -eq $Value) { return '' }
        return [System.Net.WebUtility]::HtmlEncode($Value.ToString())
    }

    function Nvl {
        param([object]$Value, [string]$Default = 'N/A')
        if ($null -eq $Value -or $Value.ToString() -eq '') { return $Default }
        return $Value.ToString()
    }

    function Get-StatusCss {
        param([string]$Status)
        switch ($Status) {
            'Good'              { return 'status-good' }
            'Watch'             { return 'status-watch' }
            'ActionRecommended' { return 'status-action' }
            'NotAvailable'      { return 'status-na' }
            default             { return 'status-info' }
        }
    }

    function Get-StatusLabel {
        param([string]$Status)
        switch ($Status) {
            'Good'              { return 'Healthy' }
            'Watch'             { return 'Watch' }
            'ActionRecommended' { return 'Action Recommended' }
            'NotAvailable'      { return 'Not Available' }
            default             { return 'Info' }
        }
    }

    function Get-BadgeCss {
        param([string]$Status)
        switch ($Status) {
            'Good'              { return 'badge-good' }
            'Watch'             { return 'badge-watch' }
            'ActionRecommended' { return 'badge-action' }
            'NotAvailable'      { return 'badge-na' }
            default             { return 'badge-info' }
        }
    }

    function Get-EffortCss {
        param([string]$Effort)
        switch ($Effort) { 'Low' { 'effort-low' } 'High' { 'effort-high' } default { 'effort-med' } }
    }

    # ── section data shortcuts ──────────────────────────────────────────────

    $sections        = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'sections')
    $kpis            = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'kpis')
    $recommendations = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'recommendations')
    $execBullets     = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'executiveSummaryBullets')
    $healthStatement = Nvl (Get-TenantReviewProperty -InputObject $Narrative -Name 'tenantHealthStatement') 'Review the findings in this report and act on the highlighted recommendations.'

    $attentionItems = @(Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'attentionItems')
    $tenantSummary  = Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'summary'
    $userSummary    = Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'summary'
    $licSummary     = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
    $anaSummary     = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
    $mbxSummary     = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
    $spSummary      = Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'summary'
    $teamsSummary   = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
    $deviceSummary  = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
    $copilotSummary = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'

    $orgName        = Nvl (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'organizationName') $TenantName

    # ── find narrative section helper ────────────────────────────────────────

    function Get-Section { param([string]$Dataset) $sections | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'dataset') -eq $Dataset } | Select-Object -First 1 }

    # ════════════════════════════════════════════════════════════════════════
    # MARKDOWN REPORT
    # ════════════════════════════════════════════════════════════════════════

    $md = [System.Collections.Generic.List[string]]::new()

    $md.Add("# Microsoft 365 Tenant Review")
    $md.Add("")
    $md.Add("**Organisation:** $orgName")
    $md.Add("**Review period:** $ReviewPeriod")
    $md.Add("**Prepared by:** RTH Tech Services")
    $md.Add("**Generated:** $generatedAt")
    $md.Add("")
    $md.Add("---")
    $md.Add("")

    # Executive Summary
    $md.Add("## Executive Summary")
    $md.Add("")
    $md.Add("> $healthStatement")
    $md.Add("")
    foreach ($bullet in $execBullets) { $md.Add("- $bullet") }
    $md.Add("")

    # KPI Summary Table
    $md.Add("## Key Metrics at a Glance")
    $md.Add("")
    $md.Add("| Area | Metric | Value | Status |")
    $md.Add("| --- | --- | ---: | --- |")
    foreach ($kpi in $kpis) {
        $area  = Nvl (Get-TenantReviewProperty -InputObject $kpi -Name 'area')
        $label = Nvl (Get-TenantReviewProperty -InputObject $kpi -Name 'label')
        $val   = Nvl (Get-TenantReviewProperty -InputObject $kpi -Name 'value')
        $stat  = Nvl (Get-TenantReviewProperty -InputObject $kpi -Name 'status') 'Informational'
        $md.Add("| $area | $label | $val | $(Get-StatusLabel $stat) |")
    }
    $md.Add("")

    # Narrative sections
    $sectionOrder = @('TenantOverview','LicenseInventory','LicenseUserAnalysis','UserInventory','MailboxInventory','SharePoint','Teams','Devices','Copilot')
    $sectionTitles = @{
        'TenantOverview'      = 'Tenant Overview'
        'LicenseInventory'    = 'Cost & Licensing Review'
        'LicenseUserAnalysis' = 'License Optimisation'
        'UserInventory'       = 'Identity & User Review'
        'MailboxInventory'    = 'Exchange & Mail Flow'
        'SharePoint'          = 'SharePoint & OneDrive'
        'Teams'               = 'Teams Collaboration'
        'Devices'             = 'Devices & Endpoints'
        'Copilot'             = 'Copilot Review'
    }
    foreach ($key in $sectionOrder) {
        $sec = Get-Section -Dataset $key
        if (-not $sec) { continue }
        $title   = if ($sectionTitles[$key]) { $sectionTitles[$key] } else { $key }
        $headline = Nvl (Get-TenantReviewProperty -InputObject $sec -Name 'headline')
        $plain    = Nvl (Get-TenantReviewProperty -InputObject $sec -Name 'plainEnglish')
        $impact   = Nvl (Get-TenantReviewProperty -InputObject $sec -Name 'businessImpact')
        $action   = Nvl (Get-TenantReviewProperty -InputObject $sec -Name 'recommendedAction')
        $status   = Nvl (Get-TenantReviewProperty -InputObject $sec -Name 'status') 'Informational'
        $md.Add("## $title")
        $md.Add("")
        $md.Add("**$headline** · *$(Get-StatusLabel $status)*")
        $md.Add("")
        $md.Add($plain)
        $md.Add("")
        $md.Add("**Why it matters:** $impact")
        $md.Add("")
        $md.Add("**Recommended action:** $action")
        $md.Add("")
    }

    # Recommendations
    $md.Add("## Top Recommendations")
    $md.Add("")
    if ($recommendations.Count -eq 0) {
        $md.Add("No specific recommendations were generated from the available data.")
    } else {
        $md.Add("| # | Category | Recommendation | Effort | Impact | Status |")
        $md.Add("| --- | --- | --- | --- | --- | --- |")
        foreach ($rec in $recommendations) {
            $pri    = Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'priority')
            $cat    = Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'category')
            $title  = Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'title')
            $effort = Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'effort')
            $impact = Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'impact')
            $stat   = Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'status') 'Watch'
            $md.Add("| $pri | $cat | $title | $effort | $impact | $(Get-StatusLabel $stat) |")
        }
    }
    $md.Add("")

    # Data coverage
    $md.Add("## Data Coverage")
    $md.Add("")
    $md.Add("This report was generated from the following data sources. Where data was unavailable due to licensing or permissions, sections are marked accordingly.")
    $md.Add("")
    $md.Add("| Dataset | Status |")
    $md.Add("| --- | --- |")
    foreach ($key in $Datasets.Keys) {
        $ds  = $Datasets[$key]
        $sum = Get-TenantReviewProperty -InputObject $ds -Name 'summary'
        $skipped = Get-TenantReviewProperty -InputObject $sum -Name 'skipped'
        $warnings = @(Get-TenantReviewProperty -InputObject $ds -Name 'warnings')
        $status = if ($skipped) { 'Skipped' } elseif ($warnings.Count -gt 0) { "Collected with $($warnings.Count) warning$(if($warnings.Count -ne 1){'s'})" } else { 'Collected' }
        $md.Add("| $key | $status |")
    }
    $md.Add("")

    # Appendix
    $md.Add("## Appendix")
    $md.Add("")

    # License SKUs
    $licItems = @(Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'items')
    if ($licItems.Count -gt 0) {
        $md.Add("### License SKUs")
        $md.Add("")
        $md.Add("| SKU | Purchased | Assigned | Unused | Monthly Cost |")
        $md.Add("| --- | ---: | ---: | ---: | ---: |")
        foreach ($sku in ($licItems | Sort-Object { -[int](Get-TenantReviewProperty -InputObject $_ -Name 'purchasedUnits') })) {
            $name    = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'displayName') (Get-TenantReviewProperty -InputObject $sku -Name 'skuPartNumber')
            $pur     = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'purchasedUnits')
            $asgn    = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'assignedUnits')
            $unus    = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'unusedUnits')
            $mCost   = Get-TenantReviewProperty -InputObject $sku -Name 'estimatedMonthlyCost'
            $costStr = if ($null -ne $mCost -and [double]$mCost -gt 0) { "CAD `$$([Math]::Round([double]$mCost,2))" } else { '—' }
            $md.Add("| $name | $pur | $asgn | $unus | $costStr |")
        }
        $md.Add("")
    }

    # Mailboxes with forwarding
    $mbxItems = @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'items')
    $fwdItems = @($mbxItems | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'forwardingEnabled') -eq $true })
    if ($fwdItems.Count -gt 0) {
        $md.Add("### Mailboxes with Forwarding")
        $md.Add("")
        $md.Add("| Display Name | Email | Type | External? |")
        $md.Add("| --- | --- | --- | --- |")
        foreach ($mbx in $fwdItems) {
            $dname  = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'displayName')
            $email  = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'primarySmtpAddress')
            $type   = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'recipientTypeDetails')
            $extFwd = Get-TenantReviewProperty -InputObject $mbx -Name 'externalForwardingSuspected'
            $extStr = if ($extFwd) { 'Yes' } else { 'No' }
            $md.Add("| $dname | $email | $type | $extStr |")
        }
        $md.Add("")
    }

    # Top SharePoint sites
    $spItems = @(Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'items')
    if ($spItems.Count -gt 0) {
        $md.Add("### Largest SharePoint Sites")
        $md.Add("")
        $md.Add("| Site | Storage Used |")
        $md.Add("| --- | ---: |")
        foreach ($site in ($spItems | Sort-Object { -[double](Get-TenantReviewProperty -InputObject $_ -Name 'storageUsageGB') } | Select-Object -First 10)) {
            $title = Nvl (Get-TenantReviewProperty -InputObject $site -Name 'title')
            $gb    = [double](Get-TenantReviewProperty -InputObject $site -Name 'storageUsageGB')
            $md.Add("| $title | $('{0:N2}' -f $gb) GB |")
        }
        $md.Add("")
    }

    # Teams
    $teamsItems = @(Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'items')
    if ($teamsItems.Count -gt 0) {
        $md.Add("### Teams Inventory")
        $md.Add("")
        $md.Add("| Team | Visibility | Last Activity | Inactive? |")
        $md.Add("| --- | --- | --- | --- |")
        foreach ($team in $teamsItems) {
            $tname    = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'displayName')
            $vis      = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'visibility')
            $lastDate = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'lastActivityDate') '—'
            $inactive = Get-TenantReviewProperty -InputObject $team -Name 'inactive'
            $inactStr = if ($inactive -eq $true) { 'Yes' } elseif ($inactive -eq $false) { 'No' } else { '—' }
            $md.Add("| $tname | $vis | $lastDate | $inactStr |")
        }
        $md.Add("")
    }

    # Stale devices
    $devItems   = @(Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'items')
    $staleDevs  = @($devItems | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'isStale') -eq $true })
    if ($staleDevs.Count -gt 0) {
        $md.Add("### Stale Devices")
        $md.Add("")
        $md.Add("| Device | OS | Last Sign-In (days ago) | Enabled |")
        $md.Add("| --- | --- | ---: | --- |")
        foreach ($dev in $staleDevs) {
            $dname   = Nvl (Get-TenantReviewProperty -InputObject $dev -Name 'displayName')
            $os      = Nvl (Get-TenantReviewProperty -InputObject $dev -Name 'operatingSystem')
            $days    = Nvl (Get-TenantReviewProperty -InputObject $dev -Name 'daysSinceLastSignIn') '—'
            $enabled = Get-TenantReviewProperty -InputObject $dev -Name 'accountEnabled'
            $md.Add("| $dname | $os | $days | $(if($enabled){'Yes'} else {'No'}) |")
        }
        $md.Add("")
    }

    # Copilot licensed users
    $copUsers = @(Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'licensedUsers')
    if ($copUsers.Count -gt 0) {
        $md.Add("### Copilot Licensed Users")
        $md.Add("")
        $md.Add("| User | Stale? |")
        $md.Add("| --- | --- |")
        foreach ($cu in $copUsers) {
            $upn   = Nvl (Get-TenantReviewProperty -InputObject $cu -Name 'userPrincipalName')
            $stale = Get-TenantReviewProperty -InputObject $cu -Name 'isStale'
            $md.Add("| $upn | $(if($stale){'Yes'} else {'No'}) |")
        }
        $md.Add("")
    }

    $newline     = [Environment]::NewLine
    $markdownText = $md -join $newline
    Set-Content -Path $reportPath -Value $markdownText -Encoding UTF8

    # ════════════════════════════════════════════════════════════════════════
    # HTML REPORT
    # ════════════════════════════════════════════════════════════════════════

    function Build-KpiCardHtml {
        param([object]$Kpi)
        $val   = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'value'))
        $label = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'label'))
        $desc  = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'description'))
        $stat  = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'status') 'Informational'
        $css   = Get-StatusCss $stat
        $badge = Get-BadgeCss $stat
        $lbl   = Get-StatusLabel $stat
        return "<div class=`"kpi-card $css`"><div class=`"kpi-value`">$val</div><div class=`"kpi-label`">$label</div><div class=`"kpi-description`">$desc</div><span class=`"kpi-badge $badge`">$lbl</span></div>"
    }

    function Build-NarrativeCardHtml {
        param([object]$Section, [string]$Icon = '')
        $headline = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'headline'))
        $plain    = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'plainEnglish'))
        $impact   = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'businessImpact'))
        $action   = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'recommendedAction'))
        $stat     = Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'status') 'Informational'
        $css      = Get-StatusCss $stat
        $badge    = Get-BadgeCss $stat
        $lbl      = Get-StatusLabel $stat
        $iconHtml = if ($Icon) { "<span class=`"nar-icon`">$Icon</span>" } else { '' }
        return @"
<div class="narrative-card $css">
  $iconHtml<span class="narrative-badge $(Get-BadgeCss $stat)">$lbl</span>
  <div class="narrative-headline">$headline</div>
  <div class="narrative-plain">$plain</div>
  <div class="narrative-impact"><strong>Why it matters:</strong> $impact</div>
  <div class="narrative-action"><strong>Recommended action:</strong> $action</div>
</div>
"@
    }

    function Build-RecCardHtml {
        param([object]$Rec)
        $pri    = Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'priority') '—'
        $title  = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'title'))
        $why    = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'why'))
        $owner  = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'owner'))
        $effort = Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'effort') 'Medium'
        $impact = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'impact'))
        $cat    = Escape-Html (Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'category'))
        $stat   = Nvl (Get-TenantReviewProperty -InputObject $Rec -Name 'status') 'Watch'
        $efCss  = Get-EffortCss $effort
        $stBadge = Get-BadgeCss $stat
        $stLabel = Get-StatusLabel $stat
        return @"
<div class="rec-card">
  <div class="rec-priority">P$pri</div>
  <div class="rec-body">
    <div class="rec-title">$title</div>
    <div class="rec-why">$why</div>
    <div class="rec-meta">
      <span class="rec-tag">$cat</span>
      <span class="rec-tag $efCss">Effort: $effort</span>
      <span class="rec-tag impact-high">Impact: $impact</span>
      <span class="rec-tag $stBadge">$(Get-StatusLabel $stat)</span>
      <span class="rec-tag">Owner: $owner</span>
    </div>
  </div>
</div>
"@
    }

    function Build-TableRowHtml {
        param([string[]]$Cells)
        $tds = ($Cells | ForEach-Object { "<td>$(Escape-Html $_)</td>" }) -join ''
        return "<tr>$tds</tr>"
    }

    function Build-SectionHeaderHtml {
        param([string]$Icon, [string]$Title, [string]$Subtitle = '')
        $sub = if ($Subtitle) { "<div class=`"section-subtitle`">$(Escape-Html $Subtitle)</div>" } else { '' }
        return "<div class=`"section-header`"><span class=`"section-icon`">$Icon</span><div><h2>$(Escape-Html $Title)</h2>$sub</div></div>"
    }

    # ── HTML body construction ───────────────────────────────────────────────

    $kpiCardsHtml  = ($kpis | ForEach-Object { Build-KpiCardHtml -Kpi $_ }) -join "`n"
    $recCardsHtml  = if ($recommendations.Count -gt 0) {
        ($recommendations | ForEach-Object { Build-RecCardHtml -Rec $_ }) -join "`n"
    } else {
        '<p>No specific recommendations were generated from the available data.</p>'
    }
    $execBulletsHtml = ($execBullets | ForEach-Object { "<li>$(Escape-Html $_)</li>" }) -join "`n"

    # Section icons
    $sectionIcons = @{
        'TenantOverview'      = '🏢'
        'LicenseInventory'    = '💳'
        'LicenseUserAnalysis' = '📊'
        'UserInventory'       = '👥'
        'MailboxInventory'    = '📧'
        'SharePoint'          = '📁'
        'Teams'               = '💬'
        'Devices'             = '💻'
        'Copilot'             = '✨'
    }

    # Build each narrative section HTML
    $narrativeSectionsHtml = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $sectionOrder) {
        $sec   = Get-Section -Dataset $key
        if (-not $sec) { continue }
        $title = if ($sectionTitles[$key]) { $sectionTitles[$key] } else { $key }
        $icon  = if ($sectionIcons[$key]) { $sectionIcons[$key] } else { '' }

        # Build KPI subsection for this area
        $areaName = switch ($key) {
            'TenantOverview'      { '' }
            'LicenseInventory'    { 'Licensing' }
            'LicenseUserAnalysis' { 'Licensing' }
            'UserInventory'       { 'Users' }
            'MailboxInventory'    { 'Exchange' }
            'SharePoint'          { 'SharePoint' }
            'Teams'               { 'Teams' }
            'Devices'             { 'Devices' }
            'Copilot'             { 'Copilot' }
            default               { '' }
        }
        $sectionKpis = if ($areaName) { @($kpis | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'area') -eq $areaName }) } else { @() }
        # Avoid duplicating Licensing KPIs across two sections
        if ($key -eq 'LicenseUserAnalysis') { $sectionKpis = @() }
        $kpiSubsection = if ($sectionKpis.Count -gt 0) {
            $cards = ($sectionKpis | ForEach-Object { Build-KpiCardHtml -Kpi $_ }) -join "`n"
            "<div class=`"kpi-grid`">$cards</div>"
        } else { '' }

        $narrativeHtml = Build-NarrativeCardHtml -Section $sec -Icon $icon

        $narrativeSectionsHtml.Add(@"
<div class="report-section">
  $(Build-SectionHeaderHtml -Icon $icon -Title $title)
  <div class="section-body">
    $kpiSubsection
    $narrativeHtml
  </div>
</div>
"@)
    }

    # Appendix tables
    $appendixHtml = [System.Collections.Generic.List[string]]::new()
    $appendixHtml.Add('<div class="report-section">')
    $appendixHtml.Add((Build-SectionHeaderHtml -Icon '📎' -Title 'Appendix' -Subtitle 'Supporting data tables for technical reviewers'))
    $appendixHtml.Add('<div class="section-body">')

    # License SKUs table
    $licItems = @(Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'items')
    if ($licItems.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>License SKUs</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>SKU</th><th>Status</th><th>Purchased</th><th>Assigned</th><th>Unused</th><th>Monthly Cost</th></tr></thead><tbody>')
        foreach ($sku in ($licItems | Sort-Object { -[int](Get-TenantReviewProperty -InputObject $_ -Name 'purchasedUnits') })) {
            $name    = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'displayName') (Get-TenantReviewProperty -InputObject $sku -Name 'skuPartNumber')
            $status  = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'capabilityStatus')
            $pur     = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'purchasedUnits')
            $asgn    = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'assignedUnits')
            $unus    = Nvl (Get-TenantReviewProperty -InputObject $sku -Name 'unusedUnits')
            $mCost   = Get-TenantReviewProperty -InputObject $sku -Name 'estimatedMonthlyCost'
            $costStr = if ($null -ne $mCost -and [double]$mCost -gt 0) { "CAD `$$([Math]::Round([double]$mCost,2))" } else { '—' }
            $appendixHtml.Add((Build-TableRowHtml -Cells @($name, $status, $pur, $asgn, $unus, $costStr)))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    # Users requiring review
    $userItems = @(Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'items')
    $reviewUsers = @($userItems | Where-Object {
        (Get-TenantReviewProperty -InputObject $_ -Name 'isLicensedAndDisabled') -eq $true -or
        (Get-TenantReviewProperty -InputObject $_ -Name 'isLicensedAndStale') -eq $true -or
        (Get-TenantReviewProperty -InputObject $_ -Name 'isGuest') -eq $true
    })
    if ($reviewUsers.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>Users Requiring Review</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>User</th><th>Type</th><th>Enabled</th><th>Licensed</th><th>Flag</th></tr></thead><tbody>')
        foreach ($u in $reviewUsers) {
            $upn      = Nvl (Get-TenantReviewProperty -InputObject $u -Name 'userPrincipalName')
            $utype    = Nvl (Get-TenantReviewProperty -InputObject $u -Name 'userType')
            $enabled  = if ((Get-TenantReviewProperty -InputObject $u -Name 'accountEnabled') -eq $true) { 'Yes' } else { 'No' }
            $licensed = if ((Get-TenantReviewProperty -InputObject $u -Name 'isLicensed') -eq $true) { 'Yes' } else { 'No' }
            $flags    = @()
            if ((Get-TenantReviewProperty -InputObject $u -Name 'isLicensedAndDisabled') -eq $true) { $flags += 'Disabled+Licensed' }
            if ((Get-TenantReviewProperty -InputObject $u -Name 'isLicensedAndStale') -eq $true) { $flags += 'Stale+Licensed' }
            if ((Get-TenantReviewProperty -InputObject $u -Name 'isGuest') -eq $true) { $flags += 'Guest' }
            $appendixHtml.Add((Build-TableRowHtml -Cells @($upn, $utype, $enabled, $licensed, ($flags -join ', '))))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    # Mailboxes with forwarding
    $mbxItems = @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'items')
    $fwdMbx   = @($mbxItems | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'forwardingEnabled') -eq $true })
    if ($fwdMbx.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>Mailboxes with Forwarding</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>Display Name</th><th>Email</th><th>Type</th><th>External?</th><th>Forward To</th></tr></thead><tbody>')
        foreach ($mbx in $fwdMbx) {
            $dname   = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'displayName')
            $email   = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'primarySmtpAddress')
            $mtype   = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'recipientTypeDetails')
            $extFwd  = if ((Get-TenantReviewProperty -InputObject $mbx -Name 'externalForwardingSuspected') -eq $true) { 'Yes' } else { 'No' }
            $fwdTo   = Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'forwardingSmtpAddress') (Nvl (Get-TenantReviewProperty -InputObject $mbx -Name 'forwardingAddress') '—')
            $appendixHtml.Add((Build-TableRowHtml -Cells @($dname, $email, $mtype, $extFwd, $fwdTo)))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    # Largest SharePoint sites
    $spItems = @(Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'items')
    if ($spItems.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>Largest SharePoint Sites</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>Site</th><th>Storage Used</th><th>Last Modified</th><th>Files</th></tr></thead><tbody>')
        foreach ($site in ($spItems | Sort-Object { -[double](Get-TenantReviewProperty -InputObject $_ -Name 'storageUsageGB') } | Select-Object -First 15)) {
            $stitle  = Nvl (Get-TenantReviewProperty -InputObject $site -Name 'title')
            $gb      = [double](Get-TenantReviewProperty -InputObject $site -Name 'storageUsageGB')
            $modDate = Nvl (Get-TenantReviewProperty -InputObject $site -Name 'lastContentModifiedDate') '—'
            $fcount  = Nvl (Get-TenantReviewProperty -InputObject $site -Name 'fileCount') '—'
            $appendixHtml.Add((Build-TableRowHtml -Cells @($stitle, "$('{0:N2}' -f $gb) GB", $modDate, $fcount)))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    # Teams
    $teamsItems = @(Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'items')
    if ($teamsItems.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>Teams Inventory</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>Team</th><th>Visibility</th><th>Created</th><th>Last Activity</th><th>Inactive?</th></tr></thead><tbody>')
        foreach ($team in $teamsItems) {
            $tname    = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'displayName')
            $vis      = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'visibility')
            $created  = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'createdDateTime') '—'
            $lastAct  = Nvl (Get-TenantReviewProperty -InputObject $team -Name 'lastActivityDate') '—'
            $inactive = Get-TenantReviewProperty -InputObject $team -Name 'inactive'
            $inactStr = if ($inactive -eq $true) { 'Yes' } elseif ($inactive -eq $false) { 'No' } else { '—' }
            $appendixHtml.Add((Build-TableRowHtml -Cells @($tname, $vis, $created, $lastAct, $inactStr)))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    # Stale devices
    $devItems  = @(Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'items')
    $staleDev  = @($devItems | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'isStale') -eq $true })
    if ($staleDev.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>Stale Devices</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>Device</th><th>OS</th><th>Days Since Sign-In</th><th>Enabled</th><th>Intune</th></tr></thead><tbody>')
        foreach ($dev in $staleDev) {
            $dname   = Nvl (Get-TenantReviewProperty -InputObject $dev -Name 'displayName')
            $os      = Nvl (Get-TenantReviewProperty -InputObject $dev -Name 'operatingSystem')
            $days    = Nvl (Get-TenantReviewProperty -InputObject $dev -Name 'daysSinceLastSignIn') '—'
            $enabled = if ((Get-TenantReviewProperty -InputObject $dev -Name 'accountEnabled') -eq $true) { 'Yes' } else { 'No' }
            $intune  = if ((Get-TenantReviewProperty -InputObject $dev -Name 'isManaged') -eq $true) { 'Yes' } else { 'No' }
            $appendixHtml.Add((Build-TableRowHtml -Cells @($dname, $os, $days, $enabled, $intune)))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    # Copilot licensed users
    $copUsers = @(Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'licensedUsers')
    if ($copUsers.Count -gt 0) {
        $appendixHtml.Add('<div class="appendix-section"><h3>Copilot Licensed Users</h3>')
        $appendixHtml.Add('<table class="data-table"><thead><tr><th>User</th><th>Display Name</th><th>Stale?</th></tr></thead><tbody>')
        foreach ($cu in $copUsers) {
            $upn    = Nvl (Get-TenantReviewProperty -InputObject $cu -Name 'userPrincipalName')
            $dname  = Nvl (Get-TenantReviewProperty -InputObject $cu -Name 'displayName')
            $stale  = if ((Get-TenantReviewProperty -InputObject $cu -Name 'isStale') -eq $true) { 'Yes' } else { 'No' }
            $appendixHtml.Add((Build-TableRowHtml -Cells @($upn, $dname, $stale)))
        }
        $appendixHtml.Add('</tbody></table></div>')
    }

    $appendixHtml.Add('</div></div>')

    # Data coverage table
    $coverageHtml = [System.Collections.Generic.List[string]]::new()
    $coverageIcons = @{ 'TenantOverview'='🏢'; 'LicenseInventory'='💳'; 'LicenseUserAnalysis'='📊'; 'UserInventory'='👥'; 'MailboxInventory'='📧'; 'SharePoint'='📁'; 'Teams'='💬'; 'Devices'='💻'; 'Copilot'='✨' }
    $coverageHtml.Add('<div class="coverage-grid">')
    foreach ($key in $Datasets.Keys) {
        $ds     = $Datasets[$key]
        $sum    = Get-TenantReviewProperty -InputObject $ds -Name 'summary'
        $skip   = Get-TenantReviewProperty -InputObject $sum -Name 'skipped'
        $warns  = @(Get-TenantReviewProperty -InputObject $ds -Name 'warnings')
        $ico    = if ($coverageIcons[$key]) { $coverageIcons[$key] } else { '📄' }
        $stat   = if ($skip)            { '<span style="color:#767779">Skipped</span>' }
                  elseif ($warns.Count) { '<span style="color:#f39c12">Collected with warnings</span>' }
                  else                  { '<span style="color:#27ae60">Collected</span>' }
        $coverageHtml.Add("<div class=`"coverage-item`"><div class=`"coverage-icon`">$ico</div><div><div class=`"coverage-label`">$(Escape-Html $key)</div><div class=`"coverage-status`">$stat</div></div></div>")
    }
    $coverageHtml.Add('</div>')

    # ── full HTML assembly ───────────────────────────────────────────────────

    $css = @'
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f7f8fb; color: #1a1a2e; line-height: 1.6; font-size: 15px; }

/* Cover */
.cover { background: linear-gradient(135deg, #5a2f61 0%, #6f28db 55%, #4fa7db 100%); min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 60px 40px; text-align: center; color: white; page-break-after: always; }
.cover-inner { max-width: 700px; }
.cover-eyebrow { font-size: 0.85em; letter-spacing: 3px; text-transform: uppercase; opacity: 0.7; margin-bottom: 28px; }
.cover-title { font-size: 3.2em; font-weight: 800; margin-bottom: 0.15em; line-height: 1.1; }
.cover-subtitle { font-size: 1.6em; font-weight: 300; opacity: 0.9; margin-bottom: 36px; }
.cover-badges { display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; margin-bottom: 40px; }
.cover-badge { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.3); border-radius: 6px; padding: 8px 18px; font-size: 0.85em; letter-spacing: 0.4px; }
.cover-prepared { font-size: 0.82em; opacity: 0.65; margin-top: 32px; letter-spacing: 1px; text-transform: uppercase; }
.cover-divider { width: 60px; height: 3px; background: rgba(255,255,255,0.4); margin: 0 auto 28px; border-radius: 2px; }

/* Page wrapper */
.page { max-width: 1120px; margin: 0 auto; padding: 56px 40px; }

/* Section structure */
.report-section { margin-bottom: 52px; }
.section-header { display: flex; align-items: center; gap: 16px; background: linear-gradient(90deg, #5a2f61, #6f28db); color: white; padding: 18px 28px; border-radius: 8px 8px 0 0; }
.section-header h2 { font-size: 1.35em; font-weight: 700; }
.section-subtitle { font-size: 0.82em; opacity: 0.8; margin-top: 2px; }
.section-icon { font-size: 1.7em; }
.section-body { background: white; padding: 28px 32px; border: 1px solid #e2e8f0; border-top: none; border-radius: 0 0 8px 8px; }

/* Executive summary */
.exec-health { background: linear-gradient(90deg, #ebf5fb, #f5eeff); border-left: 4px solid #4fa7db; border-radius: 4px; padding: 18px 22px; margin-bottom: 22px; font-size: 1.05em; color: #1a3a4a; }
.exec-bullets { list-style: none; padding: 0; display: grid; gap: 10px; }
.exec-bullets li { padding: 13px 16px 13px 44px; position: relative; background: #f9f9fc; border-radius: 6px; border-left: 3px solid #6f28db; font-size: 0.94em; color: #333; }
.exec-bullets li::before { content: '→'; position: absolute; left: 16px; color: #6f28db; font-weight: 700; }

/* KPI grid */
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap: 14px; margin-bottom: 28px; }
.kpi-card { background: white; border-radius: 10px; padding: 18px; border: 1px solid #e2e8f0; border-top: 4px solid #767779; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
.kpi-card.status-good    { border-top-color: #27ae60; }
.kpi-card.status-watch   { border-top-color: #f39c12; }
.kpi-card.status-action  { border-top-color: #e74c3c; }
.kpi-card.status-info    { border-top-color: #4fa7db; }
.kpi-card.status-na      { border-top-color: #aaa; }
.kpi-value { font-size: 2.2em; font-weight: 800; color: #5a2f61; line-height: 1.1; margin-bottom: 4px; }
.kpi-label { font-size: 0.75em; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: #767779; margin-bottom: 8px; }
.kpi-description { font-size: 0.8em; color: #555; line-height: 1.4; margin-bottom: 12px; }
.kpi-badge { display: inline-block; font-size: 0.68em; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; padding: 3px 10px; border-radius: 20px; }
.badge-good   { background: #d5f5e3; color: #1e8449; }
.badge-watch  { background: #fef9e7; color: #9a6b0e; }
.badge-action { background: #fde8e8; color: #c0392b; }
.badge-info   { background: #ebf5fb; color: #1a6a9a; }
.badge-na     { background: #f2f3f4; color: #566573; }

/* Narrative cards */
.narrative-card { border: 1px solid #e2e8f0; border-radius: 8px; padding: 22px 24px; margin-bottom: 16px; position: relative; }
.narrative-card.status-good   { border-left: 4px solid #27ae60; }
.narrative-card.status-watch  { border-left: 4px solid #f39c12; }
.narrative-card.status-action { border-left: 4px solid #e74c3c; }
.narrative-card.status-info   { border-left: 4px solid #4fa7db; }
.narrative-card.status-na     { border-left: 4px solid #aaa; }
.narrative-badge { float: right; margin-left: 8px; margin-bottom: 4px; font-size: 0.7em; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; padding: 3px 10px; border-radius: 20px; }
.nar-icon { float: left; margin-right: 10px; font-size: 1.3em; }
.narrative-headline { font-size: 1.08em; font-weight: 700; color: #2c3e50; margin-bottom: 10px; clear: both; }
.narrative-plain { color: #444; margin-bottom: 10px; font-size: 0.93em; }
.narrative-impact { font-size: 0.88em; background: #f8f9fb; padding: 8px 12px; border-radius: 4px; margin-bottom: 7px; color: #555; }
.narrative-impact strong { color: #5a2f61; }
.narrative-action { font-size: 0.88em; background: #eef3ff; padding: 8px 12px; border-radius: 4px; color: #2c3e50; }
.narrative-action strong { color: #4fa7db; }

/* Recommendations */
.rec-card { border: 1px solid #e2e8f0; border-radius: 8px; margin-bottom: 14px; overflow: hidden; display: flex; box-shadow: 0 1px 4px rgba(0,0,0,0.04); }
.rec-priority { background: linear-gradient(135deg, #5a2f61, #6f28db); color: white; font-weight: 800; font-size: 1.15em; min-width: 58px; display: flex; align-items: center; justify-content: center; padding: 16px; }
.rec-body { padding: 14px 18px; flex: 1; }
.rec-title { font-size: 0.97em; font-weight: 700; color: #2c3e50; margin-bottom: 6px; }
.rec-why { font-size: 0.86em; color: #555; margin-bottom: 10px; line-height: 1.5; }
.rec-meta { display: flex; flex-wrap: wrap; gap: 6px; }
.rec-tag { font-size: 0.72em; font-weight: 600; padding: 3px 10px; border-radius: 20px; border: 1px solid #d4d4d4; color: #555; background: #fafafa; }
.effort-low  { border-color: #a9d18e; color: #375623; background: #f2f7ed; }
.effort-med  { border-color: #f4b942; color: #7d5000; background: #fef8ec; }
.effort-high { border-color: #e07b7b; color: #7b0000; background: #fdf0f0; }
.impact-high { border-color: #4fa7db; color: #1a5276; background: #ebf5fb; }

/* Data coverage */
.coverage-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 10px; }
.coverage-item { display: flex; align-items: center; gap: 12px; padding: 11px 14px; background: white; border: 1px solid #e2e8f0; border-radius: 6px; }
.coverage-icon { font-size: 1.2em; width: 26px; text-align: center; }
.coverage-label { font-size: 0.84em; font-weight: 600; color: #2c3e50; }
.coverage-status { font-size: 0.76em; }

/* Tables */
.data-table { width: 100%; border-collapse: collapse; font-size: 0.83em; margin-top: 12px; }
.data-table th { background: #5a2f61; color: white; text-align: left; padding: 10px 14px; font-weight: 600; font-size: 0.78em; letter-spacing: 0.3px; }
.data-table td { padding: 9px 14px; border-bottom: 1px solid #eee; color: #333; vertical-align: middle; }
.data-table tr:nth-child(even) td { background: #f9f9fc; }
.data-table tr:hover td { background: #f0f4ff; }
.appendix-section { margin-bottom: 36px; }
.appendix-section h3 { font-size: 1.02em; font-weight: 700; color: #5a2f61; padding-bottom: 8px; border-bottom: 2px solid #e2e8f0; margin-bottom: 14px; }

/* Print */
@media print {
  .cover { min-height: auto; padding: 80px; page-break-after: always; }
  .kpi-grid { grid-template-columns: repeat(3, 1fr) !important; }
  .report-section { page-break-inside: avoid; }
  @page { margin: 0.6in; }
}
@media (max-width: 600px) {
  .page { padding: 20px 14px; }
  .kpi-grid { grid-template-columns: repeat(2, 1fr); }
  .rec-card { flex-direction: column; }
  .cover-title { font-size: 2em; }
}
'@

    $allNarrativeSections = $narrativeSectionsHtml -join "`n"
    $allAppendix          = $appendixHtml -join "`n"
    $allCoverage          = $coverageHtml -join "`n"

    $html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Microsoft 365 Tenant Review &mdash; $(Escape-Html $orgName)</title>
  <style>$css</style>
</head>
<body>

<!-- ═══ COVER ══════════════════════════════════════════════════════════════ -->
<div class="cover">
  <div class="cover-inner">
    <div class="cover-eyebrow">Microsoft 365 Tenant Review</div>
    <div class="cover-divider"></div>
    <div class="cover-title">$(Escape-Html $orgName)</div>
    <div class="cover-subtitle">Review Period: $(Escape-Html $ReviewPeriod)</div>
    <div class="cover-badges">
      <span class="cover-badge">&#128197; Generated: $(Escape-Html $generatedAt)</span>
      <span class="cover-badge">&#128274; Confidential</span>
      <span class="cover-badge">&#128269; Tenant: $(Escape-Html $TenantName)</span>
    </div>
    <div class="cover-prepared">Prepared by RTH Tech Services</div>
  </div>
</div>

<div class="page">

<!-- ═══ EXECUTIVE SUMMARY ═══════════════════════════════════════════════════ -->
<div class="report-section">
  $(Build-SectionHeaderHtml -Icon '📋' -Title 'Executive Summary')
  <div class="section-body">
    <div class="exec-health">$(Escape-Html $healthStatement)</div>
    <ul class="exec-bullets">
      $execBulletsHtml
    </ul>
  </div>
</div>

<!-- ═══ KPI DASHBOARD ════════════════════════════════════════════════════════ -->
<div class="report-section">
  $(Build-SectionHeaderHtml -Icon '📊' -Title 'Key Metrics at a Glance' -Subtitle 'Headline numbers across all reviewed areas')
  <div class="section-body">
    <div class="kpi-grid">
      $kpiCardsHtml
    </div>
  </div>
</div>

<!-- ═══ MAIN SECTIONS ════════════════════════════════════════════════════════ -->
$allNarrativeSections

<!-- ═══ TOP RECOMMENDATIONS ══════════════════════════════════════════════════ -->
<div class="report-section">
  $(Build-SectionHeaderHtml -Icon '🎯' -Title 'Top Recommendations' -Subtitle 'Prioritised actions to improve the tenant')
  <div class="section-body">
    $recCardsHtml
  </div>
</div>

<!-- ═══ DATA COVERAGE ════════════════════════════════════════════════════════ -->
<div class="report-section">
  $(Build-SectionHeaderHtml -Icon '🔍' -Title 'Data Coverage' -Subtitle 'Status of each data source used in this report')
  <div class="section-body">
    <p style="font-size:0.9em;color:#555;margin-bottom:16px;">Some data sources may be unavailable depending on licensing, permissions, or configuration. Where a source was unavailable, its section is marked accordingly.</p>
    $allCoverage
  </div>
</div>

<!-- ═══ APPENDIX ═════════════════════════════════════════════════════════════ -->
$allAppendix

</div><!-- /page -->
</body>
</html>
"@

    Set-Content -Path $reportPath -Value $markdownText -Encoding UTF8
    Set-Content -Path $htmlPath -Value $html -Encoding UTF8
    Write-Host "Report created: $reportPath"
    Write-Host "HTML report created: $htmlPath"
}

