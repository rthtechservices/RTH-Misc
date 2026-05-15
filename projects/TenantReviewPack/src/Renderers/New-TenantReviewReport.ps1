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

    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $sections = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'sections')
    $attentionItems = @(Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'attentionItems')
    $reportPath = Join-Path $OutputPath 'TenantReviewReport.md'
    $htmlPath = Join-Path $OutputPath 'TenantReviewReport.html'

    $licenseSummary = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
    $userSummary = Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'summary'
    $mailboxSummary = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
    $sharePointSummary = Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'summary'
    $teamsSummary = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
    $deviceSummary = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
    $copilotSummary = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'

    $markdown = New-Object System.Collections.Generic.List[string]
    $markdown.Add("# Tenant Review Report")
    $markdown.Add("")
    $markdown.Add("Tenant: $TenantName")
    $markdown.Add("")
    $markdown.Add("Review period: $ReviewPeriod")
    $markdown.Add("")
    $markdown.Add("Generated: $((Get-Date).ToString('o'))")
    $markdown.Add("")
    $markdown.Add("## Executive Summary")
    foreach ($section in $sections) {
        $markdown.Add("")
        $markdown.Add("### $(Get-TenantReviewProperty -InputObject $section -Name 'headline')")
        $markdown.Add("")
        $markdown.Add("$(Get-TenantReviewProperty -InputObject $section -Name 'plainEnglish')")
        $markdown.Add("")
        $markdown.Add("Business impact: $(Get-TenantReviewProperty -InputObject $section -Name 'businessImpact')")
        $markdown.Add("")
        $markdown.Add("Recommended action: $(Get-TenantReviewProperty -InputObject $section -Name 'recommendedAction')")
    }

    $markdown.Add("")
    $markdown.Add("## Big Numbers")
    $markdown.Add("")
    $markdown.Add("| Area | Metric | Value |")
    $markdown.Add("| --- | --- | ---: |")
    $markdown.Add("| Licenses | Purchased | $(Get-TenantReviewProperty -InputObject $licenseSummary -Name 'totalPurchased') |")
    $markdown.Add("| Licenses | Unused | $(Get-TenantReviewProperty -InputObject $licenseSummary -Name 'totalUnused') |")
    $markdown.Add("| Users | Total users | $(Get-TenantReviewProperty -InputObject $userSummary -Name 'totalUsers') |")
    $markdown.Add("| Users | Licensed disabled users | $(Get-TenantReviewProperty -InputObject $userSummary -Name 'licensedDisabledUsers') |")
    $markdown.Add("| Mail | External forwarding suspected | $(Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'mailboxesForwardingExternally') |")
    $markdown.Add("| SharePoint | Total sites | $(Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'totalSites') |")
    $markdown.Add("| Teams | Inactive teams | $(Get-TenantReviewProperty -InputObject $teamsSummary -Name 'inactiveTeams') |")
    $markdown.Add("| Devices | Stale devices | $(Get-TenantReviewProperty -InputObject $deviceSummary -Name 'staleDevices') |")
    $markdown.Add("| Copilot | Unused Copilot licenses | $(Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotUnused') |")

    $markdown.Add("")
    $markdown.Add("## Attention Items")
    if ($attentionItems.Count -eq 0) {
        $markdown.Add("")
        $markdown.Add("No license/user attention items were generated from the available data.")
    } else {
        foreach ($item in $attentionItems) {
            $markdown.Add("")
            $markdown.Add("- **$(Get-TenantReviewProperty -InputObject $item -Name 'severity')**: $(Get-TenantReviewProperty -InputObject $item -Name 'headline') - $(Get-TenantReviewProperty -InputObject $item -Name 'recommendedAction')")
        }
    }

    $markdown.Add("")
    $markdown.Add("## Dataset Notes")
    foreach ($key in $Datasets.Keys) {
        $dataset = $Datasets[$key]
        $warnings = @(Get-TenantReviewProperty -InputObject $dataset -Name 'warnings')
        $markdown.Add("")
        $markdown.Add("### $key")
        $markdown.Add("")
        if ($warnings.Count -gt 0) {
            foreach ($warning in $warnings) {
                $markdown.Add("- Warning: $warning")
            }
        } else {
            $markdown.Add("- No collector warnings.")
        }
        $markdown.Add(("- JSON: ``{0}.json``" -f $key))
    }

    $newline = [Environment]::NewLine
    $markdownText = $markdown -join $newline
    Set-Content -Path $reportPath -Value $markdownText -Encoding UTF8

    $encodedMarkdown = [System.Net.WebUtility]::HtmlEncode($markdownText)
    $htmlBody = $encodedMarkdown
    $htmlLines = @(
        '<!doctype html>',
        '<html>',
        '<head>',
        '  <meta charset="utf-8" />',
        '  <title>Tenant Review Report</title>',
        '  <style>',
        '    body { font-family: Segoe UI, Arial, sans-serif; color: #16202a; margin: 32px; background: #f7f9fb; }',
        '    main { max-width: 1100px; margin: auto; background: white; padding: 32px; border: 1px solid #d9e2ec; }',
        '    h1, h2, h3 { color: #0f4c81; }',
        '    table { border-collapse: collapse; width: 100%; }',
        '    th, td { border: 1px solid #d9e2ec; padding: 8px; }',
        '    code { background: #eef3f8; padding: 2px 4px; }',
        '    pre { white-space: pre-wrap; font-family: Segoe UI, Arial, sans-serif; line-height: 1.45; }',
        '  </style>',
        '</head>',
        "<body><main><pre>$htmlBody</pre></main></body>",
        '</html>'
    )
    $html = $htmlLines -join $newline
    Set-Content -Path $htmlPath -Value $html -Encoding UTF8

    Write-Host "Created report: $reportPath"
    Write-Host "Created HTML report: $htmlPath"
}
