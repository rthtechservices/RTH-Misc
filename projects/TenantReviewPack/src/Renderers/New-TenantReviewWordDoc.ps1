function New-TenantReviewWordDoc {
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

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $docPath = Join-Path $OutputPath 'TenantReviewDetailedReport.docx'
    if (Test-Path $docPath) {
        Remove-Item -Path $docPath -Force
    }

    $generatedAt = Get-Date
    $generatedLabel = $generatedAt.ToString('MMMM d, yyyy')
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)

    function Escape-WordXml {
        param([object]$Value)

        if ($null -eq $Value) {
            return ''
        }

        return [System.Security.SecurityElement]::Escape($Value.ToString())
    }

    function Nvl {
        param([object]$Value, [string]$Default = 'N/A')

        if ($null -eq $Value) {
            return $Default
        }

        $text = $Value.ToString()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $Default
        }

        return $text
    }

    function Join-WordValue {
        param([object]$Value, [string]$Separator = ', ')

        if ($null -eq $Value) {
            return 'N/A'
        }

        if ($Value -is [string]) {
            if ([string]::IsNullOrWhiteSpace($Value)) { return 'N/A' }
            return $Value
        }

        $items = @($Value | Where-Object { $null -ne $_ -and $_.ToString() -ne '' })
        if ($items.Count -eq 0) {
            return 'N/A'
        }

        return (($items | ForEach-Object { $_.ToString() }) -join $Separator)
    }

    function Format-WordBool {
        param([object]$Value)

        if ($null -eq $Value) { return 'N/A' }
        if ([bool]$Value) {
            return 'Yes'
        }

        return 'No'
    }

    function Format-WordDate {
        param([object]$Value)

        $date = ConvertTo-TenantReviewDateTime -Value $Value
        if ($null -eq $date) {
            return 'N/A'
        }

        return $date.ToString('yyyy-MM-dd')
    }

    function Format-WordDateTime {
        param([object]$Value)

        $date = ConvertTo-TenantReviewDateTime -Value $Value
        if ($null -eq $date) {
            return 'N/A'
        }

        return $date.ToString('yyyy-MM-dd HH:mm')
    }

    function Format-WordNumber {
        param([object]$Value)

        if ($null -eq $Value -or $Value.ToString() -eq '') {
            return '0'
        }

        return ('{0:N0}' -f ([decimal](ConvertTo-TenantReviewDecimal -Value $Value -Default 0)))
    }

    function Format-WordDecimal {
        param([object]$Value, [int]$Decimals = 2)

        if ($null -eq $Value -or $Value.ToString() -eq '') {
            return '0'
        }

        $format = '{0:N' + $Decimals + '}'
        return ($format -f ([decimal](ConvertTo-TenantReviewDecimal -Value $Value -Default 0)))
    }

    function Format-WordCurrency {
        param([object]$Value, [string]$Currency = 'CAD')

        if ($null -eq $Value -or $Value.ToString() -eq '') {
            return "$Currency $0.00"
        }

        $amount = [decimal](ConvertTo-TenantReviewDecimal -Value $Value -Default 0)
        return "$Currency `$$(('{0:N2}' -f $amount))"
    }

    function Get-StatusColor {
        param([string]$Status)

        switch ($Status) {
            'Good'              { return '27AE60' }
            'Watch'             { return 'F39C12' }
            'ActionRecommended' { return 'E74C3C' }
            'NotAvailable'      { return '767779' }
            default             { return '4FA7DB' }
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

    function New-WordParagraph {
        param(
            [string]$Text,
            [string]$Style = 'Normal',
            [string]$Color = $null,
            [switch]$Bold,
            [ValidateSet('left', 'center', 'right')]
            [string]$Alignment = 'left'
        )

        $safeText = if ($null -eq $Text) { '' } else { $Text }
        $spaceAttr = if ($safeText -match '^\s|\s$|  ') { ' xml:space="preserve"' } else { '' }
        $alignXml = if ($Alignment -ne 'left') { "<w:jc w:val=`"$Alignment`"/>" } else { '' }
        $colorXml = if ($Color) { "<w:color w:val=`"$Color`"/>" } else { '' }
        $boldXml = if ($Bold) { '<w:b/><w:bCs/>' } else { '' }

        return "<w:p><w:pPr><w:pStyle w:val=`"$Style`"/>$alignXml</w:pPr><w:r><w:rPr>$boldXml$colorXml</w:rPr><w:t$spaceAttr>$(Escape-WordXml $safeText)</w:t></w:r></w:p>"
    }

    function New-WordPageBreak {
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    }

    function New-WordTableCell {
        param(
            [string]$Text,
            [switch]$Header,
            [string]$Fill = 'FFFFFF'
        )

        $cellText = if ($null -eq $Text -or $Text -eq '') { ' ' } else { $Text }
        $color = if ($Header) { 'FFFFFF' } else { '1F1F1F' }
        $boldXml = if ($Header) { '<w:b/><w:bCs/>' } else { '' }
        $fillXml = if ($Header) { "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$Fill`"/>" } else { '' }
        $spaceAttr = if ($cellText -match '^\s|\s$|  ') { ' xml:space="preserve"' } else { '' }

        return "<w:tc><w:tcPr>$fillXml<w:tcW w:w=`"0`" w:type=`"auto`"/></w:tcPr><w:p><w:pPr><w:spacing w:after=`"80`"/></w:pPr><w:r><w:rPr>$boldXml<w:color w:val=`"$color`"/></w:rPr><w:t$spaceAttr>$(Escape-WordXml $cellText)</w:t></w:r></w:p></w:tc>"
    }

    function New-WordTable {
        param(
            [string[]]$Headers,
            [object[]]$Rows
        )

        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add('<w:tbl>')
        $parts.Add('<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="D7D9E0"/><w:left w:val="single" w:sz="8" w:space="0" w:color="D7D9E0"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="D7D9E0"/><w:right w:val="single" w:sz="8" w:space="0" w:color="D7D9E0"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="E6E8EF"/><w:insideV w:val="single" w:sz="6" w:space="0" w:color="E6E8EF"/></w:tblBorders></w:tblPr>')

        if ($Headers -and $Headers.Count -gt 0) {
            $parts.Add('<w:tr>')
            foreach ($header in $Headers) {
                $parts.Add((New-WordTableCell -Text $header -Header -Fill '5A2F61'))
            }
            $parts.Add('</w:tr>')
        }

        foreach ($row in @($Rows)) {
            $parts.Add('<w:tr>')
            foreach ($cell in @($row)) {
                $parts.Add((New-WordTableCell -Text (Nvl $cell ' ')))
            }
            $parts.Add('</w:tr>')
        }

        $parts.Add('</w:tbl>')
        return ($parts -join '')
    }

    function Add-BodyParagraph {
        param([System.Collections.Generic.List[string]]$Body, [string]$Text, [string]$Style = 'Normal', [string]$Color = $null, [switch]$Bold, [string]$Alignment = 'left')
        $Body.Add((New-WordParagraph -Text $Text -Style $Style -Color $Color -Bold:$Bold -Alignment $Alignment))
    }

    function Add-BodyTable {
        param([System.Collections.Generic.List[string]]$Body, [string[]]$Headers, [object[]]$Rows)
        $Body.Add((New-WordTable -Headers $Headers -Rows $Rows))
        $Body.Add('<w:p/>')
    }

    $sections = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'sections')
    $recommendations = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'recommendations')
    $execBullets = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'executiveSummaryBullets')
    $healthStatement = Nvl (Get-TenantReviewProperty -InputObject $Narrative -Name 'tenantHealthStatement') 'Review the findings and detailed sections in this report.'

    function Get-Section {
        param([string]$Dataset)

        return ($sections | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'dataset') -eq $Dataset } | Select-Object -First 1)
    }

    $sectionTitles = @{
        'TenantOverview'      = 'Tenant Overview'
        'LicenseInventory'    = 'Licensing and Cost'
        'LicenseUserAnalysis' = 'License Analysis'
        'UserInventory'       = 'Users and Identity'
        'MailboxInventory'    = 'Exchange and Mail Flow'
        'SharePoint'          = 'SharePoint and OneDrive'
        'Teams'               = 'Teams Collaboration'
        'Devices'             = 'Devices and Endpoints'
        'Copilot'             = 'Copilot'
    }

    $body = [System.Collections.Generic.List[string]]::new()

    $tenantSummary = Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'summary'
    $userSummary = Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'summary'
    $licenseSummary = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
    $licenseAnalysisSummary = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
    $mailboxSummary = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
    $sharePointSummary = Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'summary'
    $teamsSummary = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
    $deviceSummary = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
    $copilotSummary = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'
    $orgName = Nvl (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'organizationName') $TenantName

    Add-BodyParagraph -Body $body -Text 'Microsoft 365 Tenant Review' -Style 'Title' -Color '5A2F61' -Bold -Alignment 'center'
    Add-BodyParagraph -Body $body -Text $orgName -Style 'Subtitle' -Color '4FA7DB' -Alignment 'center'
    Add-BodyParagraph -Body $body -Text "Review period: $ReviewPeriod" -Style 'Subtitle' -Alignment 'center'
    Add-BodyParagraph -Body $body -Text 'Prepared by: RTH Tech Services' -Style 'Subtitle' -Alignment 'center'
    Add-BodyParagraph -Body $body -Text "Generated: $generatedLabel" -Style 'Subtitle' -Alignment 'center'
    $body.Add((New-WordPageBreak))

    Add-BodyParagraph -Body $body -Text 'Executive Summary' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyParagraph -Body $body -Text $healthStatement -Style 'Normal'
    foreach ($bullet in $execBullets) {
        Add-BodyParagraph -Body $body -Text ("- " + (Nvl $bullet)) -Style 'Normal'
    }
    foreach ($key in @('TenantOverview','LicenseInventory','LicenseUserAnalysis','UserInventory','MailboxInventory','SharePoint','Teams','Devices','Copilot')) {
        $section = Get-Section -Dataset $key
        if (-not $section) { continue }
        $title = $sectionTitles[$key]
        $headline = Nvl (Get-TenantReviewProperty -InputObject $section -Name 'headline') $title
        $detail = Get-TenantReviewProperty -InputObject $section -Name 'detailedAnalysis'
        if (-not $detail) {
            $detail = Get-TenantReviewProperty -InputObject $section -Name 'plainEnglish'
        }
        $impact = Nvl (Get-TenantReviewProperty -InputObject $section -Name 'businessImpact')
        $action = Nvl (Get-TenantReviewProperty -InputObject $section -Name 'recommendedAction')
        $status = Nvl (Get-TenantReviewProperty -InputObject $section -Name 'status') 'Informational'
        Add-BodyParagraph -Body $body -Text $title -Style 'Heading2' -Color '4FA7DB' -Bold
        Add-BodyParagraph -Body $body -Text ("Status: " + (Get-StatusLabel $status)) -Style 'Normal' -Color (Get-StatusColor $status) -Bold
        Add-BodyParagraph -Body $body -Text $headline -Style 'Heading3' -Color '5A2F61' -Bold
        Add-BodyParagraph -Body $body -Text (Nvl $detail) -Style 'Normal'
        Add-BodyParagraph -Body $body -Text ("Why it matters: " + $impact) -Style 'Normal'
        Add-BodyParagraph -Body $body -Text ("Recommended action: " + $action) -Style 'Normal'
    }

    Add-BodyParagraph -Body $body -Text 'Priority Recommendations' -Style 'Heading1' -Color '5A2F61' -Bold
    if ($recommendations.Count -gt 0) {
        $recRows = @(
            foreach ($rec in $recommendations) {
                ,@(
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'priority')),
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'category')),
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'title')),
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'why')),
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'owner')),
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'effort')),
                    $(Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'impact')),
                    $(Get-StatusLabel (Nvl (Get-TenantReviewProperty -InputObject $rec -Name 'status') 'Watch'))
                )
            }
        )
        Add-BodyTable -Body $body -Headers @('Priority','Category','Recommendation','Reason','Owner','Effort','Impact','Status') -Rows $recRows
    } else {
        Add-BodyParagraph -Body $body -Text 'No recommendations were generated from the available data.' -Style 'Normal'
    }

    Add-BodyParagraph -Body $body -Text 'Tenant Overview' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Organisation', $orgName),
        ,@('Default domain', $(Nvl (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'defaultDomain'))),
        ,@('Initial domain', $(Nvl (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'initialDomain'))),
        ,@('Created', $(Format-WordDate (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'createdDateTime'))),
        ,@('Domain count', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'domainCount'))),
        ,@('Verified domains', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'verifiedDomainCount'))),
        ,@('Federated domains', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'federatedDomainCount'))),
        ,@('Assigned plans', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'assignedPlansCount'))),
        ,@('Technical notifications', $(Join-WordValue (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'technicalNotificationMails')))
    )
    $domainRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'id')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isDefault')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isInitial')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isVerified')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'authenticationType')),
                $(Join-WordValue (Get-TenantReviewProperty -InputObject $item -Name 'supportedServices'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Domain Inventory' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Domain','Default','Initial','Verified','Authentication','Supported Services') -Rows $domainRows

    Add-BodyParagraph -Body $body -Text 'Users and Identity' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Total users', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'totalUsers'))),
        ,@('Members', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'memberUsers'))),
        ,@('Guests', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'guestUsers'))),
        ,@('Enabled users', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'enabledUsers'))),
        ,@('Disabled users', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'disabledUsers'))),
        ,@('Licensed users', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'licensedUsers'))),
        ,@('Stale users', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $userSummary -Name 'staleUsers'))),
        ,@('Sign-in data status', $(Nvl (Get-TenantReviewProperty -InputObject $userSummary -Name 'signInActivityStatus')))
    )
    $userRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userPrincipalName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'mail')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userType')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'accountEnabled')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isLicensed')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'licenseCount')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'department')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'jobTitle')),
                $(Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'createdDateTime')),
                $(Format-WordDateTime (Get-TenantReviewProperty -InputObject $item -Name 'lastSuccessfulSignInDateTime')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'daysSinceLastSuccessfulSignIn')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isStale'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Active Directory Accounts' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Display Name','UPN','Mail','Type','Enabled','Licensed','License Count','Department','Job Title','Created','Last Successful Sign-In','Days Since Sign-In','Stale') -Rows $userRows

    Add-BodyParagraph -Body $body -Text 'Licensing and Cost' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Purchased units', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'totalPurchased'))),
        ,@('Assigned units', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'totalAssigned'))),
        ,@('Unused units', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'totalUnused'))),
        ,@('Estimated monthly cost', $(Format-WordCurrency (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'estimatedMonthlyCost') (Nvl (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'currency') 'CAD'))),
        ,@('Unused monthly cost', $(Format-WordCurrency (Get-TenantReviewProperty -InputObject $licenseAnalysisSummary -Name 'estimatedUnusedMonthlyCost') 'CAD')),
        ,@('Missing price map SKUs', $(Join-WordValue (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'skusMissingPrice')))
    )
    $licenseRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'purchasedUnits')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'assignedUnits')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'unusedUnits')),
                $(Format-WordCurrency (Get-TenantReviewProperty -InputObject $item -Name 'estimatedMonthlyCost') (Nvl (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'currency') 'CAD'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'SKU Inventory' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('SKU','Purchased','Assigned','Unused','Monthly Cost') -Rows $licenseRows
    $attentionRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'attentionItems')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'severity')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'headline')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'plainEnglish')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'businessImpact')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'recommendedAction'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'License Attention Items' -Style 'Heading2' -Color '4FA7DB' -Bold
    if ($attentionRows.Count -gt 0) {
        Add-BodyTable -Body $body -Headers @('Severity','Headline','Plain English','Business Impact','Recommended Action') -Rows $attentionRows
    } else {
        Add-BodyParagraph -Body $body -Text 'No license attention items were generated.' -Style 'Normal'
    }

    Add-BodyParagraph -Body $body -Text 'Exchange and Mail Flow' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Total mailboxes', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'totalMailboxes'))),
        ,@('User mailboxes', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'userMailboxes'))),
        ,@('Shared mailboxes', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'sharedMailboxes'))),
        ,@('Mailboxes with forwarding', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'mailboxesWithForwarding'))),
        ,@('External forwarding suspected', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'mailboxesForwardingExternally'))),
        ,@('Enabled transport rules', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'enabledTransportRules'))),
        ,@('Inbox rules found', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'inboxForwardingRulesFound'))),
        ,@('Mailbox statistics included', $(Format-WordBool (Get-TenantReviewProperty -InputObject $mailboxSummary -Name 'mailboxStatisticsIncluded')))
    )
    $mailboxRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'primarySmtpAddress')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'recipientTypeDetails')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isShared')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'forwardingEnabled')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'forwardingAddress')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'forwardingSmtpAddress')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'deliverToMailboxAndForward')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'externalForwardingSuspected')),
                $(Format-WordDecimal (Get-TenantReviewProperty -InputObject $item -Name 'totalItemSizeGB') 2),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'itemCount'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Mailbox Inventory' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Display Name','Primary SMTP','Mailbox Type','Shared','Forwarding Enabled','Forwarding Address','Forwarding SMTP','Keep Copy','External Forwarding','Size (GB)','Items') -Rows $mailboxRows
    $transportRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'transportRules')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'name')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'state')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'mode')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'priority')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'comments')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'conditionsSummary')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'actionsSummary'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Transport Rules' -Style 'Heading2' -Color '4FA7DB' -Bold
    if ($transportRows.Count -gt 0) {
        Add-BodyTable -Body $body -Headers @('Name','State','Mode','Priority','Comments','Conditions','Actions') -Rows $transportRows
    } else {
        Add-BodyParagraph -Body $body -Text 'No transport rules were collected.' -Style 'Normal'
    }
    $inboxRuleRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'inboxForwardingRules')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'mailbox')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'name')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'enabled')),
                $(Join-WordValue (Get-TenantReviewProperty -InputObject $item -Name 'forwardTo')),
                $(Join-WordValue (Get-TenantReviewProperty -InputObject $item -Name 'redirectTo')),
                $(Join-WordValue (Get-TenantReviewProperty -InputObject $item -Name 'forwardAsAttachmentTo'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Inbox Rules with Redirect or Forward Actions' -Style 'Heading2' -Color '4FA7DB' -Bold
    if ($inboxRuleRows.Count -gt 0) {
        Add-BodyTable -Body $body -Headers @('Mailbox','Rule Name','Enabled','Forward To','Redirect To','Forward as Attachment To') -Rows $inboxRuleRows
    } else {
        Add-BodyParagraph -Body $body -Text 'No inbox forwarding rules were collected.' -Style 'Normal'
    }

    Add-BodyParagraph -Body $body -Text 'SharePoint and OneDrive' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Total sites', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'totalSites'))),
        ,@('Total storage (GB)', $(Format-WordDecimal (Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'totalStorageGB') 2)),
        ,@('External sharing enabled sites', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'externalSharingEnabledSites'))),
        ,@('OneDrive accounts', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'oneDriveAccounts'))),
        ,@('OneDrive storage (GB)', $(Format-WordDecimal (Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'oneDriveTotalStorageGB') 2)),
        ,@('Largest site', $(Nvl (Get-TenantReviewProperty -InputObject $sharePointSummary -Name 'largestSiteTitle')))
    )
    $sharePointRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'title')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'url')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'template')),
                $(Format-WordDecimal (Get-TenantReviewProperty -InputObject $item -Name 'storageUsageGB') 2),
                $(Format-WordDecimal (Get-TenantReviewProperty -InputObject $item -Name 'storageQuotaGB') 2),
                $(Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'lastContentModifiedDate')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'sharingCapability')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'externalSharingEnabled')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'activeFileCount')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'fileCount'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Site Inventory' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Title','URL','Template','Storage Used (GB)','Quota (GB)','Last Modified','Sharing Capability','External Sharing','Active Files','Files') -Rows $sharePointRows

    Add-BodyParagraph -Body $body -Text 'Teams Collaboration' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Total Teams', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $teamsSummary -Name 'totalTeams'))),
        ,@('Inactive Teams', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $teamsSummary -Name 'inactiveTeams'))),
        ,@('Teams with no owners', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $teamsSummary -Name 'teamsWithNoOwners')))
    )
    $teamRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'visibility')),
                $(Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'createdDateTime')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isArchived')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'ownerCount')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'memberCount')),
                $(Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'lastActivityDate')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'channelMessages')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'replyMessages')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'meetingsOrganized')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'inactive'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Team Inventory' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Team','Visibility','Created','Archived','Owners','Members','Last Activity','Channel Messages','Replies','Meetings','Inactive') -Rows $teamRows

    Add-BodyParagraph -Body $body -Text 'Devices and Endpoints' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Total devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'totalDevices'))),
        ,@('Enabled devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'enabledDevices'))),
        ,@('Disabled devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'disabledDevices'))),
        ,@('Stale devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'staleDevices'))),
        ,@('Windows devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'windowsDevices'))),
        ,@('Android devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'androidDevices'))),
        ,@('Intune managed devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'intuneManagedDevices'))),
        ,@('Unknown compliance devices', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $deviceSummary -Name 'unknownComplianceDevices')))
    )
    $deviceRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'deviceId')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'accountEnabled')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'operatingSystem')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'operatingSystemVersion')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'trustType')),
                $(Format-WordDateTime (Get-TenantReviewProperty -InputObject $item -Name 'approximateLastSignInDateTime')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'daysSinceLastSignIn')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isStale')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isManaged')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'complianceState')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userPrincipalName'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Device Inventory' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Device','Device ID','Enabled','OS','OS Version','Trust Type','Last Sign-In','Days Since Sign-In','Stale','Managed','Compliance','User') -Rows $deviceRows

    Add-BodyParagraph -Body $body -Text 'Copilot' -Style 'Heading1' -Color '5A2F61' -Bold
    Add-BodyTable -Body $body -Headers @('Property','Value') -Rows @(
        ,@('Purchased licenses', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotPurchased'))),
        ,@('Assigned licenses', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotAssigned'))),
        ,@('Unused licenses', $(Format-WordNumber (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'copilotUnused'))),
        ,@('Active Copilot users', $(Nvl (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'activeCopilotUsers'))),
        ,@('Usage report status', $(Nvl (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'usageReportStatus'))),
        ,@('Estimated monthly cost', $(Format-WordCurrency (Get-TenantReviewProperty -InputObject $copilotSummary -Name 'estimatedMonthlyCost') 'CAD'))
    )
    $copilotRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'items')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'skuPartNumber')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'purchasedUnits')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'assignedUnits')),
                $(Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'unusedUnits')),
                $(Format-WordCurrency (Get-TenantReviewProperty -InputObject $item -Name 'estimatedMonthlyCost') 'CAD')
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Copilot SKU Detail' -Style 'Heading2' -Color '4FA7DB' -Bold
    Add-BodyTable -Body $body -Headers @('Display Name','SKU Part Number','Purchased','Assigned','Unused','Monthly Cost') -Rows $copilotRows
    $copilotUserRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'licensedUsers')) {
            ,@(
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                $(Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userPrincipalName')),
                $(Join-WordValue (Get-TenantReviewProperty -InputObject $item -Name 'assignedCopilotSkuPartNumbers')),
                $(Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isStale')),
                $(Format-WordDateTime (Get-TenantReviewProperty -InputObject $item -Name 'lastSuccessfulSignInDateTime'))
            )
        }
    )
    Add-BodyParagraph -Body $body -Text 'Licensed Copilot Users' -Style 'Heading2' -Color '4FA7DB' -Bold
    if ($copilotUserRows.Count -gt 0) {
        Add-BodyTable -Body $body -Headers @('Display Name','UPN','Assigned Copilot SKUs','Stale','Last Successful Sign-In') -Rows $copilotUserRows
    } else {
        Add-BodyParagraph -Body $body -Text 'No Copilot-licensed users were identified.' -Style 'Normal'
    }

    Add-BodyParagraph -Body $body -Text 'Data Coverage' -Style 'Heading1' -Color '5A2F61' -Bold
    $coverageRows = @(
        foreach ($key in $Datasets.Keys) {
            $dataset = $Datasets[$key]
            $summary = Get-TenantReviewProperty -InputObject $dataset -Name 'summary'
            $warnings = @(Get-TenantReviewProperty -InputObject $dataset -Name 'warnings')
            $skipped = Get-TenantReviewProperty -InputObject $summary -Name 'skipped'
            $status = if ($skipped) { 'Skipped' } elseif ($warnings.Count -gt 0) { "Collected with $($warnings.Count) warning(s)" } else { 'Collected' }
            ,@(
                $key,
                $status,
                $(Join-WordValue $warnings '; ')
            )
        }
    )
    Add-BodyTable -Body $body -Headers @('Dataset','Status','Warnings') -Rows $coverageRows

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    $($body -join "`n    ")
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="900" w:bottom="1080" w:left="900" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="1F1F1F"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="0" w:after="240"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:color w:val="5A2F61"/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="120"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:color w:val="767779"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="260" w:after="140"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:color w:val="5A2F61"/>
      <w:sz w:val="30"/>
      <w:szCs w:val="30"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="180" w:after="100"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:color w:val="4FA7DB"/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="120" w:after="60"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:color w:val="5A2F61"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>
"@

    $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"@

    $rootRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

    $documentRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

    function Add-ZipEntry {
        param(
            [Parameter(Mandatory = $true)][System.IO.Compression.ZipArchive]$Archive,
            [Parameter(Mandatory = $true)][string]$EntryName,
            [Parameter(Mandatory = $true)][string]$Content
        )

        $entry = $Archive.CreateEntry($EntryName)
        $stream = $entry.Open()
        try {
            $bytes = $utf8NoBom.GetBytes($Content)
            $stream.Write($bytes, 0, $bytes.Length)
        } finally {
            $stream.Dispose()
        }
    }

    $archive = [System.IO.Compression.ZipFile]::Open($docPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipEntry -Archive $archive -EntryName '[Content_Types].xml' -Content $contentTypesXml
        Add-ZipEntry -Archive $archive -EntryName '_rels/.rels' -Content $rootRelsXml
        Add-ZipEntry -Archive $archive -EntryName 'word/document.xml' -Content $documentXml
        Add-ZipEntry -Archive $archive -EntryName 'word/styles.xml' -Content $stylesXml
        Add-ZipEntry -Archive $archive -EntryName 'word/_rels/document.xml.rels' -Content $documentRelsXml
    } finally {
        $archive.Dispose()
    }

    return $docPath
}