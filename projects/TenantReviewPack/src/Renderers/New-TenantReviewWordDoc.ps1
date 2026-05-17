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

    $colors = @{
        DeepPurple  = '5A2F61'
        AccentPurple = '6F28DB'
        Blue        = '4FA7DB'
        Grey        = '767779'
        Text        = '1F2933'
        Muted       = '5F6570'
        Border      = 'D7D9E0'
        SoftGrey    = 'F7F8FB'
        LightPurple = 'F5EEFF'
        LightBlue   = 'EBF5FB'
        White       = 'FFFFFF'
        Green       = '27AE60'
        Amber       = 'F39C12'
        Red         = 'E74C3C'
    }

    $script:UsableWidth = 10440
    $portraitSection = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="540" w:footer="540" w:gutter="0"/></w:sectPr>'
    $landscapeSection = '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="540" w:footer="540" w:gutter="0"/></w:sectPr>'

    function Escape-WordXml {
        param([object]$Value)
        if ($null -eq $Value) { return '' }
        return [System.Security.SecurityElement]::Escape($Value.ToString())
    }

    function Test-TenantReviewSimpleValue {
        param([object]$Value)

        if ($null -eq $Value) { return $true }
        return (
            $Value -is [string] -or
            $Value -is [char] -or
            $Value -is [bool] -or
            $Value -is [datetime] -or
            $Value -is [byte] -or
            $Value -is [int16] -or
            $Value -is [int] -or
            $Value -is [int64] -or
            $Value -is [single] -or
            $Value -is [double] -or
            $Value -is [decimal]
        )
    }

    function ConvertTo-TenantReviewDisplayText {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [object]$Value,

            [string]$Default = 'N/A'
        )

        if ($null -eq $Value) {
            return $Default
        }

        if ($Value -is [string]) {
            if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
            $trimmed = $Value.Trim()
            if ($trimmed -match '^Microsoft\.Graph\.PowerShell\.Models\.') { return 'Structured Graph value omitted' }
            return $trimmed
        }

        if ($Value -is [datetime]) {
            return ([datetime]$Value).ToString('yyyy-MM-dd')
        }

        if ($Value -is [bool]) {
            if ([bool]$Value) { return 'Yes' }
            return 'No'
        }

        if (Test-TenantReviewSimpleValue -Value $Value) {
            return $Value.ToString()
        }

        if ($Value -is [System.Collections.IDictionary]) {
            $pairs = @()
            foreach ($key in $Value.Keys) {
                if ($null -eq $key) { continue }
                $keyText = $key.ToString()
                if ($keyText -match '^@odata|AdditionalProperties|Raw|Json') { continue }
                $entry = $Value[$key]
                if ($null -eq $entry) { continue }
                if (Test-TenantReviewSimpleValue -Value $entry) {
                    $entryText = ConvertTo-TenantReviewDisplayText -Value $entry -Default ''
                    if (-not [string]::IsNullOrWhiteSpace($entryText)) {
                        $pairs += "${keyText}: $entryText"
                    }
                }
            }
            if ($pairs.Count -eq 0) { return $Default }
            return (($pairs | Select-Object -First 4) -join '; ')
        }

        if ($Value -is [System.Collections.IEnumerable]) {
            $items = @()
            foreach ($item in $Value) {
                if ($null -ne $item) { $items += $item }
            }
            if ($items.Count -eq 0) { return 'None' }

            $simpleItems = @($items | Where-Object { Test-TenantReviewSimpleValue -Value $_ })
            if ($simpleItems.Count -eq $items.Count) {
                return (($items | ForEach-Object { ConvertTo-TenantReviewDisplayText -Value $_ -Default '' } | Where-Object { $_ }) -join ', ')
            }

            $summaries = @()
            foreach ($item in $items) {
                foreach ($propertyName in @('displayName','name','title','userPrincipalName','mail','primarySmtpAddress','skuPartNumber','id','value','warning','Warning')) {
                    $propertyValue = Get-TenantReviewProperty -InputObject $item -Name $propertyName
                    if ($null -ne $propertyValue -and (Test-TenantReviewSimpleValue -Value $propertyValue)) {
                        $display = ConvertTo-TenantReviewDisplayText -Value $propertyValue -Default ''
                        if ($display) {
                            $summaries += $display
                            break
                        }
                    }
                }
            }

            if ($summaries.Count -gt 0) {
                return (($summaries | Select-Object -First 6) -join ', ')
            }

            return "$($items.Count) structured item$(if ($items.Count -ne 1) { 's' })"
        }

        $knownPairs = @()
        foreach ($propertyName in @('displayName','name','title','userPrincipalName','mail','primarySmtpAddress','skuPartNumber','status','state','warning','Warning')) {
            $propertyValue = Get-TenantReviewProperty -InputObject $Value -Name $propertyName
            if ($null -ne $propertyValue -and (Test-TenantReviewSimpleValue -Value $propertyValue)) {
                $knownPairs += (ConvertTo-TenantReviewDisplayText -Value $propertyValue -Default '')
            }
        }
        if ($knownPairs.Count -gt 0) {
            return (($knownPairs | Where-Object { $_ } | Select-Object -First 3) -join ' - ')
        }

        return 'Structured value omitted'
    }

    function Nvl {
        param([object]$Value, [string]$Default = 'N/A')
        return (ConvertTo-TenantReviewDisplayText -Value $Value -Default $Default)
    }

    function Format-WordBool {
        param([object]$Value)
        if ($null -eq $Value) { return 'N/A' }
        if ([bool]$Value) { return 'Yes' }
        return 'No'
    }

    function Format-WordDate {
        param([object]$Value)
        $date = ConvertTo-TenantReviewDateTime -Value $Value
        if ($null -eq $date) { return 'N/A' }
        return $date.ToString('yyyy-MM-dd')
    }

    function Format-WordDateTime {
        param([object]$Value)
        $date = ConvertTo-TenantReviewDateTime -Value $Value
        if ($null -eq $date) { return 'N/A' }
        return $date.ToString('yyyy-MM-dd HH:mm')
    }

    function Format-WordNumber {
        param([object]$Value)
        if ($null -eq $Value -or $Value.ToString() -eq '') { return '0' }
        return ('{0:N0}' -f ([decimal](ConvertTo-TenantReviewDecimal -Value $Value -Default 0)))
    }

    function Format-WordDecimal {
        param([object]$Value, [int]$Decimals = 2)
        if ($null -eq $Value -or $Value.ToString() -eq '') { return '0' }
        $format = '{0:N' + $Decimals + '}'
        return ($format -f ([decimal](ConvertTo-TenantReviewDecimal -Value $Value -Default 0)))
    }

    function Format-WordCurrency {
        param([object]$Value, [string]$Currency = 'CAD')
        $amount = [decimal](ConvertTo-TenantReviewDecimal -Value $Value -Default 0)
        return ('{0} ${1:N2}' -f $Currency, $amount)
    }

    function Get-StatusColor {
        param([string]$Status)
        switch ($Status) {
            'Good'              { return $colors.Green }
            'Healthy'           { return $colors.Green }
            'Watch'             { return $colors.Amber }
            'ActionRecommended' { return $colors.Red }
            'Action Recommended'{ return $colors.Red }
            'NotAvailable'      { return $colors.Grey }
            'Not Available'     { return $colors.Grey }
            default             { return $colors.Blue }
        }
    }

    function Get-StatusLightColor {
        param([string]$Status)
        switch ($Status) {
            'Good'              { return 'D5F5E3' }
            'Healthy'           { return 'D5F5E3' }
            'Watch'             { return 'FEF3D7' }
            'ActionRecommended' { return 'FDE8E8' }
            'Action Recommended'{ return 'FDE8E8' }
            'NotAvailable'      { return 'F2F3F4' }
            'Not Available'     { return 'F2F3F4' }
            default             { return $colors.LightBlue }
        }
    }

    function Get-StatusLabel {
        param([string]$Status)
        switch ($Status) {
            'Good'              { return 'Healthy' }
            'Healthy'           { return 'Healthy' }
            'Watch'             { return 'Watch' }
            'ActionRecommended' { return 'Action Recommended' }
            'Action Recommended'{ return 'Action Recommended' }
            'NotAvailable'      { return 'Not Available' }
            'Not Available'     { return 'Not Available' }
            default             { return 'Info' }
        }
    }

    function New-WordRun {
        param(
            [string]$Text,
            [string]$Color = $null,
            [int]$Size = 21,
            [switch]$Bold
        )

        $boldXml = if ($Bold) { '<w:b/><w:bCs/>' } else { '' }
        $colorXml = if ($Color) { "<w:color w:val=`"$Color`"/>" } else { '' }
        $textParts = @()
        $lines = (($Text -replace "`r`n", "`n") -split "`n")
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($i -gt 0) { $textParts += '<w:br/>' }
            $spaceAttr = if ($lines[$i] -match '^\s|\s$|  ') { ' xml:space="preserve"' } else { '' }
            $textParts += "<w:t$spaceAttr>$(Escape-WordXml $lines[$i])</w:t>"
        }

        return "<w:r><w:rPr>$boldXml$colorXml<w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/></w:rPr>$($textParts -join '')</w:r>"
    }

    function New-WordParagraph {
        param(
            [string]$Text,
            [string]$Style = 'Normal',
            [string]$Color = $null,
            [int]$Size = 21,
            [switch]$Bold,
            [ValidateSet('left', 'center', 'right')]
            [string]$Alignment = 'left',
            [int]$Before = 0,
            [int]$After = 120,
            [string]$Fill = $null,
            [switch]$KeepNext
        )

        $styleXml = if ($Style) { "<w:pStyle w:val=`"$Style`"/>" } else { '' }
        $alignXml = if ($Alignment -ne 'left') { "<w:jc w:val=`"$Alignment`"/>" } else { '' }
        $fillXml = if ($Fill) { "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$Fill`"/>" } else { '' }
        $keepXml = if ($KeepNext) { '<w:keepNext/>' } else { '' }
        $safeText = Nvl $Text ''
        return "<w:p><w:pPr>$styleXml$keepXml$alignXml<w:spacing w:before=`"$Before`" w:after=`"$After`"/>$fillXml</w:pPr>$(New-WordRun -Text $safeText -Color $Color -Size $Size -Bold:$Bold)</w:p>"
    }

    function New-WordPageBreak {
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    }

    function New-WordSectionBreakNextPage {
        param([string]$SectionXml)
        return "<w:p><w:pPr>$SectionXml</w:pPr></w:p>"
    }

    function New-CellBordersXml {
        param(
            [string]$Color = 'D7D9E0',
            [int]$Size = 6,
            [string]$TopColor = $null,
            [int]$TopSize = 6,
            [string]$LeftColor = $null,
            [int]$LeftSize = 6
        )

        $top = if ($TopColor) { $TopColor } else { $Color }
        $left = if ($LeftColor) { $LeftColor } else { $Color }
        return "<w:tcBorders><w:top w:val=`"single`" w:sz=`"$TopSize`" w:space=`"0`" w:color=`"$top`"/><w:left w:val=`"single`" w:sz=`"$LeftSize`" w:space=`"0`" w:color=`"$left`"/><w:bottom w:val=`"single`" w:sz=`"$Size`" w:space=`"0`" w:color=`"$Color`"/><w:right w:val=`"single`" w:sz=`"$Size`" w:space=`"0`" w:color=`"$Color`"/></w:tcBorders>"
    }

    function New-WordTableCell {
        param(
            [string]$Text = $null,
            [string[]]$ParagraphsXml = $null,
            [int]$Width = 2400,
            [string]$Fill = 'FFFFFF',
            [string]$Color = '1F2933',
            [int]$FontSize = 18,
            [switch]$Bold,
            [switch]$Header,
            [ValidateSet('left','center','right')]
            [string]$Alignment = 'left',
            [string]$BorderColor = 'D7D9E0',
            [string]$TopBorderColor = $null,
            [int]$TopBorderSize = 6,
            [string]$LeftBorderColor = $null,
            [int]$LeftBorderSize = 6
        )

        $paragraphs = if ($ParagraphsXml) {
            $ParagraphsXml -join ''
        } else {
            $cellText = if ($null -eq $Text -or $Text -eq '') { ' ' } else { Nvl $Text ' ' }
            $textColor = if ($Header) { $colors.White } else { $Color }
            New-WordParagraph -Text $cellText -Style 'TableText' -Color $textColor -Size $FontSize -Bold:($Bold -or $Header) -Alignment $Alignment -After 40
        }

        $borders = New-CellBordersXml -Color $BorderColor -TopColor $TopBorderColor -TopSize $TopBorderSize -LeftColor $LeftBorderColor -LeftSize $LeftBorderSize
        return "<w:tc><w:tcPr><w:tcW w:w=`"$Width`" w:type=`"dxa`"/><w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$Fill`"/><w:tcMar><w:top w:w=`"120`" w:type=`"dxa`"/><w:left w:w=`"140`" w:type=`"dxa`"/><w:bottom w:w=`"120`" w:type=`"dxa`"/><w:right w:w=`"140`" w:type=`"dxa`"/></w:tcMar><w:vAlign w:val=`"center`"/>$borders</w:tcPr>$paragraphs</w:tc>"
    }

    function New-WordTable {
        param(
            [string[]]$Headers = @(),
            [object[]]$Rows = @(),
            [int[]]$Widths = $null,
            [int]$TableWidth = $script:UsableWidth,
            [int]$FontSize = 18,
            [int]$HeaderFontSize = 18,
            [string]$HeaderFill = '5A2F61'
        )

        $columnCount = if ($Headers.Count -gt 0) { $Headers.Count } elseif ($Rows.Count -gt 0) { @($Rows[0]).Count } else { 1 }
        if (-not $Widths -or $Widths.Count -ne $columnCount) {
            $base = [int][Math]::Floor($TableWidth / $columnCount)
            $Widths = @(for ($i = 0; $i -lt $columnCount; $i++) { $base })
        }

        $grid = ($Widths | ForEach-Object { "<w:gridCol w:w=`"$_`"/>" }) -join ''
        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add("<w:tbl><w:tblPr><w:tblW w:w=`"$TableWidth`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/><w:tblCellMar><w:top w:w=`"120`" w:type=`"dxa`"/><w:left w:w=`"120`" w:type=`"dxa`"/><w:bottom w:w=`"120`" w:type=`"dxa`"/><w:right w:w=`"120`" w:type=`"dxa`"/></w:tblCellMar><w:tblLook w:val=`"04A0`" w:firstRow=`"1`" w:lastRow=`"0`" w:firstColumn=`"0`" w:lastColumn=`"0`" w:noHBand=`"0`" w:noVBand=`"1`"/></w:tblPr><w:tblGrid>$grid</w:tblGrid>")

        if ($Headers.Count -gt 0) {
            $parts.Add('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                $parts.Add((New-WordTableCell -Text $Headers[$i] -Width $Widths[$i] -Fill $HeaderFill -Header -FontSize $HeaderFontSize -BorderColor $HeaderFill))
            }
            $parts.Add('</w:tr>')
        }

        $rowIndex = 0
        foreach ($row in @($Rows)) {
            $fill = if (($rowIndex % 2) -eq 1) { 'F9FAFD' } else { 'FFFFFF' }
            $cells = @($row)
            $parts.Add('<w:tr>')
            for ($i = 0; $i -lt $columnCount; $i++) {
                $value = if ($i -lt $cells.Count) { $cells[$i] } else { '' }
                $parts.Add((New-WordTableCell -Text (Nvl $value ' ') -Width $Widths[$i] -Fill $fill -FontSize $FontSize))
            }
            $parts.Add('</w:tr>')
            $rowIndex++
        }

        $parts.Add('</w:tbl>')
        return ($parts -join '')
    }

    function Add-Block {
        param([System.Collections.Generic.List[string]]$Body, [string]$Xml)
        if (-not [string]::IsNullOrWhiteSpace($Xml)) {
            $Body.Add($Xml)
        }
    }

    function Add-Spacer {
        param([System.Collections.Generic.List[string]]$Body, [int]$After = 120)
        $Body.Add("<w:p><w:pPr><w:spacing w:after=`"$After`"/></w:pPr></w:p>")
    }

    function New-FillBand {
        param([string]$Fill, [int]$Height = 360)
        return "<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblBorders><w:top w:val=`"nil`"/><w:left w:val=`"nil`"/><w:bottom w:val=`"nil`"/><w:right w:val=`"nil`"/><w:insideH w:val=`"nil`"/><w:insideV w:val=`"nil`"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=`"$script:UsableWidth`"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val=`"$Height`" w:hRule=`"atLeast`"/></w:trPr>$(New-WordTableCell -Text ' ' -Width $script:UsableWidth -Fill $Fill -BorderColor $Fill)</w:tr></w:tbl>"
    }

    function New-CoverPage {
        param([string]$OrgName)

        $paragraphs = @(
            (New-WordParagraph -Text 'MICROSOFT 365 TENANT REVIEW' -Style 'CoverEyebrow' -Color $colors.White -Size 20 -Bold -Alignment center -Before 520 -After 240),
            (New-WordParagraph -Text 'Microsoft 365 Tenant Review' -Style 'CoverTitle' -Color $colors.White -Size 54 -Bold -Alignment center -After 180),
            (New-WordParagraph -Text $OrgName -Style 'CoverSubtitle' -Color 'EAF6FF' -Size 32 -Alignment center -After 360),
            (New-WordParagraph -Text "Review period: $ReviewPeriod" -Style 'CoverMeta' -Color $colors.White -Size 22 -Alignment center -After 120),
            (New-WordParagraph -Text 'Prepared by RTH Tech Services' -Style 'CoverMeta' -Color $colors.White -Size 22 -Alignment center -After 120),
            (New-WordParagraph -Text "Generated: $generatedLabel" -Style 'CoverMeta' -Color $colors.White -Size 22 -Alignment center -After 520),
            (New-WordParagraph -Text 'Confidential client report' -Style 'CoverEyebrow' -Color 'EAF6FF' -Size 17 -Bold -Alignment center -After 360)
        )

        $main = "<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblBorders><w:top w:val=`"nil`"/><w:left w:val=`"nil`"/><w:bottom w:val=`"nil`"/><w:right w:val=`"nil`"/><w:insideH w:val=`"nil`"/><w:insideV w:val=`"nil`"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=`"$script:UsableWidth`"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val=`"10200`" w:hRule=`"atLeast`"/></w:trPr>$(New-WordTableCell -ParagraphsXml $paragraphs -Width $script:UsableWidth -Fill $colors.DeepPurple -BorderColor $colors.DeepPurple)</w:tr></w:tbl>"
        return (New-FillBand -Fill $colors.Blue -Height 320) + $main + (New-WordPageBreak)
    }

    function New-SectionHeader {
        param([string]$Title, [string]$Subtitle = '')
        $paragraphs = @(
            (New-WordParagraph -Text $Title -Style 'SectionHeader' -Color $colors.White -Size 25 -Bold -After 30)
        )
        if ($Subtitle) {
            $paragraphs += (New-WordParagraph -Text $Subtitle -Style 'SectionSubtitle' -Color 'E9DDF0' -Size 17 -After 0)
        }
        return "<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblBorders><w:top w:val=`"nil`"/><w:left w:val=`"nil`"/><w:bottom w:val=`"nil`"/><w:right w:val=`"nil`"/><w:insideH w:val=`"nil`"/><w:insideV w:val=`"nil`"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=`"$script:UsableWidth`"/></w:tblGrid><w:tr>$(New-WordTableCell -ParagraphsXml $paragraphs -Width $script:UsableWidth -Fill $colors.DeepPurple -BorderColor $colors.DeepPurple)</w:tr></w:tbl>"
    }

    function New-Callout {
        param(
            [string]$Label,
            [string]$Text,
            [string]$Fill = 'F8F9FB',
            [string]$Border = '5A2F61'
        )
        $paragraphs = @(
            (New-WordParagraph -Text $Label -Style 'CalloutLabel' -Color $Border -Size 18 -Bold -After 20),
            (New-WordParagraph -Text (Nvl $Text) -Style 'CalloutText' -Color $colors.Text -Size 19 -After 20)
        )
        return "<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/></w:tblPr><w:tblGrid><w:gridCol w:w=`"$script:UsableWidth`"/></w:tblGrid><w:tr>$(New-WordTableCell -ParagraphsXml $paragraphs -Width $script:UsableWidth -Fill $Fill -BorderColor 'DDE3EA' -LeftBorderColor $Border -LeftBorderSize 18)</w:tr></w:tbl>"
    }

    function New-ExecutiveBulletBlock {
        param([object[]]$Bullets)

        $rows = @()
        foreach ($bullet in @($Bullets)) {
            $rows += ,@(Nvl $bullet)
        }

        if ($rows.Count -eq 0) {
            $rows = ,@('No executive summary bullets were generated.')
        }

        return New-WordTable -Headers @() -Rows $rows -Widths @($script:UsableWidth) -TableWidth $script:UsableWidth -FontSize 20
    }

    function New-KpiCardCell {
        param([object]$Kpi, [int]$Width)

        $value = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'value')
        $label = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'label')
        $desc = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'description')
        $status = Nvl (Get-TenantReviewProperty -InputObject $Kpi -Name 'status') 'Informational'
        $statusLabel = Get-StatusLabel $status
        $statusColor = Get-StatusColor $status
        $statusFill = Get-StatusLightColor $status

        $paragraphs = @(
            (New-WordParagraph -Text $value -Style 'KPIValue' -Color $colors.DeepPurple -Size 34 -Bold -After 10),
            (New-WordParagraph -Text $label.ToUpperInvariant() -Style 'KPILabel' -Color $colors.Grey -Size 16 -Bold -After 50),
            (New-WordParagraph -Text $desc -Style 'KPIDescription' -Color $colors.Muted -Size 17 -After 100),
            (New-WordParagraph -Text $statusLabel.ToUpperInvariant() -Style 'StatusBadge' -Color $statusColor -Size 15 -Bold -Fill $statusFill -After 20)
        )

        return New-WordTableCell -ParagraphsXml $paragraphs -Width $Width -Fill $colors.White -BorderColor $colors.Border -TopBorderColor $statusColor -TopBorderSize 22
    }

    function New-KpiGrid {
        param([object[]]$Kpis)

        if (-not $Kpis -or $Kpis.Count -eq 0) {
            return (New-Callout -Label 'KPI data' -Text 'No KPI data was generated for this section.' -Fill $colors.SoftGrey -Border $colors.Grey)
        }

        $colWidth = [int][Math]::Floor($script:UsableWidth / 3)
        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add("<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/><w:tblCellSpacing w:w=`"120`" w:type=`"dxa`"/></w:tblPr><w:tblGrid><w:gridCol w:w=`"$colWidth`"/><w:gridCol w:w=`"$colWidth`"/><w:gridCol w:w=`"$colWidth`"/></w:tblGrid>")
        $items = @($Kpis)
        for ($i = 0; $i -lt $items.Count; $i += 3) {
            $parts.Add('<w:tr>')
            for ($j = 0; $j -lt 3; $j++) {
                $index = $i + $j
                if ($index -lt $items.Count) {
                    $parts.Add((New-KpiCardCell -Kpi $items[$index] -Width $colWidth))
                } else {
                    $parts.Add((New-WordTableCell -Text ' ' -Width $colWidth -Fill $colors.White -BorderColor $colors.White))
                }
            }
            $parts.Add('</w:tr>')
        }
        $parts.Add('</w:tbl>')
        return ($parts -join '')
    }

    function New-NarrativeCard {
        param([object]$Section)

        $headline = Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'headline')
        $plain = Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'plainEnglish')
        $status = Nvl (Get-TenantReviewProperty -InputObject $Section -Name 'status') 'Informational'
        $statusLabel = Get-StatusLabel $status
        $statusColor = Get-StatusColor $status
        $statusFill = Get-StatusLightColor $status

        $paragraphs = @(
            (New-WordParagraph -Text $statusLabel.ToUpperInvariant() -Style 'StatusBadge' -Color $statusColor -Size 15 -Bold -Fill $statusFill -After 80),
            (New-WordParagraph -Text $headline -Style 'NarrativeHeadline' -Color $colors.Text -Size 24 -Bold -After 100),
            (New-WordParagraph -Text $plain -Style 'Normal' -Color $colors.Text -Size 21 -After 60)
        )

        return "<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/></w:tblPr><w:tblGrid><w:gridCol w:w=`"$script:UsableWidth`"/></w:tblGrid><w:tr>$(New-WordTableCell -ParagraphsXml $paragraphs -Width $script:UsableWidth -Fill $colors.White -BorderColor 'E2E8F0' -LeftBorderColor $statusColor -LeftBorderSize 24)</w:tr></w:tbl>"
    }

    function New-RecommendationCard {
        param([object]$Recommendation)

        $priority = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'priority') '-'
        $category = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'category')
        $title = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'title')
        $why = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'why')
        $owner = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'owner')
        $effort = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'effort')
        $impact = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'impact')
        $status = Nvl (Get-TenantReviewProperty -InputObject $Recommendation -Name 'status') 'Watch'
        $statusLabel = Get-StatusLabel $status
        $statusColor = Get-StatusColor $status
        $statusFill = Get-StatusLightColor $status

        $priorityParagraphs = @(
            (New-WordParagraph -Text "P$priority" -Style 'RecommendationPriority' -Color $colors.White -Size 25 -Bold -Alignment center -After 0)
        )
        $bodyParagraphs = @(
            (New-WordParagraph -Text $title -Style 'NarrativeHeadline' -Color $colors.Text -Size 23 -Bold -After 50),
            (New-WordParagraph -Text $why -Style 'Normal' -Color $colors.Muted -Size 19 -After 90),
            (New-WordParagraph -Text "Category: $category    Owner: $owner    Effort: $effort    Impact: $impact" -Style 'TableText' -Color $colors.Muted -Size 17 -After 50),
            (New-WordParagraph -Text $statusLabel.ToUpperInvariant() -Style 'StatusBadge' -Color $statusColor -Size 15 -Bold -Fill $statusFill -After 0)
        )

        $priorityWidth = 900
        $bodyWidth = $script:UsableWidth - $priorityWidth
        return "<w:tbl><w:tblPr><w:tblW w:w=`"$script:UsableWidth`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/></w:tblPr><w:tblGrid><w:gridCol w:w=`"$priorityWidth`"/><w:gridCol w:w=`"$bodyWidth`"/></w:tblGrid><w:tr>$(New-WordTableCell -ParagraphsXml $priorityParagraphs -Width $priorityWidth -Fill $colors.DeepPurple -BorderColor $colors.DeepPurple)$(New-WordTableCell -ParagraphsXml $bodyParagraphs -Width $bodyWidth -Fill $colors.White -BorderColor 'E2E8F0')</w:tr></w:tbl>"
    }

    function Add-AppendixTable {
        param(
            [System.Collections.Generic.List[string]]$Body,
            [string]$Title,
            [string[]]$Headers,
            [object[]]$Rows,
            [int[]]$Widths = $null,
            [string]$EmptyText = 'No records found.'
        )

        Add-Block -Body $Body -Xml (New-WordParagraph -Text $Title -Style 'Heading2' -Color $colors.DeepPurple -Size 25 -Bold -Before 180 -After 90 -KeepNext)
        if ($Rows.Count -gt 0) {
            Add-Block -Body $Body -Xml (New-WordTable -Headers $Headers -Rows $Rows -Widths $Widths -TableWidth $script:UsableWidth -FontSize 16 -HeaderFontSize 16)
        } else {
            Add-Block -Body $Body -Xml (New-Callout -Label $Title -Text $EmptyText -Fill $colors.SoftGrey -Border $colors.Grey)
        }
        Add-Spacer -Body $Body -After 180
    }

    function Get-Section {
        param([string]$Dataset)
        return ($sections | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'dataset') -eq $Dataset } | Select-Object -First 1)
    }

    function Get-SectionKpis {
        param([string]$Dataset)
        $areaName = switch ($Dataset) {
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
        if (-not $areaName -or $Dataset -eq 'LicenseUserAnalysis') { return @() }
        return @($kpis | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'area') -eq $areaName })
    }

    function Test-TenantReviewWordDocText {
        param([string]$DocumentXml)

        $plainText = (($DocumentXml -split '<w:t[^>]*>') | Select-Object -Skip 1 | ForEach-Object {
            $piece = ($_ -split '</w:t>')[0]
            [System.Net.WebUtility]::HtmlDecode($piece)
        }) -join "`n"

        $forbiddenPatterns = @(
            'System\.Object\[\]',
            ('implementation' + 'Status'),
            ('Stub' + ' -'),
            'Microsoft\.Graph\.PowerShell\.Models',
            '@\{',
            '\{\s*"[^"]+"\s*:',
            '\[\s*\{'
        )

        foreach ($pattern in $forbiddenPatterns) {
            if ($plainText -match $pattern) {
                throw "Word document validation failed. Generated DOCX contains forbidden client-facing text matching '$pattern'."
            }
        }

        $requiredHeadings = @(
            'Executive Summary',
            'Key Metrics at a Glance',
            'Tenant Overview',
            'Cost & Licensing Review',
            'Identity & User Review',
            'Exchange & Mail Flow',
            'SharePoint & OneDrive',
            'Teams Collaboration',
            'Devices & Endpoints',
            'Copilot Review',
            'Top Recommendations',
            'Data Coverage',
            'Appendix'
        )

        foreach ($heading in $requiredHeadings) {
            if ($plainText -notmatch [regex]::Escape($heading)) {
                throw "Word document validation failed. Missing required heading: $heading"
            }
        }
    }

    $sections = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'sections')
    $kpis = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'kpis')
    $recommendations = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'recommendations')
    $execBullets = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'executiveSummaryBullets')
    $healthStatement = Nvl (Get-TenantReviewProperty -InputObject $Narrative -Name 'tenantHealthStatement') 'Review the findings and detailed sections in this report.'

    $tenantSummary = Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'summary'
    $orgName = Nvl (Get-TenantReviewProperty -InputObject $tenantSummary -Name 'organizationName') $TenantName

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

    $body = [System.Collections.Generic.List[string]]::new()
    Add-Block -Body $body -Xml (New-CoverPage -OrgName $orgName)

    Add-Block -Body $body -Xml (New-SectionHeader -Title 'Executive Summary' -Subtitle 'High-level tenant health and immediate management focus')
    Add-Block -Body $body -Xml (New-Callout -Label 'Tenant health statement' -Text $healthStatement -Fill $colors.LightBlue -Border $colors.Blue)
    Add-Spacer -Body $body -After 120
    Add-Block -Body $body -Xml (New-ExecutiveBulletBlock -Bullets $execBullets)
    Add-Block -Body $body -Xml (New-WordPageBreak)

    Add-Block -Body $body -Xml (New-SectionHeader -Title 'Key Metrics at a Glance' -Subtitle 'Headline numbers across all reviewed areas')
    Add-Block -Body $body -Xml (New-KpiGrid -Kpis $kpis)
    Add-Block -Body $body -Xml (New-WordPageBreak)

    foreach ($key in $sectionOrder) {
        $section = Get-Section -Dataset $key
        if (-not $section) { continue }

        $title = if ($sectionTitles[$key]) { $sectionTitles[$key] } else { $key }
        Add-Block -Body $body -Xml (New-SectionHeader -Title $title)

        $sectionKpis = @(Get-SectionKpis -Dataset $key)
        if ($sectionKpis.Count -gt 0) {
            Add-Block -Body $body -Xml (New-KpiGrid -Kpis $sectionKpis)
            Add-Spacer -Body $body -After 120
        }

        Add-Block -Body $body -Xml (New-NarrativeCard -Section $section)
        Add-Spacer -Body $body -After 80
        Add-Block -Body $body -Xml (New-Callout -Label 'Why it matters' -Text (Get-TenantReviewProperty -InputObject $section -Name 'businessImpact') -Fill $colors.SoftGrey -Border $colors.DeepPurple)
        Add-Spacer -Body $body -After 80
        Add-Block -Body $body -Xml (New-Callout -Label 'Recommended action' -Text (Get-TenantReviewProperty -InputObject $section -Name 'recommendedAction') -Fill 'EEF3FF' -Border $colors.Blue)
        Add-Block -Body $body -Xml (New-WordPageBreak)
    }

    Add-Block -Body $body -Xml (New-SectionHeader -Title 'Top Recommendations' -Subtitle 'Prioritised actions to improve the tenant')
    if ($recommendations.Count -gt 0) {
        foreach ($recommendation in $recommendations) {
            Add-Block -Body $body -Xml (New-RecommendationCard -Recommendation $recommendation)
            Add-Spacer -Body $body -After 120
        }
    } else {
        Add-Block -Body $body -Xml (New-Callout -Label 'Recommendations' -Text 'No specific recommendations were generated from the available data.' -Fill $colors.SoftGrey -Border $colors.Grey)
    }
    Add-Block -Body $body -Xml (New-WordPageBreak)

    Add-Block -Body $body -Xml (New-SectionHeader -Title 'Data Coverage' -Subtitle 'Status of each data source used in this report')
    Add-Block -Body $body -Xml (New-WordParagraph -Text 'Some data sources may be unavailable depending on licensing, permissions, or configuration. Where a source was unavailable, its status is marked below without exposing raw warnings in the client-facing report.' -Style 'Normal' -Color $colors.Muted -Size 20 -After 160)
    $coverageRows = @(
        foreach ($key in $Datasets.Keys) {
            $dataset = $Datasets[$key]
            $summary = Get-TenantReviewProperty -InputObject $dataset -Name 'summary'
            $warnings = @(Get-TenantReviewProperty -InputObject $dataset -Name 'warnings')
            $skipped = Get-TenantReviewProperty -InputObject $summary -Name 'skipped'
            $status = if ($skipped) { 'Skipped' } elseif ($warnings.Count -gt 0) { "Collected with $($warnings.Count) warning$(if ($warnings.Count -ne 1) { 's' })" } else { 'Collected' }
            ,@($key, $status)
        }
    )
    Add-Block -Body $body -Xml (New-WordTable -Headers @('Dataset','Status') -Rows $coverageRows -Widths @(6400,4040) -FontSize 18)
    Add-Block -Body $body -Xml (New-WordSectionBreakNextPage -SectionXml $portraitSection)

    $script:UsableWidth = 14400
    Add-Block -Body $body -Xml (New-SectionHeader -Title 'Appendix' -Subtitle 'Curated supporting tables for technical reviewers')

    $domainRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'items')) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'id')),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isDefault')),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isVerified')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'authenticationType')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'supportedServices') 'None')
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Domain Inventory' -Headers @('Domain','Default','Verified','Authentication','Supported Services') -Rows $domainRows -Widths @(4200,1400,1400,2200,5200)

    $licenseSummary = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
    $currency = Nvl (Get-TenantReviewProperty -InputObject $licenseSummary -Name 'currency') 'CAD'
    $licenseRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'items')) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName') (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'skuPartNumber'))),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'capabilityStatus') 'N/A'),
                (Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'purchasedUnits')),
                (Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'assignedUnits')),
                (Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'unusedUnits')),
                (Format-WordCurrency (Get-TenantReviewProperty -InputObject $item -Name 'estimatedMonthlyCost') $currency)
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'License SKU Inventory' -Headers @('SKU','Status','Purchased','Assigned','Unused','Monthly Cost') -Rows $licenseRows -Widths @(5000,1800,1700,1700,1700,2500)

    $userReviewRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'items')) {
            $flags = @()
            if ((Get-TenantReviewProperty -InputObject $item -Name 'accountEnabled') -eq $false) { $flags += 'Disabled' }
            if ((Get-TenantReviewProperty -InputObject $item -Name 'isStale') -eq $true) { $flags += 'Stale' }
            if ((Get-TenantReviewProperty -InputObject $item -Name 'isLicensedAndDisabled') -eq $true) { $flags += 'Licensed disabled' }
            if ((Get-TenantReviewProperty -InputObject $item -Name 'isLicensedAndStale') -eq $true) { $flags += 'Licensed stale' }
            if ((Get-TenantReviewProperty -InputObject $item -Name 'isGuest') -eq $true) { $flags += 'Guest' }
            if ($flags.Count -gt 0) {
                ,@(
                    (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                    (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userPrincipalName')),
                    (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userType')),
                    (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'accountEnabled')),
                    ($flags -join ', ')
                )
            }
        }
    )
    Add-AppendixTable -Body $body -Title 'Users Requiring Review' -Headers @('Display Name','UPN','Type','Enabled','Reason') -Rows $userReviewRows -Widths @(3000,4700,1700,1400,3600) -EmptyText 'No users requiring review were identified.'

    $mailboxRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'items') | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'forwardingEnabled') -eq $true }) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'primarySmtpAddress')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'recipientTypeDetails')),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'externalForwardingSuspected')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'forwardingSmtpAddress') (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'forwardingAddress') 'N/A'))
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Mailboxes with Forwarding' -Headers @('Mailbox','Primary SMTP','Type','External','Forward To') -Rows $mailboxRows -Widths @(3000,4200,2500,1300,3400) -EmptyText 'No mailboxes with forwarding were identified.'

    $transportRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'transportRules')) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'name')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'state')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'mode')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'priority')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'conditionsSummary') 'N/A'),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'actionsSummary') 'N/A')
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Transport Rules' -Headers @('Name','State','Mode','Priority','Conditions','Actions') -Rows $transportRows -Widths @(3300,1500,1500,1300,3400,3400) -EmptyText 'No transport rules were collected.'

    $inboxRuleRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'inboxForwardingRules')) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'mailbox')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'name')),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'enabled')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'forwardTo') 'None'),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'redirectTo') 'None'),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'forwardAsAttachmentTo') 'None')
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Inbox Forwarding Rules' -Headers @('Mailbox','Rule','Enabled','Forward To','Redirect To','Attachment To') -Rows $inboxRuleRows -Widths @(3100,2700,1300,2400,2400,2500) -EmptyText 'No inbox forwarding rules were collected.'

    $sharePointRows = @(
        foreach ($item in (@(Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'items') | Sort-Object { -[double](ConvertTo-TenantReviewDecimal -Value (Get-TenantReviewProperty -InputObject $_ -Name 'storageUsageGB') -Default 0) } | Select-Object -First 15)) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'title')),
                (Format-WordDecimal (Get-TenantReviewProperty -InputObject $item -Name 'storageUsageGB') 2),
                (Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'lastContentModifiedDate')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'sharingCapability')),
                (Format-WordNumber (Get-TenantReviewProperty -InputObject $item -Name 'fileCount'))
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Largest SharePoint Sites' -Headers @('Site','Storage GB','Last Modified','Sharing','Files') -Rows $sharePointRows -Widths @(5200,1800,2400,2700,2300)

    $teamRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'items')) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'visibility')),
                (Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'createdDateTime')),
                (Format-WordDate (Get-TenantReviewProperty -InputObject $item -Name 'lastActivityDate')),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'inactive'))
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Teams Inventory' -Headers @('Team','Visibility','Created','Last Activity','Inactive') -Rows $teamRows -Widths @(5200,2100,2400,2500,2200)

    $deviceRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'items') | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'isStale') -eq $true }) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'operatingSystem')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'daysSinceLastSignIn') 'N/A'),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'accountEnabled')),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isManaged')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'complianceState'))
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Stale Devices' -Headers @('Device','OS','Days Since Sign-In','Enabled','Managed','Compliance') -Rows $deviceRows -Widths @(3500,2100,2300,1600,1600,3300) -EmptyText 'No stale devices were identified.'

    $copilotUserRows = @(
        foreach ($item in @(Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'licensedUsers')) {
            ,@(
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'displayName')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'userPrincipalName')),
                (Nvl (Get-TenantReviewProperty -InputObject $item -Name 'assignedCopilotSkuPartNumbers') 'N/A'),
                (Format-WordBool (Get-TenantReviewProperty -InputObject $item -Name 'isStale')),
                (Format-WordDateTime (Get-TenantReviewProperty -InputObject $item -Name 'lastSuccessfulSignInDateTime'))
            )
        }
    )
    Add-AppendixTable -Body $body -Title 'Copilot Licensed Users' -Headers @('Display Name','UPN','Assigned Copilot SKUs','Stale','Last Sign-In') -Rows $copilotUserRows -Widths @(3100,4400,3300,1300,2300) -EmptyText 'No Copilot-licensed users were identified.'

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    $($body -join "`n    ")
    $landscapeSection
  </w:body>
</w:document>
"@

    Test-TenantReviewWordDocText -DocumentXml $documentXml

    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Aptos"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
        <w:color w:val="1F2933"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CoverTitle"><w:name w:val="CoverTitle"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:bCs/><w:sz w:val="54"/><w:szCs w:val="54"/><w:color w:val="FFFFFF"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CoverSubtitle"><w:name w:val="CoverSubtitle"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="EAF6FF"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CoverEyebrow"><w:name w:val="CoverEyebrow"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:spacing w:val="40"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="FFFFFF"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CoverMeta"><w:name w:val="CoverMeta"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="FFFFFF"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="SectionHeader"><w:name w:val="SectionHeader"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:bCs/><w:sz w:val="25"/><w:szCs w:val="25"/><w:color w:val="FFFFFF"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="SectionSubtitle"><w:name w:val="SectionSubtitle"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="17"/><w:szCs w:val="17"/><w:color w:val="E9DDF0"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="280" w:after="140"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="5A2F61"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="180" w:after="100"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="5A2F61"/><w:sz w:val="25"/><w:szCs w:val="25"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="KPIValue"><w:name w:val="KPIValue"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:color w:val="5A2F61"/><w:sz w:val="34"/><w:szCs w:val="34"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="KPILabel"><w:name w:val="KPILabel"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:color w:val="767779"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="KPIDescription"><w:name w:val="KPIDescription"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="5F6570"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="NarrativeHeadline"><w:name w:val="NarrativeHeadline"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:color w:val="1F2933"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CalloutLabel"><w:name w:val="CalloutLabel"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="CalloutText"><w:name w:val="CalloutText"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="StatusBadge"><w:name w:val="StatusBadge"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:sz w:val="15"/><w:szCs w:val="15"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="RecommendationPriority"><w:name w:val="RecommendationPriority"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="25"/><w:szCs w:val="25"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="TableText"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="40" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
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
