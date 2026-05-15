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

    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $outlinePath = Join-Path $OutputPath 'TenantReviewDeckOutline.md'
    $sections = @(Get-TenantReviewProperty -InputObject $Narrative -Name 'sections')
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# Tenant Review Deck Outline")
    $lines.Add("")
    $lines.Add("Tenant: $TenantName")
    $lines.Add("")
    $lines.Add("Review period: $ReviewPeriod")
    $lines.Add("")

    $slideNumber = 1
    foreach ($section in $sections) {
        $lines.Add("## Slide $slideNumber - $(Get-TenantReviewProperty -InputObject $section -Name 'dataset')")
        $lines.Add("")
        $lines.Add("Title: $(Get-TenantReviewProperty -InputObject $section -Name 'headline')")
        $lines.Add("")
        $lines.Add("Big number: $(Get-TenantReviewProperty -InputObject $section -Name 'headline')")
        $lines.Add("")
        $lines.Add("Plain-English explanation: $(Get-TenantReviewProperty -InputObject $section -Name 'plainEnglish')")
        $lines.Add("")
        $lines.Add("Business impact: $(Get-TenantReviewProperty -InputObject $section -Name 'businessImpact')")
        $lines.Add("")
        $lines.Add("Recommended action: $(Get-TenantReviewProperty -InputObject $section -Name 'recommendedAction')")
        $lines.Add("")
        $slideNumber++
    }

    $newline = [Environment]::NewLine
    Set-Content -Path $outlinePath -Value ($lines -join $newline) -Encoding UTF8
    Write-Host "Created deck outline: $outlinePath"
}
