[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [string]$ReviewPeriod,

    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),

    [string]$SettingsPath = (Join-Path $PSScriptRoot 'config\sample.settings.json'),

    [switch]$SkipAI,

    [switch]$SkipRender
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Import-LocalScript {
    param([Parameter(Mandatory = $true)][string]$RelativePath)
    . (Join-Path $PSScriptRoot $RelativePath)
}

Import-LocalScript 'src\Shared\Connect-TenantReviewServices.ps1'
Import-LocalScript 'src\Shared\Export-TenantReviewJson.ps1'
Import-LocalScript 'src\Collectors\Get-TenantOverview.ps1'
Import-LocalScript 'src\Collectors\Get-LicenseInventory.ps1'
Import-LocalScript 'src\Collectors\Get-UserInventory.ps1'
Import-LocalScript 'src\Collectors\Get-MailboxInventory.ps1'
Import-LocalScript 'src\Collectors\Get-SharePointInventory.ps1'
Import-LocalScript 'src\Collectors\Get-TeamsInventory.ps1'
Import-LocalScript 'src\Collectors\Get-DeviceInventory.ps1'
Import-LocalScript 'src\Collectors\Get-CopilotInventory.ps1'
Import-LocalScript 'src\AI\Invoke-AINarrative.ps1'
Import-LocalScript 'src\Renderers\New-TenantReviewReport.ps1'
Import-LocalScript 'src\Renderers\New-TenantReviewDeck.ps1'

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$runId = Get-Date -Format 'yyyyMMdd-HHmmss'
$runPath = Join-Path $OutputPath $runId
New-Item -ItemType Directory -Path $runPath -Force | Out-Null

Write-Host "TenantReviewPack" -ForegroundColor Cyan
Write-Host "Tenant: $TenantName"
Write-Host "Review Period: $ReviewPeriod"
Write-Host "Output: $runPath"

$settings = if (Test-Path $SettingsPath) {
    Get-Content -Path $SettingsPath -Raw | ConvertFrom-Json
} else {
    [pscustomobject]@{}
}

Connect-TenantReviewServices -Settings $settings

$datasets = [ordered]@{
    TenantOverview     = Get-TenantOverview -TenantName $TenantName -ReviewPeriod $ReviewPeriod
    LicenseInventory  = Get-LicenseInventory
    UserInventory     = Get-UserInventory
    MailboxInventory  = Get-MailboxInventory
    SharePoint        = Get-SharePointInventory
    Teams             = Get-TeamsInventory
    Devices           = Get-DeviceInventory
    Copilot           = Get-CopilotInventory
}

foreach ($key in $datasets.Keys) {
    Export-TenantReviewJson -InputObject $datasets[$key] -Path (Join-Path $runPath "$key.json")
}

if (-not $SkipAI) {
    $narrative = Invoke-AINarrative -Datasets $datasets -Settings $settings
    Export-TenantReviewJson -InputObject $narrative -Path (Join-Path $runPath 'Narrative.json')
} else {
    $narrative = [pscustomobject]@{
        skipped = $true
        reason = 'SkipAI was specified.'
    }
}

if (-not $SkipRender) {
    New-TenantReviewReport -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
    New-TenantReviewDeck -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
}

Write-Host "Tenant review package generation complete." -ForegroundColor Green
