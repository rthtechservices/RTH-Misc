[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [string]$ReviewPeriod,

    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),

    [string]$SettingsPath = (Join-Path $PSScriptRoot 'config\sample.settings.json'),

    [string]$ConnectConfigPath = (Join-Path $PSScriptRoot 'ConnectConfig.json'),

    [switch]$SkipAI,

    [switch]$SkipRender
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-LocalScript {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $scriptPath = Join-Path $PSScriptRoot $RelativePath
    if (-not (Test-Path $scriptPath)) {
        throw "Required script was not found: $scriptPath"
    }

    return $scriptPath
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $false)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Import-TenantReviewJsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    try {
        return Get-Content -Path $Path -Raw | ConvertFrom-Json
    } catch {
        throw "Unable to load $Description from '$Path'. Ensure the file is valid JSON. $($_.Exception.Message)"
    }
}

function New-InteractiveConnectConfigFromSettings {
    param([Parameter(Mandatory = $false)][object]$Settings)

    $graphSettings = Get-ObjectPropertyValue -InputObject $Settings -Name 'graph'
    $authMode = Get-ObjectPropertyValue -InputObject $graphSettings -Name 'authMode'
    if (-not $authMode) {
        return $null
    }

    if ($authMode -ne 'Interactive') {
        return $null
    }

    return [pscustomobject]@{
        auth = [pscustomobject]@{
            mode = 'Interactive'
        }
    }
}

. (Resolve-LocalScript 'src\Shared\TenantReview.Helpers.ps1')
. (Resolve-LocalScript 'src\Shared\Connect-TenantReviewServices.ps1')
. (Resolve-LocalScript 'src\Shared\Export-TenantReviewJson.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-TenantOverview.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-LicenseInventory.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-UserInventory.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-MailboxInventory.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-SharePointInventory.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-TeamsInventory.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-DeviceInventory.ps1')
. (Resolve-LocalScript 'src\Collectors\Get-CopilotInventory.ps1')
. (Resolve-LocalScript 'src\Analyzers\Get-LicenseUserAnalysis.ps1')
. (Resolve-LocalScript 'src\AI\Invoke-AINarrative.ps1')
. (Resolve-LocalScript 'src\Renderers\New-TenantReviewReport.ps1')
. (Resolve-LocalScript 'src\Renderers\New-TenantReviewDeck.ps1')

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
    Import-TenantReviewJsonFile -Path $SettingsPath -Description 'settings'
} else {
    [pscustomobject]@{}
}

$connectConfig = if (Test-Path $ConnectConfigPath) {
    Import-TenantReviewJsonFile -Path $ConnectConfigPath -Description 'connection configuration'
} else {
    $fallbackConfig = New-InteractiveConnectConfigFromSettings -Settings $settings
    if ($null -ne $fallbackConfig) {
        Write-Warning "ConnectConfigPath '$ConnectConfigPath' was not found. Falling back to Interactive mode from SettingsPath for backward compatibility."
        $fallbackConfig
    } else {
        throw @"
Connection configuration was not found at '$ConnectConfigPath'.
AppCertificate mode requires a ConnectConfig.json file shaped like:
{
  "auth": {
    "mode": "AppCertificate",
    "tenantId": "...",
    "clientId": "...",
    "certificateThumbprint": "...",
    "exchangeOrganization": "contoso.onmicrosoft.com"
  }
}
Create the file or pass -ConnectConfigPath to the correct JSON file.
"@
    }
}

$licensePricingSettings = Get-ObjectPropertyValue -InputObject $settings -Name 'licensePricing'
$licensePricingEnabled = Get-ObjectPropertyValue -InputObject $licensePricingSettings -Name 'enabled'
$priceMapPath = $null
$defaultCurrency = Get-ObjectPropertyValue -InputObject $licensePricingSettings -Name 'defaultCurrency'
if (-not $defaultCurrency) {
    $defaultCurrency = 'CAD'
}

if ($licensePricingEnabled -ne $false) {
    $configuredPriceMapPath = Get-ObjectPropertyValue -InputObject $licensePricingSettings -Name 'priceMapPath'
    if ($configuredPriceMapPath) {
        if ([System.IO.Path]::IsPathRooted($configuredPriceMapPath)) {
            $priceMapPath = $configuredPriceMapPath
        } else {
            $priceMapPath = Join-Path $PSScriptRoot $configuredPriceMapPath
        }
    }
}

$connectionStatus = Connect-TenantReviewServices -Settings $settings -ConnectConfig $connectConfig
Write-Host ("Graph connected: {0}; auth mode: {1}; certificate found: {2}" -f $connectionStatus.GraphConnected, $connectionStatus.GraphAuthMode, $connectionStatus.CertificateFound)
if ($connectionStatus.Warnings.Count -gt 0) {
    foreach ($warning in $connectionStatus.Warnings) {
        Write-Warning $warning
    }
}

$includeInboxRules = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('exchangeOnline', 'includeInboxRules') -Default $false)
$includeMailboxStatistics = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('exchangeOnline', 'includeMailboxStatistics') -Default $false)
$inboxRuleMailboxLimit = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('exchangeOnline', 'inboxRuleMailboxLimit') -Default 200)
$includeOneDrive = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('sharePoint', 'includeOneDrive') -Default $true)
$sharePointReportPeriodDays = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('sharePoint', 'reportPeriodDays') -Default 90)
$teamsReportPeriodDays = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('teams', 'reportPeriodDays') -Default 90)
$includeOwnersAndMembers = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('teams', 'includeOwnersAndMembers') -Default $false)
$deviceStaleAfterDays = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('devices', 'staleAfterDays') -Default 90)
$includeIntune = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('devices', 'includeIntune') -Default $true)

$datasets = [ordered]@{}
$datasets['TenantOverview'] = Get-TenantOverview -TenantName $TenantName -ReviewPeriod $ReviewPeriod
$datasets['LicenseInventory'] = Get-LicenseInventory -PriceMapPath $priceMapPath -DefaultCurrency $defaultCurrency
$datasets['UserInventory'] = Get-UserInventory
$datasets['LicenseUserAnalysis'] = Get-LicenseUserAnalysis -LicenseInventory $datasets.LicenseInventory -UserInventory $datasets.UserInventory
$datasets['MailboxInventory'] = Get-MailboxInventory -IncludeInboxRules:$includeInboxRules -InboxRuleMailboxLimit $inboxRuleMailboxLimit -IncludeMailboxStatistics:$includeMailboxStatistics
$datasets['SharePoint'] = Get-SharePointInventory -ReportPeriodDays $sharePointReportPeriodDays -IncludeOneDrive:$includeOneDrive
$datasets['Teams'] = Get-TeamsInventory -ReportPeriodDays $teamsReportPeriodDays -IncludeOwnersAndMembers:$includeOwnersAndMembers
$datasets['Devices'] = Get-DeviceInventory -StaleAfterDays $deviceStaleAfterDays -IncludeIntune:$includeIntune
$datasets['Copilot'] = Get-CopilotInventory -LicenseInventory $datasets.LicenseInventory -UserInventory $datasets.UserInventory -PriceMapPath $priceMapPath

foreach ($key in $datasets.Keys) {
    Export-TenantReviewJson -InputObject $datasets[$key] -Path (Join-Path $runPath "$key.json")
}

if (-not $SkipAI) {
    $narrative = Invoke-AINarrative -Datasets $datasets -Settings $settings
} else {
    $narrative = Invoke-AINarrative -Datasets $datasets -Settings ([pscustomobject]@{
        ai = [pscustomobject]@{
            enabled = $false
        }
    })
    $narrative.warnings = @($narrative.warnings + 'SkipAI was specified; generated local deterministic narrative only.')
}
Export-TenantReviewJson -InputObject $narrative -Path (Join-Path $runPath 'Narrative.json')

if (-not $SkipRender) {
    try {
        New-TenantReviewReport -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
    } catch {
        Write-Warning "Report rendering failed after data export completed. $($_.Exception.Message)"
    }

    try {
        New-TenantReviewDeck -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
    } catch {
        Write-Warning "Deck outline rendering failed after data export completed. $($_.Exception.Message)"
    }
}

Write-Host "Tenant review package generation complete." -ForegroundColor Green
