[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [string]$ReviewPeriod,

    [string]$OutputPath,

    [string]$SettingsPath,

    [string]$ConnectConfigPath,

    [switch]$SkipAI,

    [switch]$SkipRender
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Path $MyInvocation.MyCommand.Path -Parent
} else {
    Get-Location
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $scriptRoot 'output'
}
if (-not $SettingsPath) {
    $SettingsPath = Join-Path $scriptRoot 'Settings.json'
}
if (-not $ConnectConfigPath) {
    $ConnectConfigPath = Join-Path $scriptRoot 'ConnectConfig.json'
}

function Resolve-LocalScript {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $scriptPath = Join-Path $scriptRoot $RelativePath
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

function Test-TenantReviewPlaceholderValue {
    param([Parameter(Mandatory = $false)][object]$Value)

    if ($null -eq $Value) {
        return $true
    }

    $text = $Value.ToString()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $true
    }

    return ($text -match 'contoso|example|YOUR_|00000000-0000-0000-0000-000000000000')
}

function Test-TenantReviewInteractiveHost {
    return ($Host.Name -ne 'ServerRemoteHost' -and [Environment]::UserInteractive)
}

function Read-TenantReviewRequiredOrSkip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string]$SkipToken
    )

    if (-not (Test-TenantReviewInteractiveHost)) {
        throw "$Prompt Run interactively to provide the value, or set the related enabled flag to false to skip it."
    }

    while ($true) {
        $answer = Read-Host "$Prompt Enter a value, or type '$SkipToken' to skip"
        if ($answer -eq $SkipToken) {
            return $null
        }
        if (-not [string]::IsNullOrWhiteSpace($answer)) {
            return $answer
        }
    }
}

function Ensure-TenantReviewObjectProperty {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        Add-Member -InputObject $InputObject -MemberType NoteProperty -Name $Name -Value $Value
    } else {
        $property.Value = $Value
    }
}

function Resolve-TenantReviewSettings {
    param([Parameter(Mandatory = $true)][object]$Settings)

    $sharePointSettings = Get-ObjectPropertyValue -InputObject $Settings -Name 'sharePoint'
    if ($null -eq $sharePointSettings) {
        $sharePointSettings = [pscustomobject]@{
            enabled = $true
        }
        Ensure-TenantReviewObjectProperty -InputObject $Settings -Name 'sharePoint' -Value $sharePointSettings
    }

    $sharePointEnabled = Get-ObjectPropertyValue -InputObject $sharePointSettings -Name 'enabled'
    if ($sharePointEnabled -ne $false) {
        $adminUrl = Get-ObjectPropertyValue -InputObject $sharePointSettings -Name 'adminUrl'
        if (Test-TenantReviewPlaceholderValue -Value $adminUrl) {
            Write-Warning 'SharePoint admin URL is missing or still set to a placeholder.'
            $answer = Read-TenantReviewRequiredOrSkip -Prompt 'What is the SharePoint Admin Portal URL?' -SkipToken 'skip'
            if ($null -eq $answer) {
                Ensure-TenantReviewObjectProperty -InputObject $sharePointSettings -Name 'enabled' -Value $false
                Write-Warning 'SharePoint collection was explicitly skipped for this run.'
            } else {
                Ensure-TenantReviewObjectProperty -InputObject $sharePointSettings -Name 'adminUrl' -Value $answer
                Ensure-TenantReviewObjectProperty -InputObject $sharePointSettings -Name 'enabled' -Value $true
            }
        }
    }

    return $Settings
}

function Resolve-TenantReviewAiRuntime {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Settings,

        [switch]$SkipAI
    )

    if ($SkipAI) {
        return [pscustomobject]@{
            ApiKey = $null
        }
    }

    $aiSettings = Get-ObjectPropertyValue -InputObject $Settings -Name 'ai'
    $aiEnabled = Test-TenantReviewTruthy -Value (Get-ObjectPropertyValue -InputObject $aiSettings -Name 'enabled')
    if (-not $aiEnabled) {
        return [pscustomobject]@{
            ApiKey = $null
        }
    }

    $endpoint = Get-ObjectPropertyValue -InputObject $aiSettings -Name 'endpoint'
    if (Test-TenantReviewPlaceholderValue -Value $endpoint) {
        Write-Warning 'AI is enabled but ai.endpoint is missing or still set to a placeholder.'
        $endpointAnswer = Read-TenantReviewRequiredOrSkip -Prompt 'What is the AI endpoint URL?' -SkipToken 'abort'
        if ($null -eq $endpointAnswer) {
            throw 'AI was enabled but no endpoint was provided.'
        }
        Ensure-TenantReviewObjectProperty -InputObject $aiSettings -Name 'endpoint' -Value $endpointAnswer
    }

    $directApiKey = Get-ObjectPropertyValue -InputObject $aiSettings -Name 'apiKey'
    if (-not $directApiKey) {
        $directApiKey = Get-ObjectPropertyValue -InputObject $aiSettings -Name 'apiKeyValue'
    }
    if (-not $directApiKey) {
        $directApiKey = Get-ObjectPropertyValue -InputObject $aiSettings -Name 'key'
    }

    $apiKeyVariable = Get-ObjectPropertyValue -InputObject $aiSettings -Name 'apiKeyEnvironmentVariable'
    $apiKey = if ($directApiKey) {
        $directApiKey
    } elseif ($apiKeyVariable) {
        $environmentValue = [Environment]::GetEnvironmentVariable($apiKeyVariable)
        if ($environmentValue) {
            $environmentValue
        } elseif ($apiKeyVariable.ToString().Length -ge 32) {
            $apiKeyVariable
        } else {
            $null
        }
    } else {
        $null
    }
    if (-not $apiKey) {
        Write-Warning 'AI is enabled but no API key was found in Settings.json or the configured environment variable.'
        if (-not (Test-TenantReviewInteractiveHost)) {
            throw 'AI is enabled but the API key was unavailable. Add ai.apiKey to Settings.json or set the configured ai.apiKeyEnvironmentVariable before running non-interactively.'
        }

        $secureKey = Read-Host 'Enter the AI API key for this run, or press Enter to abort' -AsSecureString
        $plainKey = [System.Net.NetworkCredential]::new('', $secureKey).Password
        if ([string]::IsNullOrWhiteSpace($plainKey)) {
            throw 'AI was enabled but no API key was provided.'
        }
        $apiKey = $plainKey
    }

    return [pscustomobject]@{
        ApiKey = $apiKey
    }
}

function Get-TenantReviewRunWarnings {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Datasets,

        [Parameter(Mandatory = $false)]
        [object]$Narrative,

        [Parameter(Mandatory = $false)]
        [string[]]$AdditionalWarnings = @()
    )

    $warnings = @()
    foreach ($warning in $AdditionalWarnings) {
        if ($warning) {
            $warnings += [pscustomobject]@{ Source = 'Connection'; Warning = $warning }
        }
    }
    foreach ($key in $Datasets.Keys) {
        foreach ($warning in @(Get-TenantReviewProperty -InputObject $Datasets[$key] -Name 'warnings')) {
            if ($warning) {
                $warnings += [pscustomobject]@{ Source = $key; Warning = $warning }
            }
        }
    }
    foreach ($warning in @(Get-TenantReviewProperty -InputObject $Narrative -Name 'warnings')) {
        if ($warning) {
            $warnings += [pscustomobject]@{ Source = 'Narrative'; Warning = $warning }
        }
    }

    return @($warnings)
}

function New-TenantReviewSkippedDataset {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    [pscustomobject]@{
        dataset     = $Name
        generatedAt = (Get-Date).ToString('o')
        summary     = [pscustomobject]@{
            skipped    = $true
            skipReason = $Reason
        }
        items       = @()
        warnings    = @()
    }
}

function Invoke-TenantReviewCollector {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Collector
    )

    while ($true) {
        $result = & $Collector
        $collectorWarnings = @(
            foreach ($warning in @(Get-TenantReviewProperty -InputObject $result -Name 'warnings')) {
                if ($warning) { $warning }
            }
        )

        if ($collectorWarnings.Count -eq 0) {
            return $result
        }

        Write-Warning "$Name produced $($collectorWarnings.Count) warning(s). The script will not bury these in JSON without your decision."
        foreach ($warning in $collectorWarnings) {
            Write-Warning "[$Name] $warning"
        }

        if (-not (Test-TenantReviewInteractiveHost)) {
            throw "$Name produced warnings in a non-interactive run. Re-run interactively to retry/skip, or fix the missing permission/module/configuration."
        }

        $canRetrySharePointGraphOnly = (
            $Name -eq 'SharePoint' -and
            $script:TenantReviewSharePointSiteSource -eq 'Auto' -and
            (($collectorWarnings -join "`n") -match 'PnP|SharePoint Online site collection|Get-SPOSite')
        )

        if ($canRetrySharePointGraphOnly) {
            $answer = Read-Host "Type 'graph' to retry SharePoint using Graph reports only, 'retry' after fixing access, 'skip' to skip this collector, or 'abort' to stop"
            if ($answer -eq 'graph') {
                $script:TenantReviewSharePointSiteSource = 'GraphReports'
                Write-Warning 'Retrying SharePoint collection using Graph reports only. PnP/SPO tenant site detail will not be collected for this run.'
                continue
            }
        } else {
            $answer = Read-Host "Type 'retry' after fixing the issue, 'skip' to skip this collector, or 'abort' to stop"
        }

        switch ($answer) {
            'retry' {
                continue
            }
            'skip' {
                $reason = ($collectorWarnings -join ' | ')
                $script:TenantReviewRunDecisions += [pscustomobject]@{
                    Source   = $Name
                    Decision = 'Skipped'
                    Reason   = $reason
                }
                return (New-TenantReviewSkippedDataset -Name $Name -Reason $reason)
            }
            'abort' {
                throw "$Name produced warnings and the run was aborted by user decision."
            }
            default {
                Write-Warning "Unknown response '$answer'. Choose retry, skip, or abort."
            }
        }
    }
}

function Invoke-TenantReviewConnection {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Settings,

        [Parameter(Mandatory = $true)]
        [object]$ConnectConfig
    )

    while ($true) {
        $status = Connect-TenantReviewServices -Settings $Settings -ConnectConfig $ConnectConfig
        $connectionWarnings = @(
            foreach ($warning in @($status.Warnings)) {
                if ($warning) { $warning }
            }
        )

        if ($connectionWarnings.Count -eq 0) {
            return $status
        }

        Write-Warning "Connection setup produced $($connectionWarnings.Count) warning(s). The script will not continue without your decision."
        foreach ($warning in $connectionWarnings) {
            Write-Warning "[Connection] $warning"
        }

        if (-not (Test-TenantReviewInteractiveHost)) {
            throw 'Connection setup produced warnings in a non-interactive run. Re-run interactively to retry/continue, or fix the missing permission/module/configuration.'
        }

        $answer = Read-Host "Type 'retry' after fixing the issue, 'continue' to continue but mark the run incomplete, or 'abort' to stop"
        switch ($answer) {
            'retry' {
                continue
            }
            'continue' {
                return $status
            }
            'abort' {
                throw 'Connection setup produced warnings and the run was aborted by user decision.'
            }
            default {
                Write-Warning "Unknown response '$answer'. Choose retry, continue, or abort."
            }
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
. (Resolve-LocalScript 'src\Renderers\New-TenantReviewWordDoc.ps1')

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
    Resolve-TenantReviewSettings -Settings (Import-TenantReviewJsonFile -Path $SettingsPath -Description 'settings')
} else {
    throw "Settings file was not found at '$SettingsPath'. Create TenantReviewPack\Settings.json or pass -SettingsPath to the correct root settings file. The script does not fall back to sample settings."
}
$aiRuntime = Resolve-TenantReviewAiRuntime -Settings $settings -SkipAI:$SkipAI

$connectConfig = if (Test-Path $ConnectConfigPath) {
    Import-TenantReviewJsonFile -Path $ConnectConfigPath -Description 'connection configuration'
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
            $priceMapPath = Join-Path $scriptRoot $configuredPriceMapPath
        }
    }
}

$connectionStatus = Invoke-TenantReviewConnection -Settings $settings -ConnectConfig $connectConfig
Write-Host ("Graph connected: {0}; auth mode: {1}; certificate found: {2}" -f $connectionStatus.GraphConnected, $connectionStatus.GraphAuthMode, $connectionStatus.CertificateFound)

$includeInboxRules = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('exchangeOnline', 'includeInboxRules') -Default $false)
$includeMailboxStatistics = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('exchangeOnline', 'includeMailboxStatistics') -Default $false)
$inboxRuleMailboxLimit = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('exchangeOnline', 'inboxRuleMailboxLimit') -Default 200)
$includeUserSignInActivity = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('users', 'includeSignInActivity') -Default $true)
$includeOneDrive = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('sharePoint', 'includeOneDrive') -Default $true)
$sharePointReportPeriodDays = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('sharePoint', 'reportPeriodDays') -Default 90)
$script:TenantReviewSharePointSiteSource = Get-TenantReviewConfigValue -Settings $settings -Path @('sharePoint', 'siteSource') -Default 'Auto'
if ($script:TenantReviewSharePointSiteSource -notin @('Auto', 'PnP', 'SPO', 'GraphReports')) {
    throw "sharePoint.siteSource must be one of Auto, PnP, SPO, or GraphReports. Current value is '$script:TenantReviewSharePointSiteSource'."
}
$teamsReportPeriodDays = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('teams', 'reportPeriodDays') -Default 90)
$includeOwnersAndMembers = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('teams', 'includeOwnersAndMembers') -Default $false)
$deviceStaleAfterDays = [int](Get-TenantReviewConfigValue -Settings $settings -Path @('devices', 'staleAfterDays') -Default 90)
$includeIntune = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('devices', 'includeIntune') -Default $true)
$includeCopilotUsageReport = Test-TenantReviewTruthy -Value (Get-TenantReviewConfigValue -Settings $settings -Path @('copilot', 'includeUsageReport') -Default $true)

$script:TenantReviewRunDecisions = @()
$datasets = [ordered]@{}
$datasets['TenantOverview'] = Invoke-TenantReviewCollector -Name 'TenantOverview' -Collector { Get-TenantOverview -TenantName $TenantName -ReviewPeriod $ReviewPeriod }
$datasets['LicenseInventory'] = Invoke-TenantReviewCollector -Name 'LicenseInventory' -Collector { Get-LicenseInventory -PriceMapPath $priceMapPath -DefaultCurrency $defaultCurrency }
$datasets['UserInventory'] = Invoke-TenantReviewCollector -Name 'UserInventory' -Collector { Get-UserInventory -IncludeSignInActivity:$includeUserSignInActivity }
$datasets['LicenseUserAnalysis'] = Get-LicenseUserAnalysis -LicenseInventory $datasets.LicenseInventory -UserInventory $datasets.UserInventory
$datasets['MailboxInventory'] = Invoke-TenantReviewCollector -Name 'MailboxInventory' -Collector { Get-MailboxInventory -IncludeInboxRules:$includeInboxRules -InboxRuleMailboxLimit $inboxRuleMailboxLimit -IncludeMailboxStatistics:$includeMailboxStatistics }
$datasets['SharePoint'] = Invoke-TenantReviewCollector -Name 'SharePoint' -Collector { Get-SharePointInventory -ReportPeriodDays $sharePointReportPeriodDays -IncludeOneDrive:$includeOneDrive -SiteSource $script:TenantReviewSharePointSiteSource }
$datasets['Teams'] = Invoke-TenantReviewCollector -Name 'Teams' -Collector { Get-TeamsInventory -ReportPeriodDays $teamsReportPeriodDays -IncludeOwnersAndMembers:$includeOwnersAndMembers }
$datasets['Devices'] = Invoke-TenantReviewCollector -Name 'Devices' -Collector { Get-DeviceInventory -StaleAfterDays $deviceStaleAfterDays -IncludeIntune:$includeIntune }
$datasets['Copilot'] = Invoke-TenantReviewCollector -Name 'Copilot' -Collector { Get-CopilotInventory -LicenseInventory $datasets.LicenseInventory -UserInventory $datasets.UserInventory -PriceMapPath $priceMapPath -IncludeUsageReport:$includeCopilotUsageReport }

foreach ($key in $datasets.Keys) {
    Export-TenantReviewJson -InputObject $datasets[$key] -Path (Join-Path $runPath "$key.json")
}

if (-not $SkipAI) {
    $narrative = Invoke-AINarrative -Datasets $datasets -Settings $settings -RuntimeApiKey $aiRuntime.ApiKey
} else {
    $narrative = Invoke-AINarrative -Datasets $datasets -Settings ([pscustomobject]@{
        ai = [pscustomobject]@{
            enabled = $false
        }
    })
}
Export-TenantReviewJson -InputObject $narrative -Path (Join-Path $runPath 'Narrative.json')

if (-not $SkipRender) {
    try {
        New-TenantReviewReport -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
    } catch {
        throw "Report rendering failed after data export completed. $($_.Exception.Message)"
    }

    try {
        New-TenantReviewDeck -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
    } catch {
        throw "Deck outline rendering failed after data export completed. $($_.Exception.Message)"
    }

    try {
        New-TenantReviewWordDoc -TenantName $TenantName -ReviewPeriod $ReviewPeriod -Datasets $datasets -Narrative $narrative -OutputPath $runPath
    } catch {
        throw "Word document rendering failed after data export completed. $($_.Exception.Message)"
    }
}

$runWarnings = @(Get-TenantReviewRunWarnings -Datasets $datasets -Narrative $narrative -AdditionalWarnings $connectionStatus.Warnings)
if ($runWarnings.Count -gt 0) {
    Write-Warning "Tenant review package generation completed with $($runWarnings.Count) warning(s). Review the warnings below and the JSON warning arrays before treating the package as complete."
    foreach ($warning in $runWarnings) {
        Write-Warning ("[{0}] {1}" -f $warning.Source, $warning.Warning)
    }
    exit 2
}

if ($script:TenantReviewRunDecisions.Count -gt 0) {
    Write-Warning "Tenant review package generation completed with $($script:TenantReviewRunDecisions.Count) user-approved skip decision(s). The package is not a full successful tenant review."
    foreach ($decision in $script:TenantReviewRunDecisions) {
        Write-Warning ("[{0}] {1}: {2}" -f $decision.Source, $decision.Decision, $decision.Reason)
    }
    exit 2
}

Write-Host "Tenant review package generation complete." -ForegroundColor Green
