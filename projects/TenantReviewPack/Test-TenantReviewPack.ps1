[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

Write-Host "TenantReviewPack validation"
Write-Host "Project root: $projectRoot"

$scriptFiles = @(Get-ChildItem -Path $projectRoot -Recurse -Filter '*.ps1' -File)
foreach ($scriptFile in $scriptFiles) {
    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName, [ref]$tokens, [ref]$parseErrors) | Out-Null
    if ($parseErrors.Count -gt 0) {
        foreach ($parseError in $parseErrors) {
            $failures.Add(("Parser error in {0}: {1}" -f $scriptFile.FullName, $parseError.Message))
        }
    }
}

$safeDotSourceFiles = @(Get-ChildItem -Path (Join-Path $projectRoot 'src') -Recurse -Filter '*.ps1' -File)
foreach ($scriptFile in $safeDotSourceFiles) {
    try {
        . $scriptFile.FullName
    } catch {
        $failures.Add(("Failed to dot-source {0}: {1}" -f $scriptFile.FullName, $_.Exception.Message))
    }
}

$expectedFunctions = @(
    'Get-TenantReviewProperty',
    'Connect-TenantReviewServices',
    'Export-TenantReviewJson',
    'Get-TenantOverview',
    'Get-LicenseInventory',
    'Get-UserInventory',
    'Get-MailboxInventory',
    'Get-SharePointInventory',
    'Get-TeamsInventory',
    'Get-DeviceInventory',
    'Get-CopilotInventory',
    'Get-LicenseUserAnalysis',
    'Invoke-AINarrative',
    'New-TenantReviewReport',
    'New-TenantReviewDeck'
)

foreach ($functionName in $expectedFunctions) {
    if (-not (Get-Command -Name $functionName -CommandType Function -ErrorAction SilentlyContinue)) {
        $failures.Add("Expected function was not loaded: $functionName")
    }
}

try {
    $licenseInventory = [pscustomobject]@{
        dataset = 'LicenseInventory'
        summary = [pscustomobject]@{
            totalPurchased = 10
            totalUnused = 2
        }
        items = @(
            [pscustomobject]@{
                skuId = '11111111-1111-1111-1111-111111111111'
                skuPartNumber = 'Microsoft_365_Copilot'
                displayName = 'Microsoft 365 Copilot'
                purchasedUnits = 10
                assignedUnits = 8
                unusedUnits = 2
                estimatedMonthlyCost = 400
                estimatedUnusedMonthlyCost = 80
                estimatedUnusedAnnualCost = 960
            }
        )
        warnings = @()
    }
    $userInventory = [pscustomobject]@{
        dataset = 'UserInventory'
        summary = [pscustomobject]@{
            totalUsers = 2
            licensedDisabledUsers = 1
            licensedStaleUsers = 1
        }
        items = @(
            [pscustomobject]@{
                displayName = 'Disabled User'
                userPrincipalName = 'disabled@example.com'
                assignedSkuIds = @('11111111-1111-1111-1111-111111111111')
                isLicensed = $true
                isGuest = $false
                isStale = $true
                isLicensedAndDisabled = $true
                isLicensedAndStale = $true
                lastSuccessfulSignInDateTime = '2025-01-01T00:00:00Z'
            }
        )
        warnings = @()
    }

    $analysis = Get-LicenseUserAnalysis -LicenseInventory $licenseInventory -UserInventory $userInventory
    if ((Get-TenantReviewProperty -InputObject (Get-TenantReviewProperty -InputObject $analysis -Name 'summary') -Name 'disabledLicensedUserCount') -ne 1) {
        $failures.Add('Analyzer mock test did not detect disabled licensed user.')
    }

    $datasets = [ordered]@{
        TenantOverview = [pscustomobject]@{ dataset = 'TenantOverview'; summary = [pscustomobject]@{}; items = @(); warnings = @() }
        LicenseInventory = $licenseInventory
        UserInventory = $userInventory
        LicenseUserAnalysis = $analysis
        MailboxInventory = [pscustomobject]@{ dataset = 'MailboxInventory'; summary = [pscustomobject]@{ mailboxesForwardingExternally = 0 }; items = @(); warnings = @() }
        SharePoint = [pscustomobject]@{ dataset = 'SharePointInventory'; summary = [pscustomobject]@{ totalSites = 1; externalSharingEnabledSites = 0 }; items = @(); warnings = @() }
        Teams = [pscustomobject]@{ dataset = 'TeamsInventory'; summary = [pscustomobject]@{ inactiveTeams = 0 }; items = @(); warnings = @() }
        Devices = [pscustomobject]@{ dataset = 'DeviceInventory'; summary = [pscustomobject]@{ staleDevices = 0 }; items = @(); warnings = @() }
        Copilot = [pscustomobject]@{ dataset = 'CopilotInventory'; summary = [pscustomobject]@{ copilotUnused = 2 }; items = @(); warnings = @() }
    }
    $narrative = Invoke-AINarrative -Datasets $datasets -Settings ([pscustomobject]@{ ai = [pscustomobject]@{ enabled = $false } })
    if (@(Get-TenantReviewProperty -InputObject $narrative -Name 'sections').Count -eq 0) {
        $failures.Add('Narrative mock test did not produce sections.')
    }

    $renderPath = Join-Path ([System.IO.Path]::GetTempPath()) ("TenantReviewPackTest-" + [guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $renderPath -Force | Out-Null
    try {
        New-TenantReviewReport -TenantName 'Test Tenant' -ReviewPeriod 'Test Period' -Datasets $datasets -Narrative $narrative -OutputPath $renderPath
        New-TenantReviewDeck -TenantName 'Test Tenant' -ReviewPeriod 'Test Period' -Datasets $datasets -Narrative $narrative -OutputPath $renderPath
        foreach ($expectedOutput in @('TenantReviewReport.md', 'TenantReviewReport.html', 'TenantReviewDeckOutline.md')) {
            if (-not (Test-Path (Join-Path $renderPath $expectedOutput))) {
                $failures.Add("Renderer mock test did not create $expectedOutput.")
            }
        }
    } finally {
        if (Test-Path $renderPath) {
            Remove-Item -Path $renderPath -Recurse -Force
        }
    }
} catch {
    $failures.Add("Mock analyzer/narrative/renderer tests failed: $($_.Exception.Message)")
}

$sourceFiles = @(Get-ChildItem -Path (Join-Path $projectRoot 'src') -Recurse -File -Include '*.ps1')
$stubPatterns = @('Stub -', 'implementationStatus', 'placeholder')
foreach ($sourceFile in $sourceFiles) {
    $content = Get-Content -Path $sourceFile.FullName -Raw
    foreach ($pattern in $stubPatterns) {
        if ($content -match [regex]::Escape($pattern)) {
            $failures.Add("Source file still contains '$pattern': $($sourceFile.FullName)")
        }
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'Validation failed.' -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host " - $failure" -ForegroundColor Red
    }

    [pscustomobject]@{
        passed = $false
        checkedScripts = $scriptFiles.Count
        failures = @($failures)
    }
    exit 1
}

Write-Host 'Validation passed.' -ForegroundColor Green
[pscustomobject]@{
    passed = $true
    checkedScripts = $scriptFiles.Count
    functions = $expectedFunctions
}
