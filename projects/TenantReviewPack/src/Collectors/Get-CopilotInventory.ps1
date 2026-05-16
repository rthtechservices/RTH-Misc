function Get-CopilotInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$LicenseInventory,

        [Parameter(Mandatory = $false)]
        [object]$UserInventory,

        [Parameter(Mandatory = $false)]
        [string]$PriceMapPath,

        [switch]$IncludeUsageReport,

        [switch]$IncludeRaw
    )

    $warnings = @()
    $usageReportAvailable = $false
    $usageReportStatus = if ($IncludeUsageReport) { 'Requested' } else { 'NotRequested' }
    $licenseItems = @(Get-TenantReviewProperty -InputObject $LicenseInventory -Name 'items')
    if ($licenseItems.Count -eq 0) {
        $warnings += 'License inventory was not passed or contained no items; falling back to Get-MgSubscribedSku.'
        if (Get-Command -Name Get-MgSubscribedSku -ErrorAction SilentlyContinue) {
            try {
                $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
                $licenseItems = foreach ($sku in $skus) {
                    $prepaid = Get-TenantReviewProperty -InputObject $sku -Name 'PrepaidUnits'
                    [pscustomobject]@{
                        skuId                       = Get-TenantReviewProperty -InputObject $sku -Name 'SkuId'
                        skuPartNumber              = Get-TenantReviewProperty -InputObject $sku -Name 'SkuPartNumber'
                        displayName                 = Get-TenantReviewProperty -InputObject $sku -Name 'SkuPartNumber'
                        purchasedUnits             = [int](Get-TenantReviewProperty -InputObject $prepaid -Name 'Enabled')
                        assignedUnits              = [int](Get-TenantReviewProperty -InputObject $sku -Name 'ConsumedUnits')
                        unusedUnits                = [int](Get-TenantReviewProperty -InputObject $prepaid -Name 'Enabled') - [int](Get-TenantReviewProperty -InputObject $sku -Name 'ConsumedUnits')
                        estimatedMonthlyCost       = $null
                        estimatedUnusedMonthlyCost = $null
                    }
                }
            } catch {
                $warnings += "Unable to fall back to Get-MgSubscribedSku for Copilot inventory. $($_.Exception.Message)"
            }
        }
    }

    $copilotSkus = @($licenseItems | Where-Object {
        $skuPartNumber = Get-TenantReviewProperty -InputObject $_ -Name 'skuPartNumber'
        $displayName = Get-TenantReviewProperty -InputObject $_ -Name 'displayName'
        ($skuPartNumber -and $skuPartNumber.ToString() -match 'COPILOT') -or
        ($displayName -and $displayName.ToString() -match 'Copilot')
    })

    if ($copilotSkus.Count -eq 0) {
        $warnings += 'No Copilot-related SKUs were found in license inventory.'
    }

    $copilotSkuIds = @($copilotSkus | ForEach-Object {
        $skuId = Get-TenantReviewProperty -InputObject $_ -Name 'skuId'
        if ($skuId) { $skuId.ToString() }
    } | Where-Object { $_ })
    $skuPartById = @{}
    foreach ($sku in $copilotSkus) {
        $skuId = Get-TenantReviewProperty -InputObject $sku -Name 'skuId'
        if ($skuId) {
            $skuPartById[$skuId.ToString()] = Get-TenantReviewProperty -InputObject $sku -Name 'skuPartNumber'
        }
    }

    $licensedUsers = @()
    $userItems = @(Get-TenantReviewProperty -InputObject $UserInventory -Name 'items')
    foreach ($user in $userItems) {
        $assignedSkuIds = @(Get-TenantReviewProperty -InputObject $user -Name 'assignedSkuIds')
        $assignedCopilotSkuIds = @($assignedSkuIds | Where-Object { $copilotSkuIds -contains $_.ToString() })
        if ($assignedCopilotSkuIds.Count -gt 0) {
            $licensedUsers += [pscustomobject]@{
                userPrincipalName            = Get-TenantReviewProperty -InputObject $user -Name 'userPrincipalName'
                displayName                  = Get-TenantReviewProperty -InputObject $user -Name 'displayName'
                assignedCopilotSkuPartNumbers = @($assignedCopilotSkuIds | ForEach-Object { $skuPartById[$_.ToString()] })
                isStale                      = Get-TenantReviewProperty -InputObject $user -Name 'isStale'
                lastSuccessfulSignInDateTime = Get-TenantReviewProperty -InputObject $user -Name 'lastSuccessfulSignInDateTime'
            }
        }
    }

    $usageReport = @()
    if ($IncludeUsageReport) {
        foreach ($commandName in @('Get-MgBetaReportMicrosoft365CopilotUsageUserDetail', 'Get-MgReportMicrosoft365CopilotUsageUserDetail')) {
            try {
                $usageReport = @(Invoke-TenantReviewGraphCsvReport -CommandName $commandName -ReportPeriodDays 90)
                if ($usageReport.Count -gt 0) {
                    $usageReportAvailable = $true
                    $usageReportStatus = 'Collected'
                    break
                }
            } catch {
                $usageReportStatus = "Unavailable: $($_.Exception.Message)"
                continue
            }
        }
        if ($usageReport.Count -eq 0 -and $usageReportStatus -eq 'Requested') {
            $usageReportStatus = 'NoRows'
        }
        if (-not $usageReportAvailable) {
            $warnings += "Copilot usage report was requested but unavailable or returned no rows. $usageReportStatus"
        }
    }

    $activeCopilotUsers = 0
    if ($usageReport.Count -gt 0) {
        $activeCopilotUsers = @($usageReport | Where-Object {
            $lastActivity = Get-TenantReviewProperty -InputObject $_ -Name 'Last Activity Date'
            -not [string]::IsNullOrWhiteSpace($lastActivity)
        }).Count
    }

    $items = foreach ($sku in $copilotSkus) {
        [pscustomobject]@{
            skuId                       = Get-TenantReviewProperty -InputObject $sku -Name 'skuId'
            skuPartNumber              = Get-TenantReviewProperty -InputObject $sku -Name 'skuPartNumber'
            displayName                 = Get-TenantReviewProperty -InputObject $sku -Name 'displayName'
            purchasedUnits             = Get-TenantReviewProperty -InputObject $sku -Name 'purchasedUnits'
            assignedUnits              = Get-TenantReviewProperty -InputObject $sku -Name 'assignedUnits'
            unusedUnits                = Get-TenantReviewProperty -InputObject $sku -Name 'unusedUnits'
            estimatedMonthlyCost       = Get-TenantReviewProperty -InputObject $sku -Name 'estimatedMonthlyCost'
            estimatedUnusedMonthlyCost = Get-TenantReviewProperty -InputObject $sku -Name 'estimatedUnusedMonthlyCost'
        }
    }

    $estimatedMonthlyCost = Get-TenantReviewPropertySum -Items $items -Property 'estimatedMonthlyCost'
    $estimatedUnusedMonthlyCost = Get-TenantReviewPropertySum -Items $items -Property 'estimatedUnusedMonthlyCost'
    $copilotPurchased = Get-TenantReviewPropertySum -Items $items -Property 'purchasedUnits' -Default 0
    $copilotAssigned = Get-TenantReviewPropertySum -Items $items -Property 'assignedUnits' -Default 0
    $copilotUnused = Get-TenantReviewPropertySum -Items $items -Property 'unusedUnits' -Default 0
    $result = [ordered]@{
        dataset       = 'CopilotInventory'
        generatedAt   = (Get-Date).ToString('o')
        summary       = [pscustomobject]@{
            copilotSkus                   = $items.Count
            copilotPurchased              = [int]$copilotPurchased
            copilotAssigned               = [int]$copilotAssigned
            copilotUnused                 = [int]$copilotUnused
            activeCopilotUsers            = if ($usageReport.Count -gt 0) { $activeCopilotUsers } else { $null }
            inactiveCopilotLicensedUsers  = @($licensedUsers | Where-Object { $_.isStale -eq $true }).Count
            usageReportRequested          = [bool]$IncludeUsageReport
            usageReportAvailable          = $usageReportAvailable
            usageReportStatus             = $usageReportStatus
            estimatedMonthlyCost          = if ($null -ne $estimatedMonthlyCost) { [decimal]::Round([decimal]$estimatedMonthlyCost, 2) } else { $null }
            estimatedUnusedMonthlyCost    = if ($null -ne $estimatedUnusedMonthlyCost) { [decimal]::Round([decimal]$estimatedUnusedMonthlyCost, 2) } else { $null }
        }
        items         = @($items)
        licensedUsers = @($licensedUsers)
        warnings      = @($warnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = [pscustomobject]@{
            usageReport = @($usageReport)
        }
    }

    [pscustomobject]$result
}
