function Get-LicenseInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$PriceMapPath,

        [Parameter(Mandatory = $false)]
        [string]$DefaultCurrency = 'CAD',

        [switch]$IncludeRaw
    )

    function Get-ReviewProperty {
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

    function Get-PriceMapEntry {
        param(
            [Parameter(Mandatory = $false)]
            [object]$PriceMap,

            [Parameter(Mandatory = $true)]
            [string]$SkuPartNumber
        )

        if ($null -eq $PriceMap) {
            return $null
        }

        foreach ($property in $PriceMap.PSObject.Properties) {
            if ($property.Name.Equals($SkuPartNumber, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $property.Value
            }
        }

        return $null
    }

    function ConvertTo-ReviewInt {
        param([Parameter(Mandatory = $false)][object]$Value)

        if ($null -eq $Value) {
            return 0
        }

        return [int]$Value
    }

    $warnings = @()
    $priceMap = $null
    if ($PriceMapPath) {
        if (Test-Path $PriceMapPath) {
            try {
                $priceMap = Get-Content -Path $PriceMapPath -Raw | ConvertFrom-Json
            } catch {
                $warnings += "License price map could not be loaded from '$PriceMapPath'. Costs will be null. $($_.Exception.Message)"
            }
        } else {
            $warnings += "License price map was not found at '$PriceMapPath'. Costs will be null."
        }
    } else {
        $warnings += 'License price map path was not provided. Costs will be null.'
    }

    $subscribedSkuCommand = Get-Command -Name Get-MgSubscribedSku -ErrorAction Stop
    if ($subscribedSkuCommand.Parameters.ContainsKey('All')) {
        $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
    } else {
        $skus = @(Get-MgSubscribedSku -ErrorAction Stop)
    }

    $items = @()
    $skusMissingPrice = @()
    $summaryEstimatedMonthlyCost = [decimal]0
    $summaryEstimatedAnnualCost = [decimal]0
    $summaryEstimatedUnusedMonthlyCost = [decimal]0
    $summaryEstimatedUnusedAnnualCost = [decimal]0
    $hasAnyPrice = $false

    foreach ($sku in $skus) {
        $prepaidUnits = Get-ReviewProperty -InputObject $sku -Name 'PrepaidUnits'
        $purchasedUnits = ConvertTo-ReviewInt -Value (Get-ReviewProperty -InputObject $prepaidUnits -Name 'Enabled')
        $assignedUnits = ConvertTo-ReviewInt -Value (Get-ReviewProperty -InputObject $sku -Name 'ConsumedUnits')
        $suspendedUnits = ConvertTo-ReviewInt -Value (Get-ReviewProperty -InputObject $prepaidUnits -Name 'Suspended')
        $warningUnits = ConvertTo-ReviewInt -Value (Get-ReviewProperty -InputObject $prepaidUnits -Name 'Warning')
        $unusedUnits = $purchasedUnits - $assignedUnits
        if ($unusedUnits -lt 0) {
            $unusedUnits = 0
        }

        $skuPartNumber = [string](Get-ReviewProperty -InputObject $sku -Name 'SkuPartNumber')
        $priceEntry = Get-PriceMapEntry -PriceMap $priceMap -SkuPartNumber $skuPartNumber
        $displayName = $skuPartNumber
        $monthlyUnitCost = $null
        $currency = $DefaultCurrency
        $estimatedMonthlyCost = $null
        $estimatedAnnualCost = $null
        $estimatedUnusedMonthlyCost = $null
        $estimatedUnusedAnnualCost = $null

        if ($null -ne $priceEntry) {
            $mappedDisplayName = Get-ReviewProperty -InputObject $priceEntry -Name 'displayName'
            $mappedMonthlyCost = Get-ReviewProperty -InputObject $priceEntry -Name 'monthlyCost'
            $mappedCurrency = Get-ReviewProperty -InputObject $priceEntry -Name 'currency'

            if ($mappedDisplayName) {
                $displayName = $mappedDisplayName
            }
            if ($mappedCurrency) {
                $currency = $mappedCurrency
            }
            if ($null -ne $mappedMonthlyCost) {
                $monthlyUnitCost = [decimal]$mappedMonthlyCost
                $estimatedMonthlyCost = [decimal]::Round(($monthlyUnitCost * $purchasedUnits), 2)
                $estimatedAnnualCost = [decimal]::Round(($estimatedMonthlyCost * 12), 2)
                $estimatedUnusedMonthlyCost = [decimal]::Round(($monthlyUnitCost * $unusedUnits), 2)
                $estimatedUnusedAnnualCost = [decimal]::Round(($estimatedUnusedMonthlyCost * 12), 2)
                $summaryEstimatedMonthlyCost += $estimatedMonthlyCost
                $summaryEstimatedAnnualCost += $estimatedAnnualCost
                $summaryEstimatedUnusedMonthlyCost += $estimatedUnusedMonthlyCost
                $summaryEstimatedUnusedAnnualCost += $estimatedUnusedAnnualCost
                $hasAnyPrice = $true
            }
        } else {
            $skusMissingPrice += $skuPartNumber
        }

        $servicePlans = @()
        $rawServicePlans = @(Get-ReviewProperty -InputObject $sku -Name 'ServicePlans')
        foreach ($servicePlan in $rawServicePlans) {
            if ($null -eq $servicePlan) {
                continue
            }

            $servicePlans += [pscustomobject]@{
                servicePlanName   = Get-ReviewProperty -InputObject $servicePlan -Name 'ServicePlanName'
                provisioningStatus = Get-ReviewProperty -InputObject $servicePlan -Name 'ProvisioningStatus'
                appliesTo          = Get-ReviewProperty -InputObject $servicePlan -Name 'AppliesTo'
            }
        }

        $items += [pscustomobject]@{
            skuId                       = Get-ReviewProperty -InputObject $sku -Name 'SkuId'
            skuPartNumber              = $skuPartNumber
            displayName                 = $displayName
            capabilityStatus           = Get-ReviewProperty -InputObject $sku -Name 'CapabilityStatus'
            purchasedUnits             = $purchasedUnits
            assignedUnits              = $assignedUnits
            suspendedUnits             = $suspendedUnits
            warningUnits               = $warningUnits
            unusedUnits                = $unusedUnits
            monthlyUnitCost            = $monthlyUnitCost
            currency                   = $currency
            estimatedMonthlyCost       = $estimatedMonthlyCost
            estimatedAnnualCost        = $estimatedAnnualCost
            estimatedUnusedMonthlyCost = $estimatedUnusedMonthlyCost
            estimatedUnusedAnnualCost  = $estimatedUnusedAnnualCost
            servicePlanCount           = $servicePlans.Count
            servicePlans               = @($servicePlans)
        }
    }

    $totalPurchased = ($items | Measure-Object -Property purchasedUnits -Sum).Sum
    $totalAssigned = ($items | Measure-Object -Property assignedUnits -Sum).Sum
    $totalUnused = ($items | Measure-Object -Property unusedUnits -Sum).Sum
    if ($null -eq $totalPurchased) { $totalPurchased = 0 }
    if ($null -eq $totalAssigned) { $totalAssigned = 0 }
    if ($null -eq $totalUnused) { $totalUnused = 0 }

    $result = [ordered]@{
        dataset     = 'LicenseInventory'
        generatedAt = (Get-Date).ToString('o')
        summary     = [pscustomobject]@{
            totalSkus                    = $items.Count
            totalPurchased               = [int]$totalPurchased
            totalAssigned                = [int]$totalAssigned
            totalUnused                  = [int]$totalUnused
            estimatedMonthlyCost         = if ($hasAnyPrice) { [decimal]::Round($summaryEstimatedMonthlyCost, 2) } else { $null }
            estimatedAnnualCost          = if ($hasAnyPrice) { [decimal]::Round($summaryEstimatedAnnualCost, 2) } else { $null }
            estimatedUnusedMonthlyCost   = if ($hasAnyPrice) { [decimal]::Round($summaryEstimatedUnusedMonthlyCost, 2) } else { $null }
            estimatedUnusedAnnualCost    = if ($hasAnyPrice) { [decimal]::Round($summaryEstimatedUnusedAnnualCost, 2) } else { $null }
            currency                     = $DefaultCurrency
            skusMissingPrice             = @($skusMissingPrice | Sort-Object -Unique)
        }
        items       = @($items)
        warnings    = @($warnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = @($skus)
    }

    [pscustomobject]$result
}
