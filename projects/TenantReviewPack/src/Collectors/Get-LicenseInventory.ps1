function Get-LicenseInventory {
    [CmdletBinding()]
    param()

    [pscustomobject]@{
        dataset = 'LicenseInventory'
        summary = [pscustomobject]@{
            totalSkus = $null
            totalPurchased = $null
            totalAssigned = $null
            totalUnused = $null
            estimatedMonthlyCost = $null
        }
        items = @()
        implementationStatus = 'Stub - add Microsoft 365 license collection logic here.'
    }
}
