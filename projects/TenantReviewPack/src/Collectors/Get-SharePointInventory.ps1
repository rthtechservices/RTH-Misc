function Get-SharePointInventory {
    [CmdletBinding()]
    param(
        [switch]$IncludeRaw,

        [int]$ReportPeriodDays = 90,

        [switch]$IncludeOneDrive,

        [ValidateSet('Auto', 'PnP', 'SPO', 'GraphReports')]
        [string]$SiteSource = 'Auto'
    )

    $warnings = @()
    $sites = @()
    $oneDriveItems = @()
    $rawReports = [ordered]@{}
    $sharePointReportAvailable = $false
    $sharePointReportStatus = 'NotRequested'
    $oneDriveReportAvailable = $false
    $oneDriveReportStatus = if ($IncludeOneDrive) { 'Requested' } else { 'NotRequested' }

    if ($SiteSource -ne 'GraphReports') {
        $pnpCommand = Get-Command -Name Get-PnPTenantSite -ErrorAction SilentlyContinue
        $spoCommand = Get-Command -Name Get-SPOSite -ErrorAction SilentlyContinue
        if ($SiteSource -in @('Auto', 'PnP') -and $pnpCommand) {
            try {
                $sites = @(Get-PnPTenantSite -Detailed -ErrorAction Stop)
            } catch {
                $warnings += "PnP tenant site collection failed. $($_.Exception.Message)"
            }
        } elseif ($SiteSource -eq 'PnP') {
            $warnings += 'PnP.PowerShell/Get-PnPTenantSite is unavailable.'
        } elseif ($SiteSource -in @('Auto', 'SPO') -and $spoCommand) {
            try {
                $sites = @(Get-SPOSite -Limit All -ErrorAction Stop)
            } catch {
                $warnings += "SharePoint Online site collection failed. $($_.Exception.Message)"
            }
        } elseif ($SiteSource -eq 'SPO') {
            $warnings += 'Microsoft.Online.SharePoint.PowerShell/Get-SPOSite is unavailable.'
        }
    }

    $usageReport = @()
    try {
        $usageReport = @(Invoke-TenantReviewGraphCsvReport -CommandName 'Get-MgReportSharePointSiteUsageDetail' -ReportPeriodDays $ReportPeriodDays)
        $rawReports['sharePointSiteUsageDetail'] = @($usageReport)
        $sharePointReportAvailable = $true
        $sharePointReportStatus = 'Collected'
    } catch {
        $sharePointReportStatus = "Unavailable: $($_.Exception.Message)"
        if ($SiteSource -eq 'GraphReports') {
            $warnings += "SharePoint Graph usage report unavailable. $($_.Exception.Message)"
        }
    }

    $usageByUrl = @{}
    foreach ($row in $usageReport) {
        $url = Get-TenantReviewProperty -InputObject $row -Name 'Site Url'
        if (-not $url) { $url = Get-TenantReviewProperty -InputObject $row -Name 'Site URL' }
        if ($url) {
            $usageByUrl[$url.ToString().ToLowerInvariant()] = $row
        }
    }

    $items = @()
    foreach ($site in $sites) {
        $url = Get-TenantReviewProperty -InputObject $site -Name 'Url'
        $usage = if ($url -and $usageByUrl.ContainsKey($url.ToString().ToLowerInvariant())) { $usageByUrl[$url.ToString().ToLowerInvariant()] } else { $null }
        $storageMB = Get-TenantReviewProperty -InputObject $site -Name 'StorageUsageCurrent'
        $storageQuotaMB = Get-TenantReviewProperty -InputObject $site -Name 'StorageQuota'
        $usageStorageBytes = Get-TenantReviewProperty -InputObject $usage -Name 'Storage Used (Byte)'
        $usageStorageQuotaBytes = Get-TenantReviewProperty -InputObject $usage -Name 'Storage Allocated (Byte)'
        $sharingCapability = Get-TenantReviewProperty -InputObject $site -Name 'SharingCapability'

        $items += [pscustomobject]@{
            title                   = Get-TenantReviewProperty -InputObject $site -Name 'Title'
            url                     = if ($url) { $url.ToString() } else { $null }
            template                = Get-TenantReviewProperty -InputObject $site -Name 'Template'
            storageUsageGB          = if ($null -ne $storageMB) { ConvertTo-TenantReviewGB -Value $storageMB -InputUnit 'MB' } else { ConvertTo-TenantReviewGB -Value $usageStorageBytes -InputUnit 'Bytes' }
            storageQuotaGB          = if ($null -ne $storageQuotaMB) { ConvertTo-TenantReviewGB -Value $storageQuotaMB -InputUnit 'MB' } else { ConvertTo-TenantReviewGB -Value $usageStorageQuotaBytes -InputUnit 'Bytes' }
            lastContentModifiedDate = if (Get-TenantReviewProperty -InputObject $site -Name 'LastContentModifiedDate') { Get-TenantReviewProperty -InputObject $site -Name 'LastContentModifiedDate' } else { Get-TenantReviewProperty -InputObject $usage -Name 'Last Activity Date' }
            sharingCapability       = $sharingCapability
            externalSharingEnabled  = ($sharingCapability -and $sharingCapability.ToString() -notin @('Disabled', 'ExistingExternalUserSharingOnly'))
            activeFileCount         = Get-TenantReviewProperty -InputObject $usage -Name 'Active File Count'
            fileCount               = Get-TenantReviewProperty -InputObject $usage -Name 'File Count'
        }
    }

    if ($items.Count -eq 0 -and $usageReport.Count -gt 0) {
        foreach ($row in $usageReport) {
            $url = Get-TenantReviewProperty -InputObject $row -Name 'Site Url'
            if (-not $url) { $url = Get-TenantReviewProperty -InputObject $row -Name 'Site URL' }
            $items += [pscustomobject]@{
                title                   = Get-TenantReviewProperty -InputObject $row -Name 'Owner Display Name'
                url                     = $url
                template                = $null
                storageUsageGB          = ConvertTo-TenantReviewGB -Value (Get-TenantReviewProperty -InputObject $row -Name 'Storage Used (Byte)') -InputUnit 'Bytes'
                storageQuotaGB          = ConvertTo-TenantReviewGB -Value (Get-TenantReviewProperty -InputObject $row -Name 'Storage Allocated (Byte)') -InputUnit 'Bytes'
                lastContentModifiedDate = Get-TenantReviewProperty -InputObject $row -Name 'Last Activity Date'
                sharingCapability       = $null
                externalSharingEnabled  = $null
                activeFileCount         = Get-TenantReviewProperty -InputObject $row -Name 'Active File Count'
                fileCount               = Get-TenantReviewProperty -InputObject $row -Name 'File Count'
            }
        }
    }

    if ($IncludeOneDrive) {
        try {
            $oneDriveReport = @(Invoke-TenantReviewGraphCsvReport -CommandName 'Get-MgReportOneDriveUsageAccountDetail' -ReportPeriodDays $ReportPeriodDays)
            $rawReports['oneDriveUsageAccountDetail'] = @($oneDriveReport)
            $oneDriveReportAvailable = $true
            $oneDriveReportStatus = 'Collected'
            foreach ($row in $oneDriveReport) {
                $oneDriveItems += [pscustomobject]@{
                    ownerPrincipalName = Get-TenantReviewProperty -InputObject $row -Name 'Owner Principal Name'
                    ownerDisplayName   = Get-TenantReviewProperty -InputObject $row -Name 'Owner Display Name'
                    siteUrl            = Get-TenantReviewProperty -InputObject $row -Name 'Site URL'
                    lastActivityDate   = Get-TenantReviewProperty -InputObject $row -Name 'Last Activity Date'
                    storageUsedGB      = ConvertTo-TenantReviewGB -Value (Get-TenantReviewProperty -InputObject $row -Name 'Storage Used (Byte)') -InputUnit 'Bytes'
                    activeFileCount    = Get-TenantReviewProperty -InputObject $row -Name 'Active File Count'
                    fileCount          = Get-TenantReviewProperty -InputObject $row -Name 'File Count'
                }
            }
        } catch {
            $oneDriveReportStatus = "Unavailable: $($_.Exception.Message)"
            $warnings += "OneDrive Graph usage report was requested but unavailable. $($_.Exception.Message)"
        }
    }

    $largestSite = $items | Sort-Object -Property storageUsageGB -Descending | Select-Object -First 1
    $totalStorage = Get-TenantReviewPropertySum -Items $items -Property 'storageUsageGB'
    $totalStorageQuota = Get-TenantReviewPropertySum -Items $items -Property 'storageQuotaGB'
    $oneDriveStorage = Get-TenantReviewPropertySum -Items $oneDriveItems -Property 'storageUsedGB'
    $oneDriveActiveFiles = Get-TenantReviewPropertySum -Items $oneDriveItems -Property 'activeFileCount'

    $result = [ordered]@{
        dataset       = 'SharePointInventory'
        generatedAt   = (Get-Date).ToString('o')
        summary       = [pscustomobject]@{
            totalSites                  = $items.Count
            totalStorageGB              = if ($null -ne $totalStorage) { [decimal]::Round([decimal]$totalStorage, 2) } else { $null }
            totalStorageQuotaGB         = if ($null -ne $totalStorageQuota) { [decimal]::Round([decimal]$totalStorageQuota, 2) } else { $null }
            largestSiteTitle            = Get-TenantReviewProperty -InputObject $largestSite -Name 'title'
            largestSiteUrl              = Get-TenantReviewProperty -InputObject $largestSite -Name 'url'
            largestSiteStorageGB        = Get-TenantReviewProperty -InputObject $largestSite -Name 'storageUsageGB'
            externalSharingEnabledSites = @($items | Where-Object { $_.externalSharingEnabled -eq $true }).Count
            oneDriveAccounts            = $oneDriveItems.Count
            oneDriveTotalStorageGB      = if ($null -ne $oneDriveStorage) { [decimal]::Round([decimal]$oneDriveStorage, 2) } else { $null }
            oneDriveActiveFiles         = if ($null -ne $oneDriveActiveFiles) { [int]$oneDriveActiveFiles } else { $null }
            reportPeriod                = "D$ReportPeriodDays"
            siteSource                  = $SiteSource
            sharePointReportAvailable   = $sharePointReportAvailable
            sharePointReportStatus      = $sharePointReportStatus
            oneDriveReportRequested     = [bool]$IncludeOneDrive
            oneDriveReportAvailable     = $oneDriveReportAvailable
            oneDriveReportStatus        = $oneDriveReportStatus
        }
        items         = @($items)
        oneDriveItems = @($oneDriveItems)
        warnings      = @($warnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = $rawReports
    }

    [pscustomobject]$result
}
