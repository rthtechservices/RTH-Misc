function Get-DeviceInventory {
    [CmdletBinding()]
    param(
        [int]$StaleAfterDays = 90,

        [switch]$IncludeRaw,

        [switch]$IncludeIntune
    )

    $warnings = @()
    $devices = @()
    $managedDevices = @()
    $managedByAzureDeviceId = @{}
    $devicesWithoutLastSignIn = 0
    $intuneCollectionAvailable = $false
    $intuneCollectionStatus = if ($IncludeIntune) { 'Requested' } else { 'NotRequested' }
    $staleCutoff = (Get-Date).AddDays(-1 * $StaleAfterDays)

    if (Get-Command -Name Invoke-TenantReviewGraphRestRequest -ErrorAction SilentlyContinue) {
        try {
            $devices = @(Invoke-TenantReviewGraphRestRequest -Uri 'devices?$select=id,deviceId,displayName,accountEnabled,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime' -All)
        } catch {
            $warnings += "Unable to collect Entra devices from Microsoft Graph REST. $($_.Exception.Message)"
        }
    } elseif (Get-Command -Name Get-MgDevice -ErrorAction SilentlyContinue) {
        try {
            $devices = @(Get-MgDevice -All -Property @('Id', 'DeviceId', 'DisplayName', 'AccountEnabled', 'OperatingSystem', 'OperatingSystemVersion', 'TrustType', 'ApproximateLastSignInDateTime') -ErrorAction Stop)
        } catch {
            $warnings += "Unable to collect Entra devices from Microsoft Graph. $($_.Exception.Message)"
        }
    } else {
        $warnings += 'Get-MgDevice is not available. Install/import Microsoft.Graph.Identity.DirectoryManagement for device inventory.'
    }

    if ($IncludeIntune) {
        if (Get-Command -Name Invoke-TenantReviewGraphRestRequest -ErrorAction SilentlyContinue) {
            try {
                $managedDevices = @(Invoke-TenantReviewGraphRestRequest -Uri 'deviceManagement/managedDevices' -All)
                $intuneCollectionAvailable = $true
                $intuneCollectionStatus = 'Collected'
                foreach ($managedDevice in $managedDevices) {
                    $azureDeviceId = Get-TenantReviewProperty -InputObject $managedDevice -Name 'AzureAdDeviceId'
                    if (-not $azureDeviceId) { $azureDeviceId = Get-TenantReviewProperty -InputObject $managedDevice -Name 'AzureADDeviceId' }
                    if (-not $azureDeviceId) { $azureDeviceId = Get-TenantReviewProperty -InputObject $managedDevice -Name 'azureADDeviceId' }
                    if ($azureDeviceId) {
                        $managedByAzureDeviceId[$azureDeviceId.ToString()] = $managedDevice
                    }
                }
            } catch {
                $intuneCollectionStatus = "Unavailable: $($_.Exception.Message)"
                $warnings += "Intune managed device collection was requested but unavailable via Graph REST. $($_.Exception.Message)"
            }
        } elseif (Get-Command -Name Get-MgDeviceManagementManagedDevice -ErrorAction SilentlyContinue) {
            try {
                $managedDevices = @(Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop)
                $intuneCollectionAvailable = $true
                $intuneCollectionStatus = 'Collected'
                foreach ($managedDevice in $managedDevices) {
                    $azureDeviceId = Get-TenantReviewProperty -InputObject $managedDevice -Name 'AzureAdDeviceId'
                    if (-not $azureDeviceId) { $azureDeviceId = Get-TenantReviewProperty -InputObject $managedDevice -Name 'AzureADDeviceId' }
                    if ($azureDeviceId) {
                        $managedByAzureDeviceId[$azureDeviceId.ToString()] = $managedDevice
                    }
                }
            } catch {
                $intuneCollectionStatus = "Unavailable: $($_.Exception.Message)"
                $warnings += "Intune managed device collection was requested but unavailable. $($_.Exception.Message)"
            }
        } else {
            $intuneCollectionStatus = 'Unavailable: Get-MgDeviceManagementManagedDevice is not available.'
            $warnings += 'Intune managed device collection was requested but Get-MgDeviceManagementManagedDevice is not available.'
        }
    }

    $items = @()
    foreach ($device in $devices) {
        $deviceId = Get-TenantReviewProperty -InputObject $device -Name 'DeviceId'
        $managed = if ($deviceId -and $managedByAzureDeviceId.ContainsKey($deviceId.ToString())) { $managedByAzureDeviceId[$deviceId.ToString()] } else { $null }
        $lastSignIn = ConvertTo-TenantReviewDateTime -Value (Get-TenantReviewProperty -InputObject $device -Name 'ApproximateLastSignInDateTime')
        $daysSinceLastSignIn = $null
        $isStale = $false
        if ($null -ne $lastSignIn) {
            $daysSinceLastSignIn = [int]([datetime]::UtcNow - $lastSignIn.ToUniversalTime()).TotalDays
            $isStale = $lastSignIn -lt $staleCutoff
        } else {
            $devicesWithoutLastSignIn++
        }

        $items += [pscustomobject]@{
            deviceId                      = $deviceId
            displayName                   = Get-TenantReviewProperty -InputObject $device -Name 'DisplayName'
            accountEnabled                = Get-TenantReviewProperty -InputObject $device -Name 'AccountEnabled'
            operatingSystem               = Get-TenantReviewProperty -InputObject $device -Name 'OperatingSystem'
            operatingSystemVersion        = Get-TenantReviewProperty -InputObject $device -Name 'OperatingSystemVersion'
            trustType                     = Get-TenantReviewProperty -InputObject $device -Name 'TrustType'
            approximateLastSignInDateTime = if ($lastSignIn) { $lastSignIn.ToString('o') } else { $null }
            daysSinceLastSignIn           = $daysSinceLastSignIn
            isStale                       = $isStale
            isManaged                     = [bool]$managed
            complianceState               = Get-TenantReviewProperty -InputObject $managed -Name 'ComplianceState'
            userPrincipalName             = Get-TenantReviewProperty -InputObject $managed -Name 'UserPrincipalName'
        }
    }

    $uniqueWarnings = @($warnings | Sort-Object -Unique)
    $result = [ordered]@{
        dataset        = 'DeviceInventory'
        generatedAt    = (Get-Date).ToString('o')
        staleAfterDays = $StaleAfterDays
        summary        = [pscustomobject]@{
            totalDevices                    = $items.Count
            enabledDevices                  = @($items | Where-Object { $_.accountEnabled -eq $true }).Count
            disabledDevices                 = @($items | Where-Object { $_.accountEnabled -eq $false }).Count
            staleDevices                    = @($items | Where-Object { $_.isStale -eq $true }).Count
            windowsDevices                  = @($items | Where-Object { $_.operatingSystem -match 'Windows' }).Count
            macDevices                      = @($items | Where-Object { $_.operatingSystem -match 'Mac|macOS' }).Count
            iosDevices                      = @($items | Where-Object { $_.operatingSystem -match 'iOS|iPad' }).Count
            androidDevices                  = @($items | Where-Object { $_.operatingSystem -match 'Android' }).Count
            intuneManagedDevices            = @($items | Where-Object { $_.isManaged }).Count
            intuneCollectionRequested       = [bool]$IncludeIntune
            intuneCollectionAvailable       = $intuneCollectionAvailable
            intuneCollectionStatus          = $intuneCollectionStatus
            compliantDevices                = @($items | Where-Object { $_.complianceState -eq 'compliant' }).Count
            nonCompliantDevices             = @($items | Where-Object { $_.complianceState -eq 'noncompliant' }).Count
            unknownComplianceDevices        = @($items | Where-Object { -not $_.complianceState -or $_.complianceState -eq 'unknown' }).Count
            devicesWithoutLastSignIn        = $devicesWithoutLastSignIn
        }
        items          = @($items)
        warnings       = @($uniqueWarnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = [pscustomobject]@{
            devices        = @($devices)
            managedDevices = @($managedDevices)
        }
    }

    [pscustomobject]$result
}
