function Get-TenantOverview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [string]$ReviewPeriod,

        [switch]$IncludeRaw,

        [switch]$IncludeSensitive
    )

    $warnings = @()
    $organizations = @()
    $domains = @()

    if (Get-Command -Name Get-MgOrganization -ErrorAction SilentlyContinue) {
        try {
            $organizations = @(Get-MgOrganization -ErrorAction Stop)
        } catch {
            $warnings += "Unable to collect organization details from Microsoft Graph. $($_.Exception.Message)"
        }
    } else {
        $warnings += 'Get-MgOrganization is not available. Install Microsoft.Graph.Identity.DirectoryManagement for tenant overview details.'
    }

    if (Get-Command -Name Get-MgDomain -ErrorAction SilentlyContinue) {
        try {
            $domains = @(Get-MgDomain -All -ErrorAction Stop)
        } catch {
            try {
                $domains = @(Get-MgDomain -ErrorAction Stop)
            } catch {
                $warnings += "Unable to collect domain details from Microsoft Graph. $($_.Exception.Message)"
            }
        }
    } else {
        $warnings += 'Get-MgDomain is not available. Install Microsoft.Graph.Identity.DirectoryManagement for domain inventory.'
    }

    $organization = $organizations | Select-Object -First 1
    $defaultDomain = $domains | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'IsDefault') -eq $true } | Select-Object -First 1
    $initialDomain = $domains | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'IsInitial') -eq $true } | Select-Object -First 1
    $verifiedDomains = @($domains | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'IsVerified') -eq $true })
    $federatedDomains = @($domains | Where-Object { (Get-TenantReviewProperty -InputObject $_ -Name 'AuthenticationType') -eq 'Federated' })

    $domainItems = @()
    foreach ($domain in $domains) {
        $domainItems += [pscustomobject]@{
            id                 = Get-TenantReviewProperty -InputObject $domain -Name 'Id'
            isDefault          = Get-TenantReviewProperty -InputObject $domain -Name 'IsDefault'
            isInitial          = Get-TenantReviewProperty -InputObject $domain -Name 'IsInitial'
            isVerified         = Get-TenantReviewProperty -InputObject $domain -Name 'IsVerified'
            authenticationType = Get-TenantReviewProperty -InputObject $domain -Name 'AuthenticationType'
            supportedServices  = @(Get-TenantReviewProperty -InputObject $domain -Name 'SupportedServices')
        }
    }

    $tenantId = Get-TenantReviewProperty -InputObject $organization -Name 'Id'
    $result = [ordered]@{
        dataset      = 'TenantOverview'
        tenantName   = $TenantName
        reviewPeriod = $ReviewPeriod
        generatedAt  = (Get-Date).ToString('o')
        summary      = [pscustomobject]@{
            organizationName              = Get-TenantReviewProperty -InputObject $organization -Name 'DisplayName'
            tenantIdPresent               = [bool]$tenantId
            defaultDomain                 = Get-TenantReviewProperty -InputObject $defaultDomain -Name 'Id'
            initialDomain                 = Get-TenantReviewProperty -InputObject $initialDomain -Name 'Id'
            domainCount                   = $domains.Count
            verifiedDomainCount           = $verifiedDomains.Count
            federatedDomainCount          = $federatedDomains.Count
            passwordValidityPeriodInDays  = Get-TenantReviewProperty -InputObject $organization -Name 'PasswordValidityPeriodInDays'
            technicalNotificationMails    = @(Get-TenantReviewProperty -InputObject $organization -Name 'TechnicalNotificationMails')
            assignedPlansCount            = Get-TenantReviewSafeCount -Value (Get-TenantReviewProperty -InputObject $organization -Name 'AssignedPlans')
            createdDateTime               = Get-TenantReviewProperty -InputObject $organization -Name 'CreatedDateTime'
            generatedAt                   = (Get-Date).ToString('o')
        }
        items       = @($domainItems)
        warnings    = @($warnings)
    }

    if ($IncludeSensitive) {
        $result.summary | Add-Member -NotePropertyName tenantId -NotePropertyValue $tenantId -Force
    }

    if ($IncludeRaw) {
        $result['raw'] = [pscustomobject]@{
            organizations = @($organizations)
            domains       = @($domains)
        }
    }

    [pscustomobject]$result
}
