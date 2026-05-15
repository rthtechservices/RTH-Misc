function Get-TenantOverview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [string]$ReviewPeriod
    )

    [pscustomobject]@{
        dataset       = 'TenantOverview'
        tenantName    = $TenantName
        reviewPeriod  = $ReviewPeriod
        generatedAt   = (Get-Date).ToString('o')
        summary       = [pscustomobject]@{
            domainsConfigured = $null
            verifiedDomains   = $null
            defaultDomain     = $null
            tenantId          = $null
            organizationName  = $null
        }
        raw           = @()
        implementationStatus = 'Stub - connect to Microsoft Graph organization/domains endpoints.'
    }
}
