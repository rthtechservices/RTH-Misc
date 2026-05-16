function Invoke-AINarrative {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Datasets,

        [Parameter(Mandatory = $false)]
        [object]$Settings,

        [Parameter(Mandatory = $false)]
        [string]$RuntimeApiKey
    )

    function New-LocalNarrativeSection {
        param(
            [string]$Dataset,
            [object]$Summary
        )

        switch ($Dataset) {
            'LicenseUserAnalysis' {
                $unused = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'unusedLicenseCount')
                $disabled = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'disabledLicensedUserCount')
                $stale = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'staleLicensedUserCount')
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$unused unused licenses, $disabled disabled licensed users, $stale stale licensed users"
                    plainEnglish      = 'License assignment should be reviewed for unused capacity and accounts that are disabled or no longer active.'
                    businessImpact    = 'Cleaning up licenses can reduce recurring spend and improve account governance.'
                    recommendedAction = 'Review disabled and stale licensed accounts before the next renewal or monthly true-up.'
                    confidence        = 'High'
                }
            }
            'MailboxInventory' {
                $external = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'mailboxesForwardingExternally')
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$external mailboxes show possible external forwarding"
                    plainEnglish      = 'Mailbox forwarding and transport rules can move mail outside the tenant.'
                    businessImpact    = 'Unexpected forwarding can increase data leakage and compliance risk.'
                    recommendedAction = 'Review forwarding-enabled mailboxes and transport rules for business justification.'
                    confidence        = 'Medium'
                }
            }
            'SharePoint' {
                $sites = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalSites')
                $external = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'externalSharingEnabledSites')
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$sites SharePoint sites reviewed; $external allow external sharing"
                    plainEnglish      = 'SharePoint storage and sharing posture were summarized from available site and usage data.'
                    businessImpact    = 'External sharing settings and large sites are useful review points for governance and storage planning.'
                    recommendedAction = 'Confirm externally shared sites still align with client and project needs.'
                    confidence        = 'Medium'
                }
            }
            'Teams' {
                $inactive = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'inactiveTeams')
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$inactive inactive Teams found in the reporting period"
                    plainEnglish      = 'Teams activity data highlights collaboration spaces that may no longer be active.'
                    businessImpact    = 'Inactive teams can add clutter and create unmanaged storage or access risk.'
                    recommendedAction = 'Archive or clean up inactive teams after owner confirmation.'
                    confidence        = 'Medium'
                }
            }
            'Devices' {
                $staleDevices = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'staleDevices')
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$staleDevices stale devices found"
                    plainEnglish      = 'Device inventory flags endpoints that have not signed in recently.'
                    businessImpact    = 'Stale devices can distort compliance reporting and may represent unmanaged access history.'
                    recommendedAction = 'Review stale and disabled devices for cleanup.'
                    confidence        = 'Medium'
                }
            }
            'Copilot' {
                $unused = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'copilotUnused')
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$unused Copilot licenses appear unused"
                    plainEnglish      = 'Copilot licensing is summarized separately because the SKU is high-value and usage reporting may vary by tenant.'
                    businessImpact    = 'Unused Copilot licenses can materially affect Microsoft 365 spend.'
                    recommendedAction = 'Compare assigned Copilot licenses with actual usage before renewal or expansion.'
                    confidence        = 'Medium'
                }
            }
            default {
                return [pscustomobject]@{
                    dataset           = $Dataset
                    headline          = "$Dataset data collected"
                    plainEnglish      = 'The dataset was collected where module and permission coverage allowed.'
                    businessImpact    = 'This gives the review package supporting evidence for tenant health and planning.'
                    recommendedAction = 'Review warnings and high-count findings before presenting the final package.'
                    confidence        = 'Medium'
                }
            }
        }
    }

    function New-LocalNarrative {
        param([object]$Datasets)

        $sections = @()
        foreach ($key in $Datasets.Keys) {
            $dataset = $Datasets[$key]
            $sections += New-LocalNarrativeSection -Dataset $key -Summary (Get-TenantReviewProperty -InputObject $dataset -Name 'summary')
        }

        [pscustomobject]@{
            dataset     = 'Narrative'
            generatedAt = (Get-Date).ToString('o')
            source      = 'LocalRuleBased'
            sections    = @($sections)
            warnings    = @()
        }
    }

    $aiSettings = Get-TenantReviewProperty -InputObject $Settings -Name 'ai'
    $aiEnabled = Test-TenantReviewTruthy -Value (Get-TenantReviewProperty -InputObject $aiSettings -Name 'enabled')
    if (-not $aiEnabled) {
        return New-LocalNarrative -Datasets $Datasets
    }

    $endpoint = Get-TenantReviewProperty -InputObject $aiSettings -Name 'endpoint'
    $apiKeyVariable = Get-TenantReviewProperty -InputObject $aiSettings -Name 'apiKeyEnvironmentVariable'
    $apiKey = if ($RuntimeApiKey) { $RuntimeApiKey } elseif ($apiKeyVariable) { [Environment]::GetEnvironmentVariable($apiKeyVariable) } else { $null }
    if (-not $endpoint -or -not $apiKey) {
        throw 'AI narrative is enabled, but the endpoint or API key is missing. Provide a valid ai.endpoint and an API key via the configured environment variable or interactive prompt.'
    }

    $datasetSummaries = @()
    foreach ($key in $Datasets.Keys) {
        $datasetSummaries += [pscustomobject]@{
            dataset  = $key
            summary  = Get-TenantReviewProperty -InputObject $Datasets[$key] -Name 'summary'
            warnings = Get-TenantReviewProperty -InputObject $Datasets[$key] -Name 'warnings'
        }
    }

    $prompt = @"
Return strict JSON only. Create concise Microsoft 365 tenant review narrative sections.
Each section must include dataset, headline, plainEnglish, businessImpact, recommendedAction, and confidence.
Use only this supplied data:
$($datasetSummaries | ConvertTo-Json -Depth 8)
"@

    try {
        $model = Get-TenantReviewProperty -InputObject $aiSettings -Name 'model'
        if ($endpoint -match '/responses') {
            $bodyObject = @{
                model = $model
                input = @(
                    @{ role = 'system'; content = 'You produce strict JSON only for client-ready Microsoft 365 review summaries.' },
                    @{ role = 'user'; content = $prompt }
                )
            }
        } else {
            $bodyObject = @{
                messages = @(
                    @{ role = 'system'; content = 'You produce strict JSON only for client-ready Microsoft 365 review summaries.' },
                    @{ role = 'user'; content = $prompt }
                )
                temperature = 0.2
            }
            if ($model) {
                $bodyObject.model = $model
            }
        }

        $body = $bodyObject | ConvertTo-Json -Depth 12
        $response = Invoke-RestMethod -Method Post -Uri $endpoint -Headers @{ 'api-key' = $apiKey; 'Content-Type' = 'application/json' } -Body $body -ErrorAction Stop
        $content = Get-TenantReviewProperty -InputObject $response -Name 'output_text'
        if (-not $content) {
            $content = Get-TenantReviewProperty -InputObject (($response.choices | Select-Object -First 1).message) -Name 'content'
        }
        if (-not $content) {
            $firstOutput = @($response.output) | Select-Object -First 1
            $firstContent = @(Get-TenantReviewProperty -InputObject $firstOutput -Name 'content') | Select-Object -First 1
            $content = Get-TenantReviewProperty -InputObject $firstContent -Name 'text'
        }
        if (-not $content) {
            throw 'AI response did not include message content.'
        }

        $parsed = $content | ConvertFrom-Json -ErrorAction Stop
        [pscustomobject]@{
            dataset     = 'Narrative'
            generatedAt = (Get-Date).ToString('o')
            source      = 'AI'
            sections    = @(Get-TenantReviewProperty -InputObject $parsed -Name 'sections')
            warnings    = @()
        }
    } catch {
        throw "AI narrative failed. The script will not use local fallback because AI was enabled. $($_.Exception.Message)"
    }
}
