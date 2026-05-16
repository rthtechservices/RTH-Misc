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

    function Get-AiSettingValue {
        param([string[]]$Names)

        foreach ($name in $Names) {
            $value = Get-TenantReviewProperty -InputObject $aiSettings -Name $name
            if ($value) {
                return $value
            }
        }

        return $null
    }

    function Resolve-TenantReviewAiApiKey {
        if ($RuntimeApiKey) {
            return $RuntimeApiKey
        }

        $directKey = Get-AiSettingValue -Names @('apiKey', 'apiKeyValue', 'key')
        if ($directKey) {
            return $directKey
        }

        $apiKeyVariable = Get-TenantReviewProperty -InputObject $aiSettings -Name 'apiKeyEnvironmentVariable'
        if (-not $apiKeyVariable) {
            return $null
        }

        $environmentValue = [Environment]::GetEnvironmentVariable($apiKeyVariable)
        if ($environmentValue) {
            return $environmentValue
        }

        if ($apiKeyVariable.ToString().Length -ge 32) {
            return $apiKeyVariable
        }

        return $null
    }

    function Resolve-TenantReviewAiRequest {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Endpoint,

            [Parameter(Mandatory = $false)]
            [string]$Model
        )

        $apiVersion = Get-AiSettingValue -Names @('apiVersion', 'azureApiVersion')
        if (-not $apiVersion) {
            $apiVersion = '2025-01-01-preview'
        }

        $trimmedEndpoint = $Endpoint.TrimEnd('/')
        $requestKind = 'ChatCompletions'
        $requestUri = $Endpoint
        $isAzure = (
            $trimmedEndpoint -match '\.openai\.azure\.com' -or
            $trimmedEndpoint -match '\.cognitiveservices\.azure\.com' -or
            $trimmedEndpoint -match '\.services\.ai\.azure\.com'
        )
        $deployment = Get-AiSettingValue -Names @('deployment', 'deploymentName', 'azureDeployment')
        if (-not $deployment) {
            $deployment = $Model
        }

        if ($trimmedEndpoint -match '/responses(\?|$)') {
            $requestKind = 'Responses'
        } elseif ($trimmedEndpoint -match '/chat/completions(\?|$)') {
            $requestKind = 'ChatCompletions'
        } elseif ($isAzure -and $trimmedEndpoint -match '/openai/deployments/[^/]+$') {
            $requestUri = "$trimmedEndpoint/chat/completions?api-version=$apiVersion"
        } elseif ($isAzure -and $trimmedEndpoint -notmatch '/openai/') {
            if (-not $deployment) {
                throw 'AI endpoint is an Azure OpenAI resource URL, but ai.deployment or ai.model is missing.'
            }
            $deploymentName = [uri]::EscapeDataString($deployment)
            $requestUri = "$trimmedEndpoint/openai/deployments/$deploymentName/chat/completions?api-version=$apiVersion"
        } elseif ($trimmedEndpoint -match 'api\.openai\.com$') {
            $requestKind = 'Responses'
            $requestUri = "$trimmedEndpoint/v1/responses"
        } elseif ($trimmedEndpoint -match 'api\.openai\.com/v1$') {
            $requestKind = 'Responses'
            $requestUri = "$trimmedEndpoint/responses"
        }

        [pscustomobject]@{
            Uri     = $requestUri
            Kind    = $requestKind
            IsAzure = $isAzure
        }
    }

    function Get-TenantReviewAiResponseText {
        param([Parameter(Mandatory = $true)][object]$Response)

        $content = Get-TenantReviewProperty -InputObject $Response -Name 'output_text'
        if ($content) {
            return $content
        }

        $choices = @(Get-TenantReviewProperty -InputObject $Response -Name 'choices')
        if ($choices.Count -gt 0) {
            $firstChoice = $choices | Select-Object -First 1
            $message = Get-TenantReviewProperty -InputObject $firstChoice -Name 'message'
            $content = Get-TenantReviewProperty -InputObject $message -Name 'content'
            if ($content) {
                return $content
            }
        }

        foreach ($outputItem in @(Get-TenantReviewProperty -InputObject $Response -Name 'output')) {
            foreach ($contentItem in @(Get-TenantReviewProperty -InputObject $outputItem -Name 'content')) {
                $text = Get-TenantReviewProperty -InputObject $contentItem -Name 'text'
                if ($text) {
                    return $text
                }
            }
        }

        return $null
    }

    function ConvertFrom-TenantReviewAiJsonText {
        param([Parameter(Mandatory = $true)][string]$Text)

        $clean = $Text.Trim()
        if ($clean -match '^```(?:json)?\s*(?<json>[\s\S]*?)\s*```$') {
            $clean = $Matches['json'].Trim()
        }

        return $clean | ConvertFrom-Json -ErrorAction Stop
    }

    $aiSettings = Get-TenantReviewProperty -InputObject $Settings -Name 'ai'
    $aiEnabled = Test-TenantReviewTruthy -Value (Get-TenantReviewProperty -InputObject $aiSettings -Name 'enabled')
    if (-not $aiEnabled) {
        return New-LocalNarrative -Datasets $Datasets
    }

    $endpoint = Get-TenantReviewProperty -InputObject $aiSettings -Name 'endpoint'
    $apiKey = Resolve-TenantReviewAiApiKey
    if (-not $endpoint -or -not $apiKey) {
        throw 'AI narrative is enabled, but the endpoint or API key is missing. Provide a valid ai.endpoint and an API key in Settings.json or the configured environment variable.'
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
Return strict JSON only as one top-level object with this exact shape:
{"sections":[{"dataset":"DatasetName","headline":"Short headline","plainEnglish":"One or two sentences","businessImpact":"One sentence","recommendedAction":"One sentence","confidence":"High|Medium|Low"}]}
Create concise Microsoft 365 tenant review narrative sections.
Each section must include dataset, headline, plainEnglish, businessImpact, recommendedAction, and confidence.
Use only this supplied data:
$($datasetSummaries | ConvertTo-Json -Depth 8)
"@

    try {
        $model = Get-TenantReviewProperty -InputObject $aiSettings -Name 'model'
        $request = Resolve-TenantReviewAiRequest -Endpoint $endpoint -Model $model
        if ($request.Kind -eq 'Responses') {
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
            }
            if ($model) {
                $bodyObject.model = $model
            }
            $temperature = Get-AiSettingValue -Names @('temperature')
            if ($null -ne $temperature) {
                $bodyObject.temperature = $temperature
            }
            $useJsonResponseFormat = Get-AiSettingValue -Names @('responseFormatJson')
            if ($null -eq $useJsonResponseFormat -or (Test-TenantReviewTruthy -Value $useJsonResponseFormat)) {
                $bodyObject.response_format = @{ type = 'json_object' }
            }
        }

        $body = $bodyObject | ConvertTo-Json -Depth 12
        $headers = if ($request.IsAzure) {
            @{ 'api-key' = $apiKey; 'Content-Type' = 'application/json' }
        } else {
            @{ Authorization = "Bearer $apiKey"; 'Content-Type' = 'application/json' }
        }
        $response = Invoke-RestMethod -Method Post -Uri $request.Uri -Headers $headers -Body $body -ErrorAction Stop
        $content = Get-TenantReviewAiResponseText -Response $response
        if (-not $content) {
            throw 'AI response did not include message content.'
        }

        $parsed = ConvertFrom-TenantReviewAiJsonText -Text $content
        $sectionsValue = Get-TenantReviewProperty -InputObject $parsed -Name 'sections'
        if ($null -eq $sectionsValue) {
            $sectionsValue = Get-TenantReviewProperty -InputObject $parsed -Name 'narrativeSections'
        }
        if ($null -eq $sectionsValue) {
            $sectionsValue = Get-TenantReviewProperty -InputObject $parsed -Name 'items'
        }
        if ($null -eq $sectionsValue) {
            $narrativeObject = Get-TenantReviewProperty -InputObject $parsed -Name 'narrative'
            $sectionsValue = Get-TenantReviewProperty -InputObject $narrativeObject -Name 'sections'
        }

        $sections = @($sectionsValue | Where-Object { $null -ne $_ })
        if ($sections.Count -eq 0 -and $parsed -is [array]) {
            $sections = @($parsed | Where-Object { $null -ne $_ })
        }
        if ($sections.Count -eq 0 -and (Get-TenantReviewProperty -InputObject $parsed -Name 'dataset')) {
            $sections = @($parsed)
        }
        if ($sections.Count -eq 0) {
            throw 'AI response JSON did not include narrative sections.'
        }
        foreach ($section in $sections) {
            foreach ($requiredProperty in @('dataset', 'headline', 'plainEnglish', 'businessImpact', 'recommendedAction', 'confidence')) {
                if (-not (Get-TenantReviewProperty -InputObject $section -Name $requiredProperty)) {
                    throw "AI response section was missing required property '$requiredProperty'."
                }
            }
        }

        [pscustomobject]@{
            dataset     = 'Narrative'
            generatedAt = (Get-Date).ToString('o')
            source      = 'AI'
            sections    = @($sections)
            warnings    = @()
        }
    } catch {
        throw "AI narrative failed. The script will not use local fallback because AI was enabled. $($_.Exception.Message)"
    }
}
