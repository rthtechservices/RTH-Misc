function Get-TenantReviewProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            if ($key.ToString().Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $InputObject[$key]
            }
        }
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }

    $additionalProperties = $InputObject.PSObject.Properties['AdditionalProperties']
    if ($null -ne $additionalProperties -and $additionalProperties.Value -is [System.Collections.IDictionary]) {
        $additional = $additionalProperties.Value
        foreach ($key in $additional.Keys) {
            if ($key.ToString().Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $additional[$key]
            }
        }
    }

    return $null
}

function Test-TenantReviewTruthy {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object]$Value)

    if ($null -eq $Value) {
        return $false
    }

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    return ($Value.ToString().ToLowerInvariant() -in @('true', '1', 'yes', 'enabled'))
}

function ConvertTo-TenantReviewDateTime {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return [datetime]$Value
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value.ToString(), [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function ConvertTo-TenantReviewDecimal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value,

        [decimal]$Default = 0
    )

    if ($null -eq $Value) {
        return $Default
    }

    $text = $Value.ToString() -replace '[^\d\.\-]', ''
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $Default
    }

    $parsed = [decimal]0
    if ([decimal]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function ConvertTo-TenantReviewGB {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value,

        [ValidateSet('Bytes', 'MB', 'GB', 'Auto')]
        [string]$InputUnit = 'Auto'
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = $Value.ToString()
    $number = ConvertTo-TenantReviewDecimal -Value $text -Default 0

    if ($InputUnit -eq 'GB') {
        return [decimal]::Round($number, 2)
    }
    if ($InputUnit -eq 'MB') {
        return [decimal]::Round(($number / 1024), 2)
    }
    if ($InputUnit -eq 'Bytes') {
        return [decimal]::Round(($number / 1GB), 2)
    }

    if ($text -match 'TB') {
        return [decimal]::Round(($number * 1024), 2)
    }
    if ($text -match 'GB') {
        return [decimal]::Round($number, 2)
    }
    if ($text -match 'MB') {
        return [decimal]::Round(($number / 1024), 2)
    }
    if ($text -match 'bytes') {
        $bytes = [regex]::Match($text, '\(([\d,]+)\s+bytes\)')
        if ($bytes.Success) {
            return [decimal]::Round((([decimal]($bytes.Groups[1].Value -replace ',', '')) / 1GB), 2)
        }
    }

    return [decimal]::Round($number, 2)
}

function Get-TenantReviewSafeCount {
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][object]$Value)

    if ($null -eq $Value) {
        return 0
    }

    return @($Value).Count
}

function Get-TenantReviewPropertySum {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Items,

        [Parameter(Mandatory = $true)]
        [string]$Property,

        [Parameter(Mandatory = $false)]
        [object]$Default = $null
    )

    $sum = [decimal]0
    $hasValue = $false

    foreach ($item in @($Items)) {
        $value = Get-TenantReviewProperty -InputObject $item -Name $Property
        if ($null -eq $value) {
            continue
        }

        $sum += ConvertTo-TenantReviewDecimal -Value $value -Default 0
        $hasValue = $true
    }

    if (-not $hasValue) {
        return $Default
    }

    return $sum
}

function ConvertFrom-TenantReviewGraphReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return @()
    }

    if ($InputObject -is [string]) {
        if ([string]::IsNullOrWhiteSpace($InputObject)) {
            return @()
        }

        return @($InputObject | ConvertFrom-Csv)
    }

    if ($InputObject -is [System.IO.FileInfo]) {
        return @(Get-Content -Path $InputObject.FullName -Raw | ConvertFrom-Csv)
    }

    return @($InputObject)
}

function Invoke-TenantReviewGraphCsvReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandName,

        [Parameter(Mandatory = $false)]
        [int]$ReportPeriodDays = 90
    )

    $reportMap = @{
        'Get-MgReportSharePointSiteUsageDetail' = "reports/getSharePointSiteUsageDetail(period='D$ReportPeriodDays')"
        'Get-MgReportOneDriveUsageAccountDetail' = "reports/getOneDriveUsageAccountDetail(period='D$ReportPeriodDays')"
        'Get-MgReportTeamActivityDetail' = "reports/getTeamsTeamActivityDetail(period='D$ReportPeriodDays')"
        'Get-MgReportTeamsTeamActivityDetail' = "reports/getTeamsTeamActivityDetail(period='D$ReportPeriodDays')"
    }
    if ($script:TenantReviewGraphAccessToken -and $reportMap.ContainsKey($CommandName)) {
        return @(Invoke-TenantReviewGraphReportRestCsv -ReportPath $reportMap[$CommandName])
    }

    $command = Get-Command -Name $CommandName -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "Graph report command '$CommandName' is not available."
    }

    $period = "D$ReportPeriodDays"
    $parameters = @{}
    if ($command.Parameters.ContainsKey('Period')) {
        $parameters.Period = $period
    }

    $tempFile = $null
    try {
        if ($command.Parameters.ContainsKey('OutFile')) {
            $tempFile = New-TemporaryFile
            $parameters.OutFile = $tempFile.FullName
            & $command @parameters | Out-Null
            if ((Test-Path $tempFile.FullName) -and (Get-Item $tempFile.FullName).Length -gt 0) {
                return @(Get-Content -Path $tempFile.FullName -Raw | ConvertFrom-Csv)
            }
            return @()
        }

        $result = & $command @parameters
        return @(ConvertFrom-TenantReviewGraphReportCsv -InputObject $result)
    } finally {
        if ($null -ne $tempFile -and (Test-Path $tempFile.FullName)) {
            Remove-Item -Path $tempFile.FullName -Force
        }
    }
}

function ConvertTo-TenantReviewBase64Url {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Set-TenantReviewGraphAppToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $now = [DateTimeOffset]::UtcNow
    $header = @{
        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-TenantReviewBase64Url -Bytes $Certificate.GetCertHash()
    } | ConvertTo-Json -Compress
    $payload = @{
        aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        iss = $ClientId
        jti = [guid]::NewGuid().ToString()
        nbf = $now.AddMinutes(-1).ToUnixTimeSeconds()
        sub = $ClientId
    } | ConvertTo-Json -Compress

    $encodedHeader = ConvertTo-TenantReviewBase64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($header))
    $encodedPayload = ConvertTo-TenantReviewBase64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($payload))
    $unsignedToken = "$encodedHeader.$encodedPayload"
    $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if ($null -eq $privateKey) {
        throw 'Configured certificate does not expose an RSA private key for Graph token acquisition.'
    }

    $signature = $privateKey.SignData(
        [Text.Encoding]::UTF8.GetBytes($unsignedToken),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $clientAssertion = "$unsignedToken.$(ConvertTo-TenantReviewBase64Url -Bytes $signature)"

    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
        client_id = $ClientId
        scope = 'https://graph.microsoft.com/.default'
        grant_type = 'client_credentials'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion = $clientAssertion
    } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop

    $script:TenantReviewGraphAccessToken = $tokenResponse.access_token
    $script:TenantReviewGraphTokenExpiresUtc = [DateTime]::UtcNow.AddSeconds([int]$tokenResponse.expires_in - 120)
}

function Invoke-TenantReviewGraphRestRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [switch]$All,

        [switch]$Raw
    )

    if (-not $script:TenantReviewGraphAccessToken) {
        throw 'Graph REST token is not available. Connect-TenantReviewServices must run first.'
    }

    $requestUri = if ($Uri -match '^https?://') { $Uri } else { "https://graph.microsoft.com/v1.0/$($Uri.TrimStart('/'))" }
    $headers = @{
        Authorization = "Bearer $script:TenantReviewGraphAccessToken"
    }

    if ($Raw) {
        return Invoke-RestMethod -Method Get -Uri $requestUri -Headers $headers -ErrorAction Stop
    }

    $items = @()
    do {
        $response = Invoke-RestMethod -Method Get -Uri $requestUri -Headers $headers -ErrorAction Stop
        if ($null -ne $response.PSObject.Properties['value']) {
            $items += @($response.value)
            $next = Get-TenantReviewProperty -InputObject $response -Name '@odata.nextLink'
            $requestUri = $next
        } else {
            return $response
        }
    } while ($All -and $requestUri)

    return @($items)
}

function Invoke-TenantReviewGraphReportRestCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReportPath
    )

    if (-not $script:TenantReviewGraphAccessToken) {
        throw 'Graph REST token is not available. Connect-TenantReviewServices must run first.'
    }

    $uri = "https://graph.microsoft.com/v1.0/$($ReportPath.TrimStart('/'))"
    $response = Invoke-WebRequest -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $script:TenantReviewGraphAccessToken" } -MaximumRedirection 5 -ErrorAction Stop
    $responseContent = if ($response.Content -is [byte[]]) {
        [Text.Encoding]::UTF8.GetString([byte[]]$response.Content)
    } else {
        $response.Content
    }

    if ([string]::IsNullOrWhiteSpace($responseContent)) {
        return @()
    }

    $content = $responseContent.TrimStart([char]0xFEFF).TrimStart()
    if ($content -notmatch '^[A-Za-z0-9_ "\(\)\/-]+,') {
        throw "Graph report response was not CSV. Content-Type: $($response.Headers['Content-Type'])"
    }

    $rows = @($content | ConvertFrom-Csv)
    if ($rows.Count -gt 0) {
        $firstRow = $rows | Select-Object -First 1
        $propertyCount = @($firstRow.PSObject.Properties).Count
        if ($propertyCount -le 1) {
            throw "Graph report response did not contain a usable CSV header. Content-Type: $($response.Headers['Content-Type'])"
        }
    }

    return @($rows)
}

function Get-TenantReviewConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Settings,

        [Parameter(Mandatory = $true)]
        [string[]]$Path,

        [Parameter(Mandatory = $false)]
        [object]$Default
    )

    $current = $Settings
    foreach ($part in $Path) {
        $current = Get-TenantReviewProperty -InputObject $current -Name $part
        if ($null -eq $current) {
            return $Default
        }
    }

    return $current
}
