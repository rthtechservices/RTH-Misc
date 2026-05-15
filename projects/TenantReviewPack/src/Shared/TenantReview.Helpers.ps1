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
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }

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
        if ($additional.Contains($Name)) {
            return $additional[$Name]
        }

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
