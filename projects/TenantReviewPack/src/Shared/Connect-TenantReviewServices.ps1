function Connect-TenantReviewServices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Settings,

        [Parameter(Mandatory = $false)]
        [object]$ConnectConfig
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

    function Test-ReviewTruthy {
        param([Parameter(Mandatory = $false)][object]$Value)

        if ($null -eq $Value) {
            return $false
        }

        if ($Value -is [bool]) {
            return [bool]$Value
        }

        return ($Value.ToString() -in @('true', '1', 'yes', 'enabled'))
    }

    function Assert-ReviewModule {
        param([Parameter(Mandatory = $true)][string]$Name)

        if (-not (Get-Module -ListAvailable -Name $Name)) {
            throw "Required PowerShell module '$Name' is not installed. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser"
        }

        Import-Module $Name -ErrorAction Stop
    }

    function Find-ReviewCertificate {
        param([Parameter(Mandatory = $true)][string]$Thumbprint)

        $normalizedThumbprint = ($Thumbprint -replace '\s', '').ToUpperInvariant()
        foreach ($storePath in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
            try {
                $certificate = Get-Item -Path (Join-Path $storePath $normalizedThumbprint) -ErrorAction SilentlyContinue
                if ($null -ne $certificate) {
                    return $certificate
                }
            } catch {
                continue
            }
        }

        return $null
    }

    function Get-ExchangeOnlineSettings {
        param(
            [Parameter(Mandatory = $false)][object]$Settings,
            [Parameter(Mandatory = $false)][object]$AuthConfig
        )

        $exchangeSettings = Get-ReviewProperty -InputObject $Settings -Name 'exchangeOnline'
        $enabled = Get-ReviewProperty -InputObject $exchangeSettings -Name 'enabled'
        $required = Get-ReviewProperty -InputObject $exchangeSettings -Name 'required'
        $organization = Get-ReviewProperty -InputObject $AuthConfig -Name 'exchangeOrganization'

        return [pscustomobject]@{
            Enabled = Test-ReviewTruthy -Value $enabled
            Required = Test-ReviewTruthy -Value $required
            Organization = $organization
        }
    }

    $warnings = @()
    $requiredGraphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Reports'
    )

    foreach ($moduleName in $requiredGraphModules) {
        Assert-ReviewModule -Name $moduleName
    }

    $authConfig = Get-ReviewProperty -InputObject $ConnectConfig -Name 'auth'
    if ($null -eq $authConfig) {
        throw 'Connection configuration is missing the auth object. Provide ConnectConfig.json with auth.mode set to AppCertificate or Interactive.'
    }

    $mode = Get-ReviewProperty -InputObject $authConfig -Name 'mode'
    if (-not $mode) {
        $mode = 'Interactive'
    }

    $tenantId = Get-ReviewProperty -InputObject $authConfig -Name 'tenantId'
    $clientId = Get-ReviewProperty -InputObject $authConfig -Name 'clientId'
    $thumbprint = Get-ReviewProperty -InputObject $authConfig -Name 'certificateThumbprint'
    $certificateFound = $false
    $exchangeOnlineConnected = $false

    Write-Verbose ("Connection config loaded. Mode: {0}; tenant id present: {1}; client id present: {2}; certificate thumbprint present: {3}" -f $mode, [bool]$tenantId, [bool]$clientId, [bool]$thumbprint)

    switch ($mode) {
        'AppCertificate' {
            $missingFields = @()
            if (-not $tenantId) { $missingFields += 'auth.tenantId' }
            if (-not $clientId) { $missingFields += 'auth.clientId' }
            if (-not $thumbprint) { $missingFields += 'auth.certificateThumbprint' }

            if ($missingFields.Count -gt 0) {
                throw ("AppCertificate mode requires these ConnectConfig.json values: {0}" -f ($missingFields -join ', '))
            }

            $certificate = Find-ReviewCertificate -Thumbprint $thumbprint
            if ($null -eq $certificate) {
                throw 'Configured certificate thumbprint was not found in Cert:\CurrentUser\My or Cert:\LocalMachine\My. Import the public/private certificate into one of those stores and retry.'
            }
            $certificateFound = $true

            Connect-MgGraph -TenantId $tenantId -ClientId $clientId -CertificateThumbprint $thumbprint -NoWelcome -ErrorAction Stop | Out-Null
        }

        'Interactive' {
            $scopes = @(
                'User.Read.All',
                'Directory.Read.All',
                'Organization.Read.All',
                'LicenseAssignment.Read.All',
                'Reports.Read.All',
                'AuditLog.Read.All'
            )

            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop | Out-Null
        }

        default {
            throw "Unsupported auth.mode '$mode'. Supported values are AppCertificate and Interactive."
        }
    }

    $exchangeSettings = Get-ExchangeOnlineSettings -Settings $Settings -AuthConfig $authConfig
    if ($exchangeSettings.Enabled) {
        if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
            $message = 'ExchangeOnlineManagement is not installed. Exchange Online connection was skipped. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
            if ($exchangeSettings.Required) {
                throw $message
            }
            $warnings += $message
        } elseif (-not $exchangeSettings.Organization -and $mode -eq 'AppCertificate') {
            $message = 'Exchange Online connection was skipped because auth.exchangeOrganization is not configured.'
            if ($exchangeSettings.Required) {
                throw $message
            }
            $warnings += $message
        } else {
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                if ($mode -eq 'AppCertificate') {
                    Connect-ExchangeOnline -AppId $clientId -CertificateThumbprint $thumbprint -Organization $exchangeSettings.Organization -ShowBanner:$false -ErrorAction Stop | Out-Null
                } else {
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
                }
                $exchangeOnlineConnected = $true
            } catch {
                $message = "Exchange Online connection was skipped after an error: $($_.Exception.Message)"
                if ($exchangeSettings.Required) {
                    throw $message
                }
                $warnings += $message
            }
        }
    }

    [pscustomobject]@{
        GraphConnected          = $true
        GraphAuthMode           = $mode
        GraphTenantIdLoaded     = [bool]$tenantId
        GraphClientIdLoaded     = [bool]$clientId
        CertificateFound        = $certificateFound
        ExchangeOnlineConnected = $exchangeOnlineConnected
        Warnings                = @($warnings)
    }
}
