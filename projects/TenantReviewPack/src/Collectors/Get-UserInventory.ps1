function Get-UserInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$StaleAfterDays = 90,

        [switch]$IncludeGuests,

        [switch]$IncludeRaw
    )

    function Get-ReviewValue {
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

    function ConvertTo-ReviewDateTime {
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

    function Get-AssignedSkuIds {
        param([Parameter(Mandatory = $false)][object]$AssignedLicenses)

        $skuIds = @()
        foreach ($license in @($AssignedLicenses)) {
            if ($null -eq $license) {
                continue
            }

            $skuId = Get-ReviewValue -InputObject $license -Name 'SkuId'
            if ($skuId) {
                $skuIds += [string]$skuId
            }
        }

        return @($skuIds)
    }

    $warnings = @()
    $propertiesWithSignIn = @(
        'Id',
        'DisplayName',
        'UserPrincipalName',
        'Mail',
        'AccountEnabled',
        'UserType',
        'CreatedDateTime',
        'AssignedLicenses',
        'AssignedPlans',
        'SignInActivity',
        'Department',
        'JobTitle',
        'CompanyName'
    )
    $propertiesWithoutSignIn = @(
        'Id',
        'DisplayName',
        'UserPrincipalName',
        'Mail',
        'AccountEnabled',
        'UserType',
        'CreatedDateTime',
        'AssignedLicenses',
        'AssignedPlans',
        'Department',
        'JobTitle',
        'CompanyName'
    )

    $signInActivityAvailable = $true
    try {
        $users = @(Get-MgUser -All -Property $propertiesWithSignIn -ErrorAction Stop)
    } catch {
        $signInActivityAvailable = $false
        $warnings += "Sign-in activity was unavailable. Retried user collection without signInActivity. $($_.Exception.Message)"
        $users = @(Get-MgUser -All -Property $propertiesWithoutSignIn -ErrorAction Stop)
    }

    $items = @()
    $summaryTotalUsers = 0
    $summaryMemberUsers = 0
    $summaryGuestUsers = 0
    $summaryEnabledUsers = 0
    $summaryDisabledUsers = 0
    $summaryLicensedUsers = 0
    $summaryUnlicensedUsers = 0
    $summaryStaleUsers = 0
    $summaryLicensedStaleUsers = 0
    $summaryLicensedDisabledUsers = 0
    $summaryGuestUsersWithLicenses = 0
    $summaryUsersWithoutSignInData = 0
    $staleCutoff = (Get-Date).AddDays(-1 * $StaleAfterDays)

    foreach ($user in $users) {
        $summaryTotalUsers++

        $userType = Get-ReviewValue -InputObject $user -Name 'UserType'
        $accountEnabled = Get-ReviewValue -InputObject $user -Name 'AccountEnabled'
        $assignedLicenses = @(Get-ReviewValue -InputObject $user -Name 'AssignedLicenses')
        $assignedSkuIds = @(Get-AssignedSkuIds -AssignedLicenses $assignedLicenses)
        $licenseCount = $assignedSkuIds.Count
        $isLicensed = $licenseCount -gt 0
        $isGuest = ($userType -eq 'Guest')
        $isEnabled = ($accountEnabled -eq $true)
        $isDisabled = ($accountEnabled -eq $false)

        if ($isGuest) {
            $summaryGuestUsers++
        } else {
            $summaryMemberUsers++
        }
        if ($isEnabled) {
            $summaryEnabledUsers++
        }
        if ($isDisabled) {
            $summaryDisabledUsers++
        }
        if ($isLicensed) {
            $summaryLicensedUsers++
        } else {
            $summaryUnlicensedUsers++
        }
        if ($isGuest -and $isLicensed) {
            $summaryGuestUsersWithLicenses++
        }

        $signInActivity = Get-ReviewValue -InputObject $user -Name 'SignInActivity'
        $lastSuccessfulSignInDateTime = ConvertTo-ReviewDateTime -Value (Get-ReviewValue -InputObject $signInActivity -Name 'LastSuccessfulSignInDateTime')
        $lastSignInDateTime = ConvertTo-ReviewDateTime -Value (Get-ReviewValue -InputObject $signInActivity -Name 'LastSignInDateTime')
        $lastNonInteractiveSignInDateTime = ConvertTo-ReviewDateTime -Value (Get-ReviewValue -InputObject $signInActivity -Name 'LastNonInteractiveSignInDateTime')
        $hasAnySignInData = ($null -ne $lastSuccessfulSignInDateTime -or $null -ne $lastSignInDateTime -or $null -ne $lastNonInteractiveSignInDateTime)
        if (-not $signInActivityAvailable -or -not $hasAnySignInData) {
            $summaryUsersWithoutSignInData++
        }

        $daysSinceLastSuccessfulSignIn = $null
        $isStale = $false
        if ($null -ne $lastSuccessfulSignInDateTime) {
            $daysSinceLastSuccessfulSignIn = [int]([datetime]::UtcNow - $lastSuccessfulSignInDateTime.ToUniversalTime()).TotalDays
            $isStale = ($lastSuccessfulSignInDateTime -lt $staleCutoff)
        }

        $isLicensedAndDisabled = ($isLicensed -and $isDisabled)
        $isLicensedAndStale = ($isLicensed -and $isStale)

        if ($isStale) {
            $summaryStaleUsers++
        }
        if ($isLicensedAndStale) {
            $summaryLicensedStaleUsers++
        }
        if ($isLicensedAndDisabled) {
            $summaryLicensedDisabledUsers++
        }

        $detailItem = [pscustomobject]@{
            id                                  = Get-ReviewValue -InputObject $user -Name 'Id'
            displayName                         = Get-ReviewValue -InputObject $user -Name 'DisplayName'
            userPrincipalName                   = Get-ReviewValue -InputObject $user -Name 'UserPrincipalName'
            mail                                = Get-ReviewValue -InputObject $user -Name 'Mail'
            userType                            = $userType
            accountEnabled                      = $accountEnabled
            department                          = Get-ReviewValue -InputObject $user -Name 'Department'
            jobTitle                            = Get-ReviewValue -InputObject $user -Name 'JobTitle'
            companyName                         = Get-ReviewValue -InputObject $user -Name 'CompanyName'
            createdDateTime                     = Get-ReviewValue -InputObject $user -Name 'CreatedDateTime'
            licenseCount                        = $licenseCount
            assignedSkuIds                      = @($assignedSkuIds)
            isLicensed                          = $isLicensed
            isGuest                             = $isGuest
            isEnabled                           = $isEnabled
            isDisabled                          = $isDisabled
            lastSuccessfulSignInDateTime        = if ($null -ne $lastSuccessfulSignInDateTime) { $lastSuccessfulSignInDateTime.ToString('o') } else { $null }
            lastSignInDateTime                  = if ($null -ne $lastSignInDateTime) { $lastSignInDateTime.ToString('o') } else { $null }
            lastNonInteractiveSignInDateTime    = if ($null -ne $lastNonInteractiveSignInDateTime) { $lastNonInteractiveSignInDateTime.ToString('o') } else { $null }
            daysSinceLastSuccessfulSignIn       = $daysSinceLastSuccessfulSignIn
            isStale                             = $isStale
            isLicensedAndDisabled               = $isLicensedAndDisabled
            isLicensedAndStale                  = $isLicensedAndStale
        }

        if (-not $isGuest -or $IncludeGuests -or $isLicensed) {
            $items += $detailItem
        }
    }

    $result = [ordered]@{
        dataset        = 'UserInventory'
        generatedAt    = (Get-Date).ToString('o')
        staleAfterDays = $StaleAfterDays
        summary        = [pscustomobject]@{
            totalUsers                 = $summaryTotalUsers
            memberUsers                = $summaryMemberUsers
            guestUsers                 = $summaryGuestUsers
            enabledUsers               = $summaryEnabledUsers
            disabledUsers              = $summaryDisabledUsers
            licensedUsers              = $summaryLicensedUsers
            unlicensedUsers            = $summaryUnlicensedUsers
            staleUsers                 = $summaryStaleUsers
            licensedStaleUsers         = $summaryLicensedStaleUsers
            licensedDisabledUsers      = $summaryLicensedDisabledUsers
            guestUsersWithLicenses     = $summaryGuestUsersWithLicenses
            usersWithoutSignInData     = $summaryUsersWithoutSignInData
        }
        items          = @($items)
        warnings       = @($warnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = @($users)
    }

    [pscustomobject]$result
}
