function Get-LicenseUserAnalysis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$LicenseInventory,

        [Parameter(Mandatory = $true)]
        [object]$UserInventory
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

    function ConvertTo-ReviewDecimal {
        param([Parameter(Mandatory = $false)][object]$Value)

        if ($null -eq $Value) {
            return [decimal]0
        }

        return [decimal]$Value
    }

    $licenseItems = @(Get-ReviewProperty -InputObject $LicenseInventory -Name 'items')
    $userItems = @(Get-ReviewProperty -InputObject $UserInventory -Name 'items')

    $unusedLicenseCount = 0
    $estimatedUnusedMonthlyCost = [decimal]0
    $estimatedUnusedAnnualCost = [decimal]0
    foreach ($license in $licenseItems) {
        $unusedLicenseCount += [int](Get-ReviewProperty -InputObject $license -Name 'unusedUnits')
        $estimatedUnusedMonthlyCost += ConvertTo-ReviewDecimal -Value (Get-ReviewProperty -InputObject $license -Name 'estimatedUnusedMonthlyCost')
        $estimatedUnusedAnnualCost += ConvertTo-ReviewDecimal -Value (Get-ReviewProperty -InputObject $license -Name 'estimatedUnusedAnnualCost')
    }

    $disabledLicensedUsers = @($userItems | Where-Object { (Get-ReviewProperty -InputObject $_ -Name 'isLicensedAndDisabled') -eq $true })
    $staleLicensedUsers = @($userItems | Where-Object { (Get-ReviewProperty -InputObject $_ -Name 'isLicensedAndStale') -eq $true })
    $guestLicensedUsers = @($userItems | Where-Object { (Get-ReviewProperty -InputObject $_ -Name 'isGuest') -eq $true -and (Get-ReviewProperty -InputObject $_ -Name 'isLicensed') -eq $true })

    $attentionItems = @()
    if ($unusedLicenseCount -gt 0) {
        $attentionItems += [pscustomobject]@{
            severity          = 'Medium'
            headline          = "$unusedLicenseCount purchased licenses appear unused"
            plainEnglish      = 'These licenses are available in the tenant but are not currently assigned to users.'
            businessImpact    = 'The tenant may be carrying license capacity that is not being used.'
            recommendedAction = 'Review upcoming hiring, project, and retention needs before reducing or reallocating licenses.'
        }
    }

    if ($disabledLicensedUsers.Count -gt 0) {
        $attentionItems += [pscustomobject]@{
            severity          = 'High'
            headline          = "$($disabledLicensedUsers.Count) disabled users still have licenses"
            plainEnglish      = 'These accounts are turned off but still have Microsoft 365 licenses assigned.'
            businessImpact    = 'The tenant may be paying for licenses that cannot currently be used.'
            recommendedAction = 'Review and remove licenses from disabled users unless there is a retention or compliance reason.'
        }
    }

    if ($staleLicensedUsers.Count -gt 0) {
        $attentionItems += [pscustomobject]@{
            severity          = 'Medium'
            headline          = "$($staleLicensedUsers.Count) licensed users appear stale"
            plainEnglish      = 'These licensed accounts have not had a recent successful sign-in based on the configured stale threshold.'
            businessImpact    = 'Some assigned licenses may not be producing active business value.'
            recommendedAction = 'Confirm whether these accounts are still needed and remove licenses where appropriate.'
        }
    }

    if ($guestLicensedUsers.Count -gt 0) {
        $attentionItems += [pscustomobject]@{
            severity          = 'Medium'
            headline          = "$($guestLicensedUsers.Count) guest users have licenses"
            plainEnglish      = 'Guest accounts are external users, and some have Microsoft 365 licenses assigned.'
            businessImpact    = 'External accounts with licenses can create avoidable cost and access review obligations.'
            recommendedAction = 'Validate that each licensed guest account still requires a tenant license.'
        }
    }

    [pscustomobject]@{
        dataset               = 'LicenseUserAnalysis'
        generatedAt           = (Get-Date).ToString('o')
        summary               = [pscustomobject]@{
            unusedLicenseCount          = $unusedLicenseCount
            estimatedUnusedMonthlyCost  = [decimal]::Round($estimatedUnusedMonthlyCost, 2)
            estimatedUnusedAnnualCost   = [decimal]::Round($estimatedUnusedAnnualCost, 2)
            disabledLicensedUserCount   = $disabledLicensedUsers.Count
            staleLicensedUserCount      = $staleLicensedUsers.Count
            guestLicensedUserCount      = $guestLicensedUsers.Count
            attentionItemCount          = $attentionItems.Count
        }
        attentionItems        = @($attentionItems)
        disabledLicensedUsers = @($disabledLicensedUsers)
        staleLicensedUsers    = @($staleLicensedUsers)
        guestLicensedUsers    = @($guestLicensedUsers)
    }
}
