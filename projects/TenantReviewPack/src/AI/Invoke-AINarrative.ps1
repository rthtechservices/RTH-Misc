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

    # ─── formatting helpers ────────────────────────────────────────────────────

    function Format-ReviewNumber {
        param([object]$Value)
        if ($null -eq $Value) { return 'N/A' }
        try { return '{0:N0}' -f [double]$Value } catch { return $Value.ToString() }
    }

    function Format-ReviewCurrency {
        param([object]$Value, [string]$Currency = 'CAD')
        if ($null -eq $Value) { return 'N/A' }
        try { return ('{0} ${1:N2}' -f $Currency, [double]$Value) } catch { return $Value.ToString() }
    }

    function Format-ReviewGB {
        param([object]$Value)
        if ($null -eq $Value) { return 'N/A' }
        try {
            $gb = [double]$Value
            if ($gb -ge 1024) { return ('{0:N1} TB' -f ($gb / 1024)) }
            return ('{0:N1} GB' -f $gb)
        } catch { return $Value.ToString() }
    }

    # ─── status scoring ────────────────────────────────────────────────────────

    function Get-MetricStatus {
        param([string]$Name, [object]$Value)
        $v = try { [int]$Value } catch { 0 }
        switch ($Name) {
            'unusedLicenses'        { if ($v -eq 0) { 'Good' } elseif ($v -le 5) { 'Watch' } else { 'ActionRecommended' } }
            'disabledLicensedUsers' { if ($v -eq 0) { 'Good' } else { 'ActionRecommended' } }
            'staleLicensedUsers'    { if ($v -eq 0) { 'Good' } elseif ($v -le 3) { 'Watch' } else { 'ActionRecommended' } }
            'guestUsers'            { if ($v -eq 0) { 'Good' } else { 'Watch' } }
            'externalForwarding'    { if ($v -eq 0) { 'Good' } else { 'ActionRecommended' } }
            'inboxForwardingRules'  { if ($v -eq 0) { 'Good' } elseif ($v -le 5) { 'Watch' } else { 'ActionRecommended' } }
            'fwdMailboxes'          { if ($v -eq 0) { 'Good' } else { 'Watch' } }
            'externalSharingSites'  { if ($v -eq 0) { 'Good' } else { 'Watch' } }
            'inactiveTeams'         { if ($v -eq 0) { 'Good' } elseif ($v -le 3) { 'Watch' } else { 'ActionRecommended' } }
            'teamsNoOwners'         { if ($v -eq 0) { 'Good' } else { 'ActionRecommended' } }
            'staleDevices'          { if ($v -eq 0) { 'Good' } elseif ($v -le 5) { 'Watch' } else { 'ActionRecommended' } }
            'copilotUnused'         { if ($v -eq 0) { 'Good' } elseif ($v -le 2) { 'Watch' } else { 'ActionRecommended' } }
            default                 { 'Informational' }
        }
    }

    # ─── KPI builder ──────────────────────────────────────────────────────────

    function New-KpiData {
        param([object]$Datasets)
        $kpis = [System.Collections.Generic.List[object]]::new()

        $userSum = Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'summary'
        if ($userSum) {
            $totalUsers  = Get-TenantReviewProperty -InputObject $userSum -Name 'totalUsers'
            $members     = Get-TenantReviewProperty -InputObject $userSum -Name 'memberUsers'
            $guests      = Get-TenantReviewProperty -InputObject $userSum -Name 'guestUsers'
            $stale       = Get-TenantReviewProperty -InputObject $userSum -Name 'staleUsers'
            $signInAvail = Get-TenantReviewProperty -InputObject $userSum -Name 'signInActivityAvailable'
            $kpis.Add([pscustomobject]@{
                area = 'Users'; label = 'Total Users'; value = Format-ReviewNumber $totalUsers; rawValue = [int]$totalUsers
                description = "All accounts in the tenant — $members member$(if([int]$members -ne 1){'s'}) and $guests guest$(if([int]$guests -ne 1){'s'})."
                status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'Users'; label = 'Guest Users'; value = Format-ReviewNumber $guests; rawValue = [int]$guests
                description = 'External users with access to tenant resources. Review periodically to confirm ongoing need.'
                status = Get-MetricStatus -Name 'guestUsers' -Value $guests
            })
            if ($signInAvail -eq $true) {
                $kpis.Add([pscustomobject]@{
                    area = 'Users'; label = 'Stale Users'; value = Format-ReviewNumber $stale; rawValue = [int]$stale
                    description = 'Accounts with no recent sign-in activity. May indicate dormant access that should be reviewed.'
                    status = Get-MetricStatus -Name 'staleLicensedUsers' -Value $stale
                })
            } else {
                $kpis.Add([pscustomobject]@{
                    area = 'Users'; label = 'Stale Users'; value = 'N/A'; rawValue = $null
                    description = 'Sign-in activity data was not collected this run. Enable sign-in reporting to detect dormant accounts.'
                    status = 'NotAvailable'
                })
            }
        }

        $licSum = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
        $anaSum = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
        if ($licSum) {
            $assigned      = Get-TenantReviewProperty -InputObject $licSum -Name 'totalAssigned'
            $unused        = Get-TenantReviewProperty -InputObject $licSum -Name 'totalUnused'
            $monthly       = Get-TenantReviewProperty -InputObject $licSum -Name 'estimatedMonthlyCost'
            $unusedMonthly = Get-TenantReviewProperty -InputObject $licSum -Name 'estimatedUnusedMonthlyCost'
            $currency      = Get-TenantReviewProperty -InputObject $licSum -Name 'currency'
            if (-not $currency) { $currency = 'CAD' }
            $kpis.Add([pscustomobject]@{
                area = 'Licensing'; label = 'Licensed Users'; value = Format-ReviewNumber $assigned; rawValue = [int]$assigned
                description = 'Microsoft 365 license assignments currently active in the tenant.'
                status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'Licensing'; label = 'Unused Licenses'; value = Format-ReviewNumber $unused; rawValue = [int]$unused
                description = 'Purchased licenses not currently assigned to any user.'
                status = Get-MetricStatus -Name 'unusedLicenses' -Value $unused
            })
            $kpis.Add([pscustomobject]@{
                area = 'Licensing'; label = 'Est. Monthly Spend'; value = Format-ReviewCurrency $monthly $currency; rawValue = [double]$monthly
                description = 'Estimated total monthly license cost based on available SKU price data.'
                status = 'Informational'
            })
            if ([double]$unusedMonthly -gt 0) {
                $kpis.Add([pscustomobject]@{
                    area = 'Licensing'; label = 'Unused Monthly Spend'; value = Format-ReviewCurrency $unusedMonthly $currency; rawValue = [double]$unusedMonthly
                    description = 'Estimated monthly cost of licenses that are purchased but not assigned to anyone.'
                    status = 'ActionRecommended'
                })
            }
        }
        if ($anaSum) {
            $disabledLic = Get-TenantReviewProperty -InputObject $anaSum -Name 'disabledLicensedUserCount'
            $staleLic    = Get-TenantReviewProperty -InputObject $anaSum -Name 'staleLicensedUserCount'
            $kpis.Add([pscustomobject]@{
                area = 'Licensing'; label = 'Disabled Licensed Users'; value = Format-ReviewNumber $disabledLic; rawValue = [int]$disabledLic
                description = 'Disabled accounts that still have Microsoft 365 licenses assigned.'
                status = Get-MetricStatus -Name 'disabledLicensedUsers' -Value $disabledLic
            })
            $kpis.Add([pscustomobject]@{
                area = 'Licensing'; label = 'Stale Licensed Users'; value = Format-ReviewNumber $staleLic; rawValue = [int]$staleLic
                description = 'Licensed accounts with no recent sign-in. May no longer need paid licenses.'
                status = Get-MetricStatus -Name 'staleLicensedUsers' -Value $staleLic
            })
        }

        $mbxSum = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
        if ($mbxSum) {
            $totalMbx  = Get-TenantReviewProperty -InputObject $mbxSum -Name 'totalMailboxes'
            $sharedMbx = Get-TenantReviewProperty -InputObject $mbxSum -Name 'sharedMailboxes'
            $fwdMbx    = Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesWithForwarding'
            $extFwd    = Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesForwardingExternally'
            $inboxFwd  = Get-TenantReviewProperty -InputObject $mbxSum -Name 'inboxForwardingRulesFound'
            $kpis.Add([pscustomobject]@{
                area = 'Exchange'; label = 'Total Mailboxes'; value = Format-ReviewNumber $totalMbx; rawValue = [int]$totalMbx
                description = "All mailboxes in Exchange Online — $sharedMbx shared mailbox$(if([int]$sharedMbx -ne 1){'es'}) included."
                status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'Exchange'; label = 'Mailboxes with Forwarding'; value = Format-ReviewNumber $fwdMbx; rawValue = [int]$fwdMbx
                description = 'Mailboxes configured to forward messages to another address.'
                status = Get-MetricStatus -Name 'fwdMailboxes' -Value $fwdMbx
            })
            $kpis.Add([pscustomobject]@{
                area = 'Exchange'; label = 'External Forwarding'; value = Format-ReviewNumber $extFwd; rawValue = [int]$extFwd
                description = 'Mailboxes suspected of forwarding messages outside the organisation. Should be reviewed.'
                status = Get-MetricStatus -Name 'externalForwarding' -Value $extFwd
            })
            $kpis.Add([pscustomobject]@{
                area = 'Exchange'; label = 'Inbox Forwarding Rules'; value = Format-ReviewNumber $inboxFwd; rawValue = [int]$inboxFwd
                description = 'Inbox rules configured to redirect or forward messages found across scanned mailboxes.'
                status = Get-MetricStatus -Name 'inboxForwardingRules' -Value $inboxFwd
            })
        }

        $spSum = Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'summary'
        if ($spSum) {
            $totalSites = Get-TenantReviewProperty -InputObject $spSum -Name 'totalSites'
            $storageGB  = Get-TenantReviewProperty -InputObject $spSum -Name 'totalStorageGB'
            $extSharing = Get-TenantReviewProperty -InputObject $spSum -Name 'externalSharingEnabledSites'
            $odAccounts = Get-TenantReviewProperty -InputObject $spSum -Name 'oneDriveAccounts'
            $odStorage  = Get-TenantReviewProperty -InputObject $spSum -Name 'oneDriveTotalStorageGB'
            $kpis.Add([pscustomobject]@{
                area = 'SharePoint'; label = 'SharePoint Sites'; value = Format-ReviewNumber $totalSites; rawValue = [int]$totalSites
                description = 'Total SharePoint Online sites in the tenant.'; status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'SharePoint'; label = 'SharePoint Storage'; value = Format-ReviewGB $storageGB; rawValue = [double]$storageGB
                description = 'Total storage used across all SharePoint sites.'; status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'SharePoint'; label = 'External Sharing Sites'; value = Format-ReviewNumber $extSharing; rawValue = [int]$extSharing
                description = 'Sites with external sharing enabled. Should be intentional and regularly reviewed.'
                status = Get-MetricStatus -Name 'externalSharingSites' -Value $extSharing
            })
            if ($null -ne $odAccounts) {
                $kpis.Add([pscustomobject]@{
                    area = 'SharePoint'; label = 'OneDrive Accounts'; value = Format-ReviewNumber $odAccounts; rawValue = [int]$odAccounts
                    description = 'Personal OneDrive accounts provisioned in the tenant.'; status = 'Informational'
                })
            }
            if ($null -ne $odStorage) {
                $kpis.Add([pscustomobject]@{
                    area = 'SharePoint'; label = 'OneDrive Storage'; value = Format-ReviewGB $odStorage; rawValue = [double]$odStorage
                    description = 'Total storage used across all OneDrive personal accounts.'; status = 'Informational'
                })
            }
        }

        $teamsSum = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
        if ($teamsSum) {
            $totalTeams    = Get-TenantReviewProperty -InputObject $teamsSum -Name 'totalTeams'
            $inactiveTeams = Get-TenantReviewProperty -InputObject $teamsSum -Name 'inactiveTeams'
            $noOwnerTeams  = Get-TenantReviewProperty -InputObject $teamsSum -Name 'teamsWithNoOwners'
            $kpis.Add([pscustomobject]@{
                area = 'Teams'; label = 'Total Teams'; value = Format-ReviewNumber $totalTeams; rawValue = [int]$totalTeams
                description = 'Microsoft Teams workspaces in the tenant.'; status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'Teams'; label = 'Inactive Teams'; value = Format-ReviewNumber $inactiveTeams; rawValue = [int]$inactiveTeams
                description = 'Teams with no recorded activity in the last 90 days. May be candidates for archiving.'
                status = Get-MetricStatus -Name 'inactiveTeams' -Value $inactiveTeams
            })
            $kpis.Add([pscustomobject]@{
                area = 'Teams'; label = 'Teams Without Owners'; value = Format-ReviewNumber $noOwnerTeams; rawValue = [int]$noOwnerTeams
                description = 'Teams with no assigned owner — an unmanaged workspace with no accountable steward.'
                status = Get-MetricStatus -Name 'teamsNoOwners' -Value $noOwnerTeams
            })
        }

        $devSum = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
        if ($devSum) {
            $totalDevices = Get-TenantReviewProperty -InputObject $devSum -Name 'totalDevices'
            $staleDevices = Get-TenantReviewProperty -InputObject $devSum -Name 'staleDevices'
            $intuneMgd    = Get-TenantReviewProperty -InputObject $devSum -Name 'intuneManagedDevices'
            $windowsDev   = Get-TenantReviewProperty -InputObject $devSum -Name 'windowsDevices'
            $androidDev   = Get-TenantReviewProperty -InputObject $devSum -Name 'androidDevices'
            $kpis.Add([pscustomobject]@{
                area = 'Devices'; label = 'Total Devices'; value = Format-ReviewNumber $totalDevices; rawValue = [int]$totalDevices
                description = "Devices registered in Entra ID — $windowsDev Windows, $androidDev Android."
                status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'Devices'; label = 'Stale Devices'; value = Format-ReviewNumber $staleDevices; rawValue = [int]$staleDevices
                description = 'Devices that have not checked in recently. May represent decommissioned endpoints.'
                status = Get-MetricStatus -Name 'staleDevices' -Value $staleDevices
            })
            $kpis.Add([pscustomobject]@{
                area = 'Devices'; label = 'Intune Managed'; value = Format-ReviewNumber $intuneMgd; rawValue = [int]$intuneMgd
                description = 'Devices actively managed through Microsoft Intune for compliance and policy enforcement.'
                status = if ([int]$intuneMgd -eq 0 -and [int]$totalDevices -gt 0) { 'Watch' } else { 'Informational' }
            })
        }

        $copSum = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'
        if ($copSum) {
            $copPurchased = Get-TenantReviewProperty -InputObject $copSum -Name 'copilotPurchased'
            $copAssigned  = Get-TenantReviewProperty -InputObject $copSum -Name 'copilotAssigned'
            $copUnused    = Get-TenantReviewProperty -InputObject $copSum -Name 'copilotUnused'
            $copMonthly   = Get-TenantReviewProperty -InputObject $copSum -Name 'estimatedMonthlyCost'
            $activeUsers  = Get-TenantReviewProperty -InputObject $copSum -Name 'activeCopilotUsers'
            $kpis.Add([pscustomobject]@{
                area = 'Copilot'; label = 'Copilot Licenses'; value = Format-ReviewNumber $copPurchased; rawValue = [int]$copPurchased
                description = 'Microsoft 365 Copilot licenses purchased for the tenant.'; status = 'Informational'
            })
            $kpis.Add([pscustomobject]@{
                area = 'Copilot'; label = 'Unused Copilot Licenses'; value = Format-ReviewNumber $copUnused; rawValue = [int]$copUnused
                description = 'Copilot licenses purchased but not yet assigned to any user.'
                status = Get-MetricStatus -Name 'copilotUnused' -Value $copUnused
            })
            if ($null -ne $activeUsers) {
                $kpis.Add([pscustomobject]@{
                    area = 'Copilot'; label = 'Active Copilot Users'; value = Format-ReviewNumber $activeUsers; rawValue = [int]$activeUsers
                    description = 'Users with Copilot who actively used it during the review period.'; status = 'Informational'
                })
            } else {
                $kpis.Add([pscustomobject]@{
                    area = 'Copilot'; label = 'Active Copilot Users'; value = 'N/A'; rawValue = $null
                    description = 'Usage activity data was not collected this run. Enable Copilot usage reporting for adoption insight.'
                    status = 'NotAvailable'
                })
            }
            $kpis.Add([pscustomobject]@{
                area = 'Copilot'; label = 'Copilot Monthly Spend'; value = Format-ReviewCurrency $copMonthly 'CAD'; rawValue = [double]$copMonthly
                description = 'Estimated monthly cost for all Copilot licenses based on available price data.'; status = 'Informational'
            })
        }

        return @($kpis)
    }

    # ─── recommendations builder ───────────────────────────────────────────────

    function New-RecommendationData {
        param([object]$Datasets)
        $recs = [System.Collections.Generic.List[object]]::new()
        $pri  = 1

        $anaSum = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
        $mbxSum = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
        $spSum  = Get-TenantReviewProperty -InputObject $Datasets['SharePoint'] -Name 'summary'
        $teaSum = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
        $devSum = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
        $copSum = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'
        $licSum = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'

        if ($anaSum) {
            $disabled   = [int](Get-TenantReviewProperty -InputObject $anaSum -Name 'disabledLicensedUserCount')
            $stale      = [int](Get-TenantReviewProperty -InputObject $anaSum -Name 'staleLicensedUserCount')
            $unused     = [int](Get-TenantReviewProperty -InputObject $anaSum -Name 'unusedLicenseCount')
            $unusedCost = [double](Get-TenantReviewProperty -InputObject $anaSum -Name 'estimatedUnusedMonthlyCost')
            if ($disabled -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Licensing'; status = 'ActionRecommended'
                    title    = "Remove licenses from $disabled disabled account$(if($disabled -ne 1){'s'})"
                    why      = "$disabled disabled user$(if($disabled -ne 1){'s are'} else {' is'}) still assigned Microsoft 365 licenses. These accounts cannot actively use them and may represent avoidable spend."
                    owner    = 'IT Admin / License Manager'; effort = 'Low'; impact = 'High'
                })
            }
            if ($stale -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Licensing'; status = 'Watch'
                    title    = "Confirm access for $stale stale licensed user$(if($stale -ne 1){'s'})"
                    why      = "$stale licensed account$(if($stale -ne 1){'s have'} else {' has'}) not had a recent successful sign-in. Confirm with managers before removing access or licenses."
                    owner    = 'IT Admin / HR'; effort = 'Low'; impact = 'Medium'
                })
            }
            if ($unused -gt 0) {
                $whyText = if ($unusedCost -gt 0) {
                    "$unused purchased license$(if($unused -ne 1){'s are'} else {' is'}) not currently assigned, representing an estimated $(Format-ReviewCurrency $unusedCost 'CAD')/month in unused capacity."
                } else {
                    "$unused purchased license$(if($unused -ne 1){'s are'} else {' is'}) not currently assigned. Confirm whether they are reserved for upcoming growth or can be reduced at renewal."
                }
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Licensing'; status = if ($unusedCost -gt 0) { 'ActionRecommended' } else { 'Watch' }
                    title    = "Review $unused unassigned license$(if($unused -ne 1){'s'}) before next renewal"
                    why      = $whyText; owner = 'License Manager / Finance'; effort = 'Low'
                    impact   = if ($unusedCost -gt 0) { 'High' } else { 'Medium' }
                })
            }
        }

        if ($mbxSum) {
            $extFwd   = [int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesForwardingExternally')
            $inboxFwd = [int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'inboxForwardingRulesFound')
            $fwdTotal = [int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesWithForwarding')
            if ($extFwd -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Exchange'; status = 'ActionRecommended'
                    title    = "Validate $extFwd mailbox$(if($extFwd -ne 1){'es'}) with suspected external forwarding"
                    why      = "External forwarding may be legitimate or unauthorised. Each instance should be confirmed with the mailbox owner before the next review."
                    owner    = 'Exchange Admin / Security'; effort = 'Low'; impact = 'High'
                })
            }
            if ($inboxFwd -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Exchange'; status = 'Watch'
                    title    = "Review $inboxFwd inbox forwarding rule$(if($inboxFwd -ne 1){'s'})"
                    why      = "Inbox rules that redirect or forward messages can route sensitive information outside the organisation if not governed. Each rule should be confirmed as intentional."
                    owner    = 'Exchange Admin'; effort = 'Medium'; impact = 'Medium'
                })
            } elseif ($fwdTotal -gt 0 -and $extFwd -eq 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Exchange'; status = 'Watch'
                    title    = "Confirm $fwdTotal mailbox forwarding configuration$(if($fwdTotal -ne 1){'s'})"
                    why      = "$fwdTotal mailbox$(if($fwdTotal -ne 1){'es have'} else {' has'}) forwarding configured. No external destinations were detected, but all forwarding should still be confirmed as intentional."
                    owner    = 'Exchange Admin'; effort = 'Low'; impact = 'Low'
                })
            }
        }

        if ($spSum) {
            $extSharing = [int](Get-TenantReviewProperty -InputObject $spSum -Name 'externalSharingEnabledSites')
            if ($extSharing -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'SharePoint'; status = 'Watch'
                    title    = "Review $extSharing SharePoint site$(if($extSharing -ne 1){'s'}) with external sharing enabled"
                    why      = "External sharing should be intentional. Sites that no longer require it should have sharing disabled to reduce data exposure risk."
                    owner    = 'SharePoint Admin'; effort = 'Low'; impact = 'Medium'
                })
            }
        }

        if ($teaSum) {
            $noOwners = [int](Get-TenantReviewProperty -InputObject $teaSum -Name 'teamsWithNoOwners')
            $inactive = [int](Get-TenantReviewProperty -InputObject $teaSum -Name 'inactiveTeams')
            if ($noOwners -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Teams'; status = 'ActionRecommended'
                    title    = "Assign owners to $noOwners Teams workspace$(if($noOwners -ne 1){'s'})"
                    why      = "Teams without owners have no accountable steward for content, membership, or lifecycle decisions. This is a governance risk that is straightforward to resolve."
                    owner    = 'Teams Admin / Department Managers'; effort = 'Low'; impact = 'Medium'
                })
            }
            if ($inactive -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Teams'; status = 'Watch'
                    title    = "Archive or clean up $inactive inactive Team$(if($inactive -ne 1){'s'})"
                    why      = "$inactive Team$(if($inactive -ne 1){'s have'} else {' has'}) had no recorded activity in the last 90 days. Archiving removes clutter and prevents unmanaged access to stale content."
                    owner    = 'Teams Admin / Team Owners'; effort = 'Low'; impact = 'Low'
                })
            }
        }

        if ($devSum) {
            $stale  = [int](Get-TenantReviewProperty -InputObject $devSum -Name 'staleDevices')
            $intune = [int](Get-TenantReviewProperty -InputObject $devSum -Name 'intuneManagedDevices')
            $total  = [int](Get-TenantReviewProperty -InputObject $devSum -Name 'totalDevices')
            if ($stale -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Devices'; status = 'Watch'
                    title    = "Clean up $stale stale device record$(if($stale -ne 1){'s'})"
                    why      = "Stale device records in Entra ID can skew compliance reporting and may represent endpoints that no longer have an active relationship with the tenant."
                    owner    = 'IT Admin / Endpoint Team'; effort = 'Low'; impact = 'Medium'
                })
            }
            if ($intune -eq 0 -and $total -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Devices'; status = 'Watch'
                    title    = 'Consider enrolling devices in Intune for management and compliance coverage'
                    why      = "No devices appear to be Intune-managed. Intune provides compliance checking, policy enforcement, and Conditional Access integration — all of which strengthen security posture."
                    owner    = 'IT Admin / Security'; effort = 'High'; impact = 'High'
                })
            }
        }

        if ($copSum) {
            $copUnused   = [int](Get-TenantReviewProperty -InputObject $copSum -Name 'copilotUnused')
            $activeUsers = Get-TenantReviewProperty -InputObject $copSum -Name 'activeCopilotUsers'
            $usageStatus = Get-TenantReviewProperty -InputObject $copSum -Name 'usageReportStatus'
            if ($copUnused -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Copilot'; status = 'ActionRecommended'
                    title    = "Review $copUnused unused Copilot license$(if($copUnused -ne 1){'s'})"
                    why      = "Microsoft 365 Copilot is a premium SKU. Unused licenses should be reviewed before the next renewal or expansion decision."
                    owner    = 'License Manager'; effort = 'Low'; impact = 'High'
                })
            }
            if ($null -eq $activeUsers -and $usageStatus -ne 'Collected') {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Copilot'; status = 'Watch'
                    title    = 'Enable Copilot usage reporting to measure adoption'
                    why      = "Usage activity data was not collected this run. Without it, it is not possible to confirm whether Copilot licenses are being actively used and delivering value."
                    owner    = 'IT Admin'; effort = 'Low'; impact = 'Medium'
                })
            }
        }

        if ($licSum) {
            $missing = @(Get-TenantReviewProperty -InputObject $licSum -Name 'skusMissingPrice')
            if ($missing.Count -gt 0) {
                $recs.Add([pscustomobject]@{
                    priority = $pri++; category = 'Licensing'; status = 'Watch'
                    title    = "Add price data for $($missing.Count) SKU$(if($missing.Count -ne 1){'s'}) missing cost information"
                    why      = "Cost reporting is incomplete without pricing for all SKUs. Adding price data gives a more accurate and credible view of total monthly spend."
                    owner    = 'License Manager'; effort = 'Low'; impact = 'Low'
                })
            }
        }

        return @($recs)
    }

    # ─── executive summary builder ─────────────────────────────────────────────

    function New-ExecutiveSummaryBullets {
        param([object]$Datasets)
        $bullets = [System.Collections.Generic.List[string]]::new()
        $tenSum  = Get-TenantReviewProperty -InputObject $Datasets['TenantOverview'] -Name 'summary'
        $userSum = Get-TenantReviewProperty -InputObject $Datasets['UserInventory'] -Name 'summary'
        $licSum  = Get-TenantReviewProperty -InputObject $Datasets['LicenseInventory'] -Name 'summary'
        $anaSum  = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
        $mbxSum  = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
        $teaSum  = Get-TenantReviewProperty -InputObject $Datasets['Teams'] -Name 'summary'
        $devSum  = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
        $copSum  = Get-TenantReviewProperty -InputObject $Datasets['Copilot'] -Name 'summary'

        if ($tenSum -and $userSum) {
            $orgName = Get-TenantReviewProperty -InputObject $tenSum -Name 'organizationName'
            $total   = Get-TenantReviewProperty -InputObject $userSum -Name 'totalUsers'
            $members = Get-TenantReviewProperty -InputObject $userSum -Name 'memberUsers'
            $guests  = Get-TenantReviewProperty -InputObject $userSum -Name 'guestUsers'
            if ($orgName) { $bullets.Add("This review covers the $orgName Microsoft 365 tenant. The tenant has $total user accounts — $members members and $guests guest$(if([int]$guests -ne 1){'s'}).") }
        }
        if ($licSum) {
            $assigned = [int](Get-TenantReviewProperty -InputObject $licSum -Name 'totalAssigned')
            $unused   = [int](Get-TenantReviewProperty -InputObject $licSum -Name 'totalUnused')
            $monthly  = [double](Get-TenantReviewProperty -InputObject $licSum -Name 'estimatedMonthlyCost')
            $currency = Get-TenantReviewProperty -InputObject $licSum -Name 'currency'
            if (-not $currency) { $currency = 'CAD' }
            $missing  = @(Get-TenantReviewProperty -InputObject $licSum -Name 'skusMissingPrice')
            if ($monthly -gt 0) {
                $costNote = if ($missing.Count -gt 0) { " (partial — $($missing.Count) SKU$(if($missing.Count -ne 1){'s'}) are missing price data)" } else { '' }
                $bullets.Add("$assigned Microsoft 365 license$(if($assigned -ne 1){'s'}) are currently assigned with an estimated monthly spend of $currency `$$([Math]::Round($monthly,2))$costNote. $unused license$(if($unused -ne 1){'s'}) appear unassigned.")
            } else {
                $bullets.Add("$assigned Microsoft 365 license$(if($assigned -ne 1){'s'}) are currently assigned. $unused license$(if($unused -ne 1){'s'}) appear unassigned. Full cost figures require price data for all SKUs.")
            }
        }
        if ($anaSum) {
            $disabled = [int](Get-TenantReviewProperty -InputObject $anaSum -Name 'disabledLicensedUserCount')
            $stale    = [int](Get-TenantReviewProperty -InputObject $anaSum -Name 'staleLicensedUserCount')
            if ($disabled -gt 0 -or $stale -gt 0) {
                $parts = @()
                if ($disabled -gt 0) { $parts += "$disabled disabled account$(if($disabled -ne 1){'s'}) still holding licenses" }
                if ($stale -gt 0) { $parts += "$stale licensed account$(if($stale -ne 1){'s'}) that $(if($stale -ne 1){'have'} else {'has'}) not signed in recently" }
                $bullets.Add("License optimisation opportunities were found: $($parts -join ', and '). These are straightforward to address before the next renewal cycle.")
            }
        }
        if ($mbxSum) {
            $extFwd   = [int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesForwardingExternally')
            $fwd      = [int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesWithForwarding')
            $inboxFwd = [int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'inboxForwardingRulesFound')
            if ($extFwd -gt 0) {
                $bullets.Add("$extFwd mailbox$(if($extFwd -ne 1){'es are'} else {' is'}) suspected of forwarding messages externally. This should be reviewed to confirm the configuration is intentional and authorised.")
            } elseif ($fwd -gt 0 -or $inboxFwd -gt 0) {
                $bullets.Add("Mail forwarding is active on $fwd mailbox$(if($fwd -ne 1){'es'}), and $inboxFwd inbox forwarding rule$(if($inboxFwd -ne 1){'s were'} else {' was'}) found. These should be confirmed as intentional.")
            }
        }
        if ($teaSum) {
            $inactive = [int](Get-TenantReviewProperty -InputObject $teaSum -Name 'inactiveTeams')
            $total    = [int](Get-TenantReviewProperty -InputObject $teaSum -Name 'totalTeams')
            if ($total -gt 0 -and $inactive -gt 0) {
                $bullets.Add("$inactive of $total Teams workspace$(if($total -ne 1){'s'}) had no recorded activity in the last 90 days and may be candidates for archiving.")
            }
        }
        if ($devSum) {
            $stale = [int](Get-TenantReviewProperty -InputObject $devSum -Name 'staleDevices')
            $total = [int](Get-TenantReviewProperty -InputObject $devSum -Name 'totalDevices')
            if ($stale -gt 0) { $bullets.Add("$stale of $total registered device$(if($total -ne 1){'s'}) appear stale. Cleaning up outdated device records supports better compliance and security reporting.") }
        }
        if ($copSum) {
            $purchased = [int](Get-TenantReviewProperty -InputObject $copSum -Name 'copilotPurchased')
            $assigned  = [int](Get-TenantReviewProperty -InputObject $copSum -Name 'copilotAssigned')
            $monthly   = [double](Get-TenantReviewProperty -InputObject $copSum -Name 'estimatedMonthlyCost')
            if ($purchased -gt 0) {
                $bullets.Add("$assigned Microsoft 365 Copilot license$(if($assigned -ne 1){'s'}) $(if($assigned -ne 1){'are'} else {'is'}) assigned at an estimated CAD `$$([Math]::Round($monthly,2))/month. Enable usage reporting to confirm active adoption.")
            }
        }
        return @($bullets)
    }

    function New-TenantHealthStatement {
        param([object]$Datasets)
        $anaSum = Get-TenantReviewProperty -InputObject $Datasets['LicenseUserAnalysis'] -Name 'summary'
        $mbxSum = Get-TenantReviewProperty -InputObject $Datasets['MailboxInventory'] -Name 'summary'
        $devSum = Get-TenantReviewProperty -InputObject $Datasets['Devices'] -Name 'summary'
        $action = 0; $watch = 0
        if ($anaSum) {
            if ([int](Get-TenantReviewProperty -InputObject $anaSum -Name 'disabledLicensedUserCount') -gt 0) { $action++ }
            if ([int](Get-TenantReviewProperty -InputObject $anaSum -Name 'staleLicensedUserCount') -gt 0) { $watch++ }
        }
        if ($mbxSum) {
            if ([int](Get-TenantReviewProperty -InputObject $mbxSum -Name 'mailboxesForwardingExternally') -gt 0) { $action++ }
        }
        if ($devSum) {
            if ([int](Get-TenantReviewProperty -InputObject $devSum -Name 'staleDevices') -gt 5) { $watch++ }
        }
        if ($action -eq 0 -and $watch -eq 0) { return 'The tenant is in good standing. No major issues were identified during this review.' }
        elseif ($action -eq 0) { return 'The tenant is generally in good shape. A small number of items are flagged for monitoring to keep the environment well-maintained.' }
        elseif ($action -le 2) { return 'The tenant review identified a small number of items recommended for action before the next renewal or review cycle.' }
        else { return 'The tenant review identified several areas that would benefit from prompt attention. Prioritised recommendations are outlined in this report.' }
    }

    # ─── local narrative section builder ─────────────────────────────────────

    function New-LocalNarrativeSection {
        param([string]$Dataset, [object]$Summary)

        switch ($Dataset) {
            'TenantOverview' {
                $org     = Get-TenantReviewProperty -InputObject $Summary -Name 'organizationName'
                $domains = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'domainCount')
                $verified = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'verifiedDomainCount')
                $federated = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'federatedDomainCount')
                $default = Get-TenantReviewProperty -InputObject $Summary -Name 'defaultDomain'
                $created = Get-TenantReviewProperty -InputObject $Summary -Name 'createdDateTime'
                $techContacts = @(Get-TenantReviewProperty -InputObject $Summary -Name 'technicalNotificationMails')
                $assignedPlans = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'assignedPlansCount')
                $year    = if ($created) { try { ([datetime]$created).Year } catch { 'Unknown' } } else { 'Unknown' }
                $detailText = "The tenant foundation shows $verified verified domain$(if($verified -ne 1){'s'}) out of $domains total, with $default serving as the default sign-in and mail identity namespace."
                if ($federated -gt 0) { $detailText += " $federated domain$(if($federated -ne 1){'s are'} else {' is'}) configured for federation, which should stay aligned with identity provider policy and lifecycle ownership." }
                if ($techContacts.Count -gt 0) { $detailText += " Technical notifications currently route to $($techContacts -join ', '), and $assignedPlans service plan$(if($assignedPlans -ne 1){'s are'} else {' is'}) assigned across the tenant." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = 'Good'
                    headline          = if ($org) { "$org — active since $year" } else { "Tenant active since $year" }
                    plainEnglish      = "The tenant for$(if($org){" $org"}) has been active since $year and is configured with $domains verified domain$(if($domains -ne 1){'s'}), with $default as the primary domain."
                    detailedAnalysis  = $detailText
                    businessImpact    = 'A well-configured tenant foundation supports reliable identity, mail, and collaboration services for the business.'
                    recommendedAction = 'Review technical notification contacts and domain settings periodically to ensure they remain current.'
                }
            }
            'LicenseInventory' {
                $assigned = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalAssigned')
                $unused   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalUnused')
                $purchased = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalPurchased')
                $monthly  = [double](Get-TenantReviewProperty -InputObject $Summary -Name 'estimatedMonthlyCost')
                $currency = Get-TenantReviewProperty -InputObject $Summary -Name 'currency'; if (-not $currency) { $currency = 'CAD' }
                $missing  = @(Get-TenantReviewProperty -InputObject $Summary -Name 'skusMissingPrice')
                $plainText = "$assigned license$(if($assigned -ne 1){'s'}) are currently assigned across all SKUs in the tenant."
                if ($missing.Count -gt 0) { $plainText += " Cost data is unavailable for $($missing.Count) SKU$(if($missing.Count -ne 1){'s'}), so total spend figures may be understated." }
                $detailText = "The tenant has $purchased purchased seats with $assigned currently assigned and $unused unassigned, providing the baseline for renewal and cost optimisation decisions."
                if ($monthly -gt 0) { $detailText += " Available pricing data estimates monthly spend at $(Format-ReviewCurrency $monthly $currency)." }
                if ($missing.Count -gt 0) { $detailText += " Pricing is still missing for $($missing.Count) SKU$(if($missing.Count -ne 1){'s'}), so the reported spend should be treated as a partial view until price mapping is completed." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = if ($unused -eq 0) { 'Good' } else { 'Watch' }
                    headline          = "$assigned licenses assigned — $unused unassigned across all SKUs"
                    plainEnglish      = $plainText
                    detailedAnalysis  = $detailText
                    businessImpact    = 'Understanding what is purchased versus assigned is the starting point for license cost management and renewal planning.'
                    recommendedAction = if ($missing.Count -gt 0) { "Add price data for the $($missing.Count) missing SKU$(if($missing.Count -ne 1){'s'}) to get a complete spend view, then review unassigned licenses before the next renewal." } else { 'Review unassigned licenses to confirm whether they are needed for upcoming users or can be reduced at renewal.' }
                }
            }
            'UserInventory' {
                $total    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalUsers')
                $members  = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'memberUsers')
                $guests   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'guestUsers')
                $disabled = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'disabledUsers')
                $licensed = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'licensedUsers')
                $stale = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'staleUsers')
                $signStatus = Get-TenantReviewProperty -InputObject $Summary -Name 'signInActivityStatus'
                $signNote = switch ($signStatus) {
                    'NotRequested' { ' Sign-in activity data was not collected this run — enable it to detect dormant accounts.' }
                    'Unavailable'  { ' Sign-in activity data was unavailable for this tenant.' }
                    default        { '' }
                }
                $detailText = "$members member account$(if($members -ne 1){'s'}) and $guests guest account$(if($guests -ne 1){'s'}) are currently present in the directory, with $licensed licensed user$(if($licensed -ne 1){'s'}) and $disabled disabled account$(if($disabled -ne 1){'s'})."
                if ($signStatus -eq 'Collected') { $detailText += " Sign-in telemetry indicates $stale stale account$(if($stale -ne 1){'s'}) based on the configured inactivity threshold, which should be compared against HR and vendor access records." }
                else { $detailText += " Sign-in telemetry was not available in this run, so dormant-account risk should be treated as under-reported until activity collection is enabled consistently." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = if ($disabled -gt 0 -or ([int]$guests -gt ([int]$total * 0.3))) { 'Watch' } else { 'Good' }
                    headline          = "$total users — $members members, $guests guest$(if($guests -ne 1){'s'})$(if($disabled -gt 0){", $disabled disabled"})"
                    plainEnglish      = "The tenant has $total user accounts: $members internal members and $guests external guest$(if($guests -ne 1){'s'}).$signNote"
                    detailedAnalysis  = $detailText
                    businessImpact    = 'Regular user reviews keep the tenant secure and reduce the risk of dormant or over-privileged accounts retaining access to business data.'
                    recommendedAction = if ($guests -gt 0) { "Review guest accounts to confirm each still has a clear, current business reason to access tenant resources." } else { 'Periodically review user accounts to ensure access aligns with current roles and employment status.' }
                }
            }
            'LicenseUserAnalysis' {
                $unused    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'unusedLicenseCount')
                $disabled  = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'disabledLicensedUserCount')
                $stale     = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'staleLicensedUserCount')
                $unusedCost = [double](Get-TenantReviewProperty -InputObject $Summary -Name 'estimatedUnusedMonthlyCost')
                $parts = @(); if ($disabled -gt 0) { $parts += "$disabled disabled account$(if($disabled -ne 1){'s'}) with licenses" }; if ($stale -gt 0) { $parts += "$stale stale licensed account$(if($stale -ne 1){'s'})" }
                $headline  = if ($parts.Count -gt 0) { "License review recommended — $($parts -join ' and ')" } else { "License assignment is clean — no obvious waste identified" }
                $plainText = if ($unusedCost -gt 0) { "$unused license$(if($unused -ne 1){'s'}) are unassigned, representing an estimated CAD `$$([Math]::Round($unusedCost,2))/month in unused capacity." }
                             elseif ($unused -gt 0)  { "$unused license$(if($unused -ne 1){'s'}) are unassigned. Some SKUs are missing price data, so the exact cost impact is not yet calculable." }
                             else                    { "All purchased licenses with available pricing appear to be assigned. No obvious unused capacity was identified." }
                $detailText = "$disabled disabled user account$(if($disabled -ne 1){'s'}) still hold paid licensing, while $stale licensed user$(if($stale -ne 1){'s'}) appear inactive enough to warrant a manager validation before renewal."
                if ($unusedCost -gt 0) { $detailText += " Unassigned capacity is estimated at CAD `$$([Math]::Round($unusedCost,2))/month, which is the clearest recurring savings opportunity in the current data set." }
                elseif ($unused -gt 0) { $detailText += " Unassigned license counts were found, but some SKU pricing is incomplete, so the financial impact is directional rather than final." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = if ($disabled -gt 0) { 'ActionRecommended' } elseif ($stale -gt 0 -or $unused -gt 5) { 'Watch' } else { 'Good' }
                    headline = $headline; plainEnglish = $plainText; detailedAnalysis = $detailText
                    businessImpact    = 'Unused and misallocated licenses represent avoidable recurring spend. Even small monthly amounts compound over a 12-month agreement.'
                    recommendedAction = if ($disabled -gt 0) { "Start by reviewing the $disabled disabled account$(if($disabled -ne 1){'s'}) still holding licenses — these are the clearest immediate candidates." } elseif ($stale -gt 0) { "Confirm whether the $stale stale licensed user$(if($stale -ne 1){'s'}) still need access before the next renewal." } else { 'Review unassigned licenses to confirm they are intentionally reserved for upcoming growth.' }
                }
            }
            'MailboxInventory' {
                $total    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalMailboxes')
                $shared   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'sharedMailboxes')
                $fwd      = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'mailboxesWithForwarding')
                $extFwd   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'mailboxesForwardingExternally')
                $rules    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'enabledTransportRules')
                $inboxFwd = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'inboxForwardingRulesFound')
                $headline = if ($extFwd -gt 0) { "External mailbox forwarding detected — review recommended" } elseif ($fwd -gt 0) { "$fwd mailbox$(if($fwd -ne 1){'es'}) with forwarding configured" } else { "$total mailboxes reviewed — no forwarding concerns detected" }
                $plainText = "The tenant has $total mailboxes, including $shared shared mailbox$(if($shared -ne 1){'es'}). $fwd $(if($fwd -ne 1){'have'} else {'has'}) forwarding configured."
                if ($extFwd -gt 0) { $plainText += " $extFwd appear$(if($extFwd -eq 1){'s'}) to forward messages externally." }
                if ($inboxFwd -gt 0) { $plainText += " $inboxFwd inbox forwarding rule$(if($inboxFwd -ne 1){'s were'} else {' was'}) found." }
                if ($rules -gt 0) { $plainText += " $rules transport rule$(if($rules -ne 1){'s are'} else {' is'}) currently active." }
                $detailText = "$total Exchange Online mailbox$(if($total -ne 1){'es were'} else {' was'}) inventoried, including $shared shared mailbox$(if($shared -ne 1){'es'}) used for team or workflow scenarios."
                if ($fwd -gt 0) { $detailText += " Forwarding is configured on $fwd mailbox$(if($fwd -ne 1){'es'}), with $extFwd suspected external destination$(if($extFwd -ne 1){'s'}) and $inboxFwd inbox rule$(if($inboxFwd -ne 1){'s'}) requiring line-by-line review in the detailed appendix." }
                if ($rules -gt 0) { $detailText += " $rules active transport rule$(if($rules -ne 1){'s'}) also affect tenant-wide mail flow and should be reviewed alongside mailbox-level forwarding." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'Medium'; status = if ($extFwd -gt 0) { 'ActionRecommended' } elseif ($fwd -gt 0 -or $inboxFwd -gt 5) { 'Watch' } else { 'Good' }
                    headline = $headline; plainEnglish = $plainText; detailedAnalysis = $detailText
                    businessImpact    = 'Mail forwarding can be legitimate, but unreviewed forwarding — especially to external addresses — creates data exposure and compliance risk.'
                    recommendedAction = if ($extFwd -gt 0) { "Validate each of the $extFwd externally forwarding mailbox$(if($extFwd -ne 1){'es'}) and remove any configurations that are no longer required." } elseif ($fwd -gt 0) { "Confirm that each mailbox forwarding configuration is intentional, documented, and still required." } else { 'Continue to monitor transport rules and inbox rules periodically.' }
                }
            }
            'SharePoint' {
                $sites   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalSites')
                $stGB    = [double](Get-TenantReviewProperty -InputObject $Summary -Name 'totalStorageGB')
                $extSh   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'externalSharingEnabledSites')
                $lgTitle = Get-TenantReviewProperty -InputObject $Summary -Name 'largestSiteTitle'
                $lgGB    = [double](Get-TenantReviewProperty -InputObject $Summary -Name 'largestSiteStorageGB')
                $plainText = "$sites SharePoint sites are active, using $('{0:N1}' -f $stGB) GB of storage."
                if ($lgTitle -and $lgGB -gt 0) { $plainText += " The largest site ('$lgTitle') uses $('{0:N1}' -f $lgGB) GB." }
                if ($extSh -gt 0) { $plainText += " $extSh site$(if($extSh -ne 1){'s have'} else {' has'}) external sharing enabled." } else { $plainText += " No sites have external sharing enabled." }
                $detailText = "SharePoint storage currently totals $('{0:N1}' -f $stGB) GB across $sites site$(if($sites -ne 1){'s'}), with '$lgTitle' currently the largest footprint at $('{0:N1}' -f $lgGB) GB."
                if ($extSh -gt 0) { $detailText += " External sharing remains enabled on $extSh site$(if($extSh -ne 1){'s'}), so the detailed report should be used to confirm business justification, ownership, and sensitivity of exposed content." }
                else { $detailText += " No externally shared sites were identified in the collected data, reducing the immediate external collaboration risk surface." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = if ($extSh -gt 0) { 'Watch' } else { 'Good' }
                    headline          = "$sites SharePoint sites — $('{0:N1}' -f $stGB) GB used$(if($extSh -gt 0){", $extSh with external sharing"})"
                    plainEnglish      = $plainText
                    detailedAnalysis  = $detailText
                    businessImpact    = 'SharePoint is often the primary document store. Governance of site size, access, and sharing settings directly affects data security and storage costs.'
                    recommendedAction = if ($extSh -gt 0) { "Review the $extSh externally shared site$(if($extSh -ne 1){'s'}) to confirm sharing is still needed and appropriately scoped." } else { 'Monitor storage growth and review large sites for lifecycle or archiving decisions.' }
                }
            }
            'Teams' {
                $total    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalTeams')
                $inactive = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'inactiveTeams')
                $noOwners = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'teamsWithNoOwners')
                $plainText = "The tenant has $total Teams workspace$(if($total -ne 1){'s'}). $inactive appear$(if($inactive -eq 1){'s'}) to have been inactive in the last 90 days."
                if ($noOwners -gt 0) { $plainText += " $noOwners team$(if($noOwners -ne 1){'s have'} else {' has'}) no assigned owner." }
                $detailText = "$total Teams workspace$(if($total -ne 1){'s were'} else {' was'}) identified in scope, and $inactive of them show no recent collaboration activity over the reporting window."
                if ($noOwners -gt 0) { $detailText += " $noOwners ownerless Team$(if($noOwners -ne 1){'s'}) create a direct governance gap because no accountable person is assigned to membership, retention, or archival decisions." }
                else { $detailText += " Owner coverage appears complete in the collected data, which simplifies lifecycle review and accountability." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = if ($noOwners -gt 0) { 'ActionRecommended' } elseif ($inactive -gt 0) { 'Watch' } else { 'Good' }
                    headline          = if ($inactive -gt 0) { "$inactive of $total Teams $(if($inactive -ne 1){'are'} else {'is'}) inactive — review for archiving" } else { "$total Teams reviewed — all appear active" }
                    plainEnglish      = $plainText
                    detailedAnalysis  = $detailText
                    businessImpact    = 'Inactive Teams create unmanaged collaboration spaces with unreviewed content and access. Teams without owners have no accountable steward.'
                    recommendedAction = if ($noOwners -gt 0) { "Assign owners to the $noOwners ownerless Team$(if($noOwners -ne 1){'s'}) first, then archive inactive ones." } elseif ($inactive -gt 0) { "Contact the owners of the $inactive inactive Team$(if($inactive -ne 1){'s'}) to confirm whether they should remain active or be archived." } else { 'No immediate action required. Review periodically for new inactive workspaces.' }
                }
            }
            'Devices' {
                $total    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'totalDevices')
                $stale    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'staleDevices')
                $windows  = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'windowsDevices')
                $intune   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'intuneManagedDevices')
                $unknown  = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'unknownComplianceDevices')
                $intuneNote = if ($intune -eq 0) { ' No devices appear to be Intune-managed.' } else { " $intune $(if($intune -ne 1){'are'} else {'is'}) Intune-managed." }
                $plainText = "$total devices are registered in Entra ID. $stale appear$(if($stale -eq 1){'s'}) stale.$intuneNote"
                if ($unknown -gt 0 -and $intune -eq 0) { $plainText += " $unknown device$(if($unknown -ne 1){'s have'} else {' has'}) unknown compliance status." }
                $detailText = "$windows Windows device$(if($windows -ne 1){'s'}) make up the known managed-desktop footprint inside a total of $total registered device$(if($total -ne 1){'s'})."
                $detailText += " $stale stale device record$(if($stale -ne 1){'s'}) and $unknown unknown-compliance device$(if($unknown -ne 1){'s'}) indicate where endpoint hygiene or reporting coverage should be tightened."
                if ($intune -eq 0) { $detailText += " No Intune-managed devices were identified in summary data, so conditional access and compliance control should be reviewed carefully against the current endpoint strategy." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'Medium'; status = if ($stale -gt 5) { 'ActionRecommended' } elseif ($stale -gt 0 -or $intune -eq 0) { 'Watch' } else { 'Good' }
                    headline          = "$total devices registered — $stale stale$(if($intune -eq 0){', Intune coverage not confirmed'})"
                    plainEnglish      = $plainText
                    detailedAnalysis  = $detailText
                    businessImpact    = 'A clean device registry supports accurate compliance reporting and helps ensure only active, known endpoints retain access to company resources.'
                    recommendedAction = if ($stale -gt 0 -and $intune -eq 0) { "Remove stale device records and explore Intune enrolment to gain device compliance visibility." } elseif ($stale -gt 0) { "Review and remove the $stale stale device record$(if($stale -ne 1){'s'})." } elseif ($intune -eq 0) { "Consider enrolling devices in Intune to gain compliance visibility and management coverage." } else { 'Device inventory is in good shape. Review periodically.' }
                }
            }
            'Copilot' {
                $purchased   = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'copilotPurchased')
                $assigned    = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'copilotAssigned')
                $unused      = [int](Get-TenantReviewProperty -InputObject $Summary -Name 'copilotUnused')
                $monthly     = [double](Get-TenantReviewProperty -InputObject $Summary -Name 'estimatedMonthlyCost')
                $usageStatus = Get-TenantReviewProperty -InputObject $Summary -Name 'usageReportStatus'
                $activeUsers = Get-TenantReviewProperty -InputObject $Summary -Name 'activeCopilotUsers'
                $usageNote   = switch ($usageStatus) {
                    'NotRequested' { ' Usage activity data was not collected this run — enable it to measure adoption.' }
                    'Unavailable'  { ' Usage activity data was unavailable for this tenant.' }
                    default        { '' }
                }
                $plainText = "$purchased Copilot license$(if($purchased -ne 1){'s'}) $(if($purchased -ne 1){'are'} else {'is'}) in the tenant, of which $assigned $(if($assigned -ne 1){'are'} else {'is'}) assigned.$usageNote"
                $detailText = "Copilot licensing currently shows $purchased purchased, $assigned assigned, and $unused unassigned seat$(if($unused -ne 1){'s'}), with estimated spend of CAD `$$([Math]::Round($monthly,2))/month."
                if ($null -ne $activeUsers) { $detailText += " Usage reporting recorded $activeUsers active Copilot user$(if([int]$activeUsers -ne 1){'s'}) during the review window, which should be compared with assigned seats to measure adoption depth." }
                else { $detailText += " Usage reporting was not available in collected data, so the current value assessment is based on assigned licenses rather than observed adoption." }
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'High'; status = if ($unused -gt 0) { 'ActionRecommended' } elseif ($usageStatus -eq 'NotRequested') { 'Watch' } else { 'Good' }
                    headline          = "$assigned Copilot license$(if($assigned -ne 1){'s'}) assigned$(if($unused -gt 0){" — $unused unused"}) at CAD `$$([Math]::Round($monthly,2))/month"
                    plainEnglish      = $plainText
                    detailedAnalysis  = $detailText
                    businessImpact    = "At CAD `$$([Math]::Round($monthly,2))/month, Copilot is a high-value SKU. Confirming adoption helps justify the investment and guides renewal decisions."
                    recommendedAction = if ($unused -gt 0) { "Reassign or remove the $unused unused Copilot license$(if($unused -ne 1){'s'}) before the next renewal." } elseif ($usageStatus -eq 'NotRequested') { "Enable Copilot usage reporting to confirm that licensed users are actively using and benefiting from the product." } else { 'Continue to monitor Copilot adoption metrics to support the next renewal decision.' }
                }
            }
            default {
                return [pscustomobject]@{
                    dataset = $Dataset; confidence = 'Medium'; status = 'Informational'
                    headline          = "$Dataset data collected"
                    plainEnglish      = 'This dataset was collected where module and permission coverage allowed.'
                    detailedAnalysis  = 'This section is supported by collected tenant data and should be reviewed alongside any warnings or skipped-coverage notes before presenting the final package.'
                    businessImpact    = 'Supporting data for tenant health and planning.'
                    recommendedAction = 'Review any collector warnings before presenting the final package.'
                }
            }
        }
    }

    function New-LocalNarrative {
        param([object]$Datasets)
        $sections = [System.Collections.Generic.List[object]]::new()
        foreach ($key in $Datasets.Keys) {
            $sections.Add((New-LocalNarrativeSection -Dataset $key -Summary (Get-TenantReviewProperty -InputObject $Datasets[$key] -Name 'summary')))
        }
        return @($sections)
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

    # ─── determine sections (AI or local) ────────────────────────────────────

    $aiSettings = Get-TenantReviewProperty -InputObject $Settings -Name 'ai'
    $aiEnabled  = Test-TenantReviewTruthy -Value (Get-TenantReviewProperty -InputObject $aiSettings -Name 'enabled')

    $sections = $null
    $source   = 'LocalRuleBased'

    if (-not $aiEnabled) {
        $sections = New-LocalNarrative -Datasets $Datasets
    } else {
        $endpoint = Get-TenantReviewProperty -InputObject $aiSettings -Name 'endpoint'
        $apiKey   = Resolve-TenantReviewAiApiKey
        if (-not $endpoint -or -not $apiKey) {
            throw 'AI narrative is enabled, but the endpoint or API key is missing. Provide a valid ai.endpoint and API key in Settings.json or the configured environment variable.'
        }

        $datasetSummaries = @(
            foreach ($key in $Datasets.Keys) {
                [pscustomobject]@{
                    dataset  = $key
                    summary  = Get-TenantReviewProperty -InputObject $Datasets[$key] -Name 'summary'
                    warnings = Get-TenantReviewProperty -InputObject $Datasets[$key] -Name 'warnings'
                }
            }
        )

        $prompt = @"
Return strict JSON only as one top-level object with this exact shape:
{"sections":[{"dataset":"DatasetName","headline":"Short headline","plainEnglish":"One or two sentences for an executive HTML report","detailedAnalysis":"Two to four sentences for a detailed technical Word report","businessImpact":"One sentence","recommendedAction":"One sentence","confidence":"High|Medium|Low","status":"Good|Watch|ActionRecommended|Informational"}]}
Create concise, professional Microsoft 365 tenant review narrative sections for a client-facing executive report and a more detailed companion technical report.
Each section must include all fields. Use only the supplied data. `plainEnglish` should stay high-level and accessible. `detailedAnalysis` should be more specific, using concrete counts, costs, or coverage details from the supplied dataset summaries.
Data:
$($datasetSummaries | ConvertTo-Json -Depth 8)
"@

        try {
            $model   = Get-TenantReviewProperty -InputObject $aiSettings -Name 'model'
            $request = Resolve-TenantReviewAiRequest -Endpoint $endpoint -Model $model
            if ($request.Kind -eq 'Responses') {
                $bodyObject = @{
                    model = $model
                    input = @(
                        @{ role = 'system'; content = 'You produce strict JSON only for client-ready Microsoft 365 tenant review narrative. Provide both an executive plain-English summary and a more detailed analysis paragraph for each section.' }
                        @{ role = 'user'; content = $prompt }
                    )
                }
            } else {
                $bodyObject = @{
                    messages = @(
                        @{ role = 'system'; content = 'You produce strict JSON only for client-ready Microsoft 365 tenant review narrative. Provide both an executive plain-English summary and a more detailed analysis paragraph for each section.' }
                        @{ role = 'user'; content = $prompt }
                    )
                }
                if ($model) { $bodyObject.model = $model }
                $temperature = Get-AiSettingValue -Names @('temperature')
                if ($null -ne $temperature) { $bodyObject.temperature = $temperature }
                $useJsonFormat = Get-AiSettingValue -Names @('responseFormatJson')
                if ($null -eq $useJsonFormat -or (Test-TenantReviewTruthy -Value $useJsonFormat)) {
                    $bodyObject.response_format = @{ type = 'json_object' }
                }
            }

            $body     = $bodyObject | ConvertTo-Json -Depth 12
            $headers  = if ($request.IsAzure) {
                @{ 'api-key' = $apiKey; 'Content-Type' = 'application/json' }
            } else {
                @{ Authorization = "Bearer $apiKey"; 'Content-Type' = 'application/json' }
            }
            $response = Invoke-RestMethod -Method Post -Uri $request.Uri -Headers $headers -Body $body -ErrorAction Stop
            $content  = Get-TenantReviewAiResponseText -Response $response
            if (-not $content) { throw 'AI response did not include message content.' }

            $parsed        = ConvertFrom-TenantReviewAiJsonText -Text $content
            $sectionsValue = Get-TenantReviewProperty -InputObject $parsed -Name 'sections'
            if ($null -eq $sectionsValue) { $sectionsValue = Get-TenantReviewProperty -InputObject $parsed -Name 'narrativeSections' }
            if ($null -eq $sectionsValue) { $sectionsValue = Get-TenantReviewProperty -InputObject $parsed -Name 'items' }
            if ($null -eq $sectionsValue) {
                $narrativeObject = Get-TenantReviewProperty -InputObject $parsed -Name 'narrative'
                $sectionsValue   = Get-TenantReviewProperty -InputObject $narrativeObject -Name 'sections'
            }
            $sections = @($sectionsValue | Where-Object { $null -ne $_ })
            if ($sections.Count -eq 0 -and $parsed -is [array]) { $sections = @($parsed | Where-Object { $null -ne $_ }) }
            if ($sections.Count -eq 0 -and (Get-TenantReviewProperty -InputObject $parsed -Name 'dataset')) { $sections = @($parsed) }
            if ($sections.Count -eq 0) { throw 'AI response JSON did not include narrative sections.' }
            foreach ($section in $sections) {
                foreach ($req in @('dataset', 'headline', 'plainEnglish', 'businessImpact', 'recommendedAction', 'confidence')) {
                    if (-not (Get-TenantReviewProperty -InputObject $section -Name $req)) {
                        throw "AI response section was missing required property '$req'."
                    }
                }
                if (-not (Get-TenantReviewProperty -InputObject $section -Name 'detailedAnalysis')) {
                    Add-Member -InputObject $section -MemberType NoteProperty -Name 'detailedAnalysis' -Value (Get-TenantReviewProperty -InputObject $section -Name 'plainEnglish') -Force
                }
            }
            $source = 'AI'
        } catch {
            throw "AI narrative failed. The script will not use local fallback because AI was enabled. $($_.Exception.Message)"
        }
    }

    # ─── always compute KPIs and recommendations locally ─────────────────────

    $kpis            = New-KpiData -Datasets $Datasets
    $recommendations = New-RecommendationData -Datasets $Datasets
    $execBullets     = New-ExecutiveSummaryBullets -Datasets $Datasets
    $healthStatement = New-TenantHealthStatement -Datasets $Datasets

    [pscustomobject]@{
        dataset                 = 'Narrative'
        generatedAt             = (Get-Date).ToString('o')
        source                  = $source
        tenantHealthStatement   = $healthStatement
        executiveSummaryBullets = @($execBullets)
        sections                = @($sections)
        kpis                    = @($kpis)
        recommendations         = @($recommendations)
        warnings                = @()
    }
}
