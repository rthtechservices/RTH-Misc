function Get-TeamsInventory {
    [CmdletBinding()]
    param(
        [switch]$IncludeRaw,

        [int]$ReportPeriodDays = 90,

        [switch]$IncludeOwnersAndMembers
    )

    $warnings = @()
    $teamGroups = @()
    $teamActivity = @()
    $raw = [ordered]@{}

    $groupCommand = Get-Command -Name Get-MgGroup -ErrorAction SilentlyContinue
    if (Get-Command -Name Invoke-TenantReviewGraphRestRequest -ErrorAction SilentlyContinue) {
        try {
            $teamGroups = @(Invoke-TenantReviewGraphRestRequest -Uri "groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,visibility,createdDateTime,resourceProvisioningOptions" -All)
        } catch {
            $warnings += "Unable to collect Microsoft 365 groups with Teams via Graph REST. $($_.Exception.Message)"
        }
    } elseif ($groupCommand) {
        try {
            $teamGroups = @(Get-MgGroup -All -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Property @('Id', 'DisplayName', 'Visibility', 'CreatedDateTime', 'ResourceProvisioningOptions') -ErrorAction Stop)
        } catch {
            $warnings += "Filtered Teams group query failed; retrying broader group query. $($_.Exception.Message)"
            try {
                $groups = @(Get-MgGroup -All -Property @('Id', 'DisplayName', 'Visibility', 'CreatedDateTime', 'ResourceProvisioningOptions') -ErrorAction Stop)
                $teamGroups = @($groups | Where-Object { @(Get-TenantReviewProperty -InputObject $_ -Name 'ResourceProvisioningOptions') -contains 'Team' })
            } catch {
                $warnings += "Unable to collect Microsoft 365 groups with Teams. $($_.Exception.Message)"
            }
        }
    } else {
        $warnings += 'Get-MgGroup is not available. Install/import Microsoft.Graph.Groups for Teams inventory.'
    }

    try {
        $teamActivity = @(Invoke-TenantReviewGraphCsvReport -CommandName 'Get-MgReportTeamActivityDetail' -ReportPeriodDays $ReportPeriodDays)
        $raw['teamActivityDetail'] = @($teamActivity)
    } catch {
        try {
            $teamActivity = @(Invoke-TenantReviewGraphCsvReport -CommandName 'Get-MgReportTeamsTeamActivityDetail' -ReportPeriodDays $ReportPeriodDays)
            $raw['teamActivityDetail'] = @($teamActivity)
        } catch {
            $warnings += "Teams activity report unavailable. $($_.Exception.Message)"
        }
    }

    $activityByTeamId = @{}
    $activityByName = @{}
    foreach ($row in $teamActivity) {
        $teamId = Get-TenantReviewProperty -InputObject $row -Name 'Team Id'
        if (-not $teamId) { $teamId = Get-TenantReviewProperty -InputObject $row -Name 'Team ID' }
        $teamName = Get-TenantReviewProperty -InputObject $row -Name 'Team Name'
        if ($teamId) { $activityByTeamId[$teamId.ToString()] = $row }
        if ($teamName) { $activityByName[$teamName.ToString().ToLowerInvariant()] = $row }
    }

    $items = @()
    $staleCutoff = (Get-Date).AddDays(-1 * $ReportPeriodDays)
    foreach ($group in $teamGroups) {
        $groupId = Get-TenantReviewProperty -InputObject $group -Name 'Id'
        $displayName = Get-TenantReviewProperty -InputObject $group -Name 'DisplayName'
        $activity = $null
        if ($groupId -and $activityByTeamId.ContainsKey($groupId.ToString())) {
            $activity = $activityByTeamId[$groupId.ToString()]
        } elseif ($displayName -and $activityByName.ContainsKey($displayName.ToString().ToLowerInvariant())) {
            $activity = $activityByName[$displayName.ToString().ToLowerInvariant()]
        }

        $ownerCount = $null
        $memberCount = $null
        if ($IncludeOwnersAndMembers) {
            if (Get-Command -Name Invoke-TenantReviewGraphRestRequest -ErrorAction SilentlyContinue) {
                try {
                    $ownerCount = @(Invoke-TenantReviewGraphRestRequest -Uri "groups/$groupId/owners?`$select=id" -All).Count
                } catch {
                    $warnings += "Unable to expand owners for team '$displayName'. $($_.Exception.Message)"
                }
            } elseif (Get-Command -Name Get-MgGroupOwner -ErrorAction SilentlyContinue) {
                try {
                    $ownerCount = @(Get-MgGroupOwner -GroupId $groupId -All -ErrorAction Stop).Count
                } catch {
                    $warnings += "Unable to expand owners for team '$displayName'. $($_.Exception.Message)"
                }
            }
            if (Get-Command -Name Invoke-TenantReviewGraphRestRequest -ErrorAction SilentlyContinue) {
                try {
                    $memberCount = @(Invoke-TenantReviewGraphRestRequest -Uri "groups/$groupId/members?`$select=id" -All).Count
                } catch {
                    $warnings += "Unable to expand members for team '$displayName'. $($_.Exception.Message)"
                }
            } elseif (Get-Command -Name Get-MgGroupMember -ErrorAction SilentlyContinue) {
                try {
                    $memberCount = @(Get-MgGroupMember -GroupId $groupId -All -ErrorAction Stop).Count
                } catch {
                    $warnings += "Unable to expand members for team '$displayName'. $($_.Exception.Message)"
                }
            }
        }

        $isArchived = $null
        if ((Get-Command -Name Invoke-TenantReviewGraphRestRequest -ErrorAction SilentlyContinue) -and $groupId) {
            try {
                $team = Invoke-TenantReviewGraphRestRequest -Uri "teams/$groupId" -Raw
                $isArchived = Get-TenantReviewProperty -InputObject $team -Name 'IsArchived'
                if ($null -eq $isArchived) { $isArchived = Get-TenantReviewProperty -InputObject $team -Name 'isArchived' }
            } catch {
                $warnings += "Unable to collect team settings for '$displayName'. $($_.Exception.Message)"
            }
        } elseif ((Get-Command -Name Get-MgTeam -ErrorAction SilentlyContinue) -and $groupId) {
            try {
                $team = Get-MgTeam -TeamId $groupId -ErrorAction Stop
                $isArchived = Get-TenantReviewProperty -InputObject $team -Name 'IsArchived'
            } catch {
                $warnings += "Unable to collect team settings for '$displayName'. $($_.Exception.Message)"
            }
        }

        $lastActivityDate = ConvertTo-TenantReviewDateTime -Value (Get-TenantReviewProperty -InputObject $activity -Name 'Last Activity Date')
        $inactive = if ($null -ne $lastActivityDate) { $lastActivityDate -lt $staleCutoff } else { $null }

        $items += [pscustomobject]@{
            teamId              = $groupId
            displayName         = $displayName
            visibility          = Get-TenantReviewProperty -InputObject $group -Name 'Visibility'
            createdDateTime     = Get-TenantReviewProperty -InputObject $group -Name 'CreatedDateTime'
            isArchived          = $isArchived
            ownerCount          = $ownerCount
            memberCount         = $memberCount
            lastActivityDate    = if ($lastActivityDate) { $lastActivityDate.ToString('yyyy-MM-dd') } else { $null }
            channelMessages     = [int](ConvertTo-TenantReviewDecimal -Value (Get-TenantReviewProperty -InputObject $activity -Name 'Channel Messages') -Default 0)
            replyMessages       = [int](ConvertTo-TenantReviewDecimal -Value (Get-TenantReviewProperty -InputObject $activity -Name 'Reply Messages') -Default 0)
            meetingsOrganized   = [int](ConvertTo-TenantReviewDecimal -Value (Get-TenantReviewProperty -InputObject $activity -Name 'Meetings Organized Count') -Default 0)
            inactive            = $inactive
        }
    }

    $totalChannelMessages = Get-TenantReviewPropertySum -Items $items -Property 'channelMessages'
    $totalMeetings = Get-TenantReviewPropertySum -Items $items -Property 'meetingsOrganized'
    $result = [ordered]@{
        dataset     = 'TeamsInventory'
        generatedAt = (Get-Date).ToString('o')
        summary     = [pscustomobject]@{
            totalTeams              = $items.Count
            activeTeams             = @($items | Where-Object { $_.inactive -eq $false }).Count
            inactiveTeams           = @($items | Where-Object { $_.inactive -eq $true }).Count
            publicTeams             = @($items | Where-Object { $_.visibility -eq 'Public' }).Count
            privateTeams            = @($items | Where-Object { $_.visibility -eq 'Private' }).Count
            archivedTeams           = @($items | Where-Object { $_.isArchived -eq $true }).Count
            teamsWithNoOwners       = @($items | Where-Object { $null -ne $_.ownerCount -and $_.ownerCount -eq 0 }).Count
            teamsWithFewMembers     = @($items | Where-Object { $null -ne $_.memberCount -and $_.memberCount -lt 3 }).Count
            totalChannelMessages    = if ($null -ne $totalChannelMessages) { [int]$totalChannelMessages } else { $null }
            totalMeetings           = if ($null -ne $totalMeetings) { [int]$totalMeetings } else { $null }
            activeTeamsReportPeriod = "D$ReportPeriodDays"
            ownerMemberCountsIncluded = [bool]$IncludeOwnersAndMembers
        }
        items       = @($items)
        warnings    = @($warnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = $raw
    }

    [pscustomobject]$result
}
