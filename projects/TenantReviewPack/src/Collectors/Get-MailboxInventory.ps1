function Get-MailboxInventory {
    [CmdletBinding()]
    param(
        [switch]$IncludeRaw,

        [switch]$IncludeInboxRules,

        [int]$InboxRuleMailboxLimit = 200,

        [switch]$IncludeMailboxStatistics
    )

    function Test-ExternalMailboxForwarding {
        param(
            [object]$ForwardingAddress,
            [object]$ForwardingSmtpAddress,
            [object]$PrimarySmtpAddress
        )

        $primaryText = if ($PrimarySmtpAddress) { $PrimarySmtpAddress.ToString() } else { '' }
        $primaryDomain = if ($primaryText -match '@(.+)$') { $Matches[1].ToLowerInvariant() } else { $null }
        $target = if ($ForwardingSmtpAddress) { $ForwardingSmtpAddress.ToString() } elseif ($ForwardingAddress) { $ForwardingAddress.ToString() } else { $null }
        if (-not $target) {
            return $false
        }
        if ($target -notmatch '@') {
            return $false
        }
        if (-not $primaryDomain) {
            return $true
        }

        return -not $target.ToLowerInvariant().EndsWith("@$primaryDomain")
    }

    $warnings = @()
    $mailboxes = @()
    $mailboxStats = @{}
    $transportRules = @()
    $inboxForwardingRules = @()
    $inboxRulesScanned = 0

    $mailboxCommand = Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue
    if (-not $mailboxCommand) {
        $mailboxCommand = Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue
    }

    if (-not $mailboxCommand) {
        $warnings += 'Exchange Online mailbox commands are not available or not connected. Install/connect ExchangeOnlineManagement for mailbox inventory.'
        return [pscustomobject]@{
            dataset        = 'MailboxInventory'
            generatedAt    = (Get-Date).ToString('o')
            summary        = [pscustomobject]@{
                totalMailboxes                 = 0
                userMailboxes                  = 0
                sharedMailboxes                = 0
                roomMailboxes                  = 0
                equipmentMailboxes             = 0
                mailboxesWithForwarding        = 0
                mailboxesForwardingExternally  = 0
                transportRules                 = 0
                enabledTransportRules          = 0
                inboxRulesScanned              = 0
                inboxForwardingRulesFound      = 0
                totalMailboxSizeGB             = $null
                largestMailboxes               = @()
            }
            items          = @()
            transportRules = @()
            warnings       = @($warnings)
        }
    }

    try {
        $parameters = @{}
        if ($mailboxCommand.Parameters.ContainsKey('ResultSize')) {
            $parameters.ResultSize = 'Unlimited'
        }
        if ($mailboxCommand.Parameters.ContainsKey('Properties')) {
            $parameters.Properties = @('ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward')
        }
        $mailboxes = @(& $mailboxCommand @parameters -ErrorAction Stop)
    } catch {
        $warnings += "Unable to collect Exchange mailboxes. $($_.Exception.Message)"
    }

    if ($IncludeMailboxStatistics) {
        $statsCommand = Get-Command -Name Get-EXOMailboxStatistics -ErrorAction SilentlyContinue
        if (-not $statsCommand) {
            $statsCommand = Get-Command -Name Get-MailboxStatistics -ErrorAction SilentlyContinue
        }
        if ($statsCommand) {
            foreach ($mailbox in $mailboxes) {
                $identity = Get-TenantReviewProperty -InputObject $mailbox -Name 'ExternalDirectoryObjectId'
                if (-not $identity) { $identity = Get-TenantReviewProperty -InputObject $mailbox -Name 'Identity' }
                if (-not $identity) { continue }
                try {
                    $stat = & $statsCommand -Identity $identity -ErrorAction Stop
                    $mailboxStats[$identity.ToString()] = $stat
                } catch {
                    $warnings += "Mailbox statistics unavailable for one mailbox. $($_.Exception.Message)"
                }
            }
        } else {
            $warnings += 'Mailbox statistics command is not available; mailbox size fields were skipped.'
        }
    }

    $transportCommand = Get-Command -Name Get-TransportRule -ErrorAction SilentlyContinue
    if ($transportCommand) {
        try {
            $transportRules = @(& $transportCommand -ErrorAction Stop)
        } catch {
            $warnings += "Unable to collect transport rules. $($_.Exception.Message)"
        }
    } else {
        $warnings += 'Get-TransportRule is not available; transport rule inventory was skipped.'
    }

    $items = @()
    foreach ($mailbox in $mailboxes) {
        $recipientTypeDetails = (Get-TenantReviewProperty -InputObject $mailbox -Name 'RecipientTypeDetails')
        $primarySmtpAddress = Get-TenantReviewProperty -InputObject $mailbox -Name 'PrimarySmtpAddress'
        $forwardingAddress = Get-TenantReviewProperty -InputObject $mailbox -Name 'ForwardingAddress'
        $forwardingSmtpAddress = Get-TenantReviewProperty -InputObject $mailbox -Name 'ForwardingSmtpAddress'
        $deliverToMailboxAndForward = Get-TenantReviewProperty -InputObject $mailbox -Name 'DeliverToMailboxAndForward'
        $forwardingEnabled = [bool]($forwardingAddress -or $forwardingSmtpAddress)
        $identity = Get-TenantReviewProperty -InputObject $mailbox -Name 'ExternalDirectoryObjectId'
        if (-not $identity) { $identity = Get-TenantReviewProperty -InputObject $mailbox -Name 'Identity' }
        $stat = if ($identity -and $mailboxStats.ContainsKey($identity.ToString())) { $mailboxStats[$identity.ToString()] } else { $null }
        $totalItemSizeGB = ConvertTo-TenantReviewGB -Value (Get-TenantReviewProperty -InputObject $stat -Name 'TotalItemSize')

        $items += [pscustomobject]@{
            displayName                 = Get-TenantReviewProperty -InputObject $mailbox -Name 'DisplayName'
            primarySmtpAddress          = if ($primarySmtpAddress) { $primarySmtpAddress.ToString() } else { $null }
            recipientTypeDetails        = if ($recipientTypeDetails) { $recipientTypeDetails.ToString() } else { $null }
            isShared                    = ($recipientTypeDetails -eq 'SharedMailbox')
            forwardingEnabled           = $forwardingEnabled
            forwardingAddress           = if ($forwardingAddress) { $forwardingAddress.ToString() } else { $null }
            forwardingSmtpAddress       = if ($forwardingSmtpAddress) { $forwardingSmtpAddress.ToString() } else { $null }
            deliverToMailboxAndForward  = $deliverToMailboxAndForward
            externalForwardingSuspected = Test-ExternalMailboxForwarding -ForwardingAddress $forwardingAddress -ForwardingSmtpAddress $forwardingSmtpAddress -PrimarySmtpAddress $primarySmtpAddress
            totalItemSizeGB             = $totalItemSizeGB
            itemCount                   = Get-TenantReviewProperty -InputObject $stat -Name 'ItemCount'
        }
    }

    if ($IncludeInboxRules) {
        $inboxRuleCommand = Get-Command -Name Get-InboxRule -ErrorAction SilentlyContinue
        if ($inboxRuleCommand) {
            foreach ($mailbox in ($mailboxes | Select-Object -First $InboxRuleMailboxLimit)) {
                $identity = Get-TenantReviewProperty -InputObject $mailbox -Name 'PrimarySmtpAddress'
                if (-not $identity) { $identity = Get-TenantReviewProperty -InputObject $mailbox -Name 'Identity' }
                if (-not $identity) { continue }
                try {
                    $rules = @(& $inboxRuleCommand -Mailbox $identity -ErrorAction Stop)
                    $inboxRulesScanned++
                    foreach ($rule in $rules) {
                        $forwardTo = @(Get-TenantReviewProperty -InputObject $rule -Name 'ForwardTo')
                        $redirectTo = @(Get-TenantReviewProperty -InputObject $rule -Name 'RedirectTo')
                        $forwardAsAttachmentTo = @(Get-TenantReviewProperty -InputObject $rule -Name 'ForwardAsAttachmentTo')
                        if ($forwardTo.Count -gt 0 -or $redirectTo.Count -gt 0 -or $forwardAsAttachmentTo.Count -gt 0) {
                            $inboxForwardingRules += [pscustomobject]@{
                                mailbox               = $identity.ToString()
                                name                  = Get-TenantReviewProperty -InputObject $rule -Name 'Name'
                                enabled               = Get-TenantReviewProperty -InputObject $rule -Name 'Enabled'
                                forwardTo             = @($forwardTo | ForEach-Object { $_.ToString() })
                                redirectTo            = @($redirectTo | ForEach-Object { $_.ToString() })
                                forwardAsAttachmentTo = @($forwardAsAttachmentTo | ForEach-Object { $_.ToString() })
                            }
                        }
                    }
                } catch {
                    $warnings += "Unable to scan inbox rules for one mailbox. $($_.Exception.Message)"
                }
            }
        } else {
            $warnings += 'Get-InboxRule is not available; inbox rule scan was skipped.'
        }
    } else {
        $warnings += 'Inbox rule scan skipped. Use -IncludeInboxRules to scan per-mailbox forwarding rules.'
    }

    $transportItems = @()
    foreach ($rule in $transportRules) {
        $transportItems += [pscustomobject]@{
            name               = Get-TenantReviewProperty -InputObject $rule -Name 'Name'
            state              = Get-TenantReviewProperty -InputObject $rule -Name 'State'
            mode               = Get-TenantReviewProperty -InputObject $rule -Name 'Mode'
            priority           = Get-TenantReviewProperty -InputObject $rule -Name 'Priority'
            comments           = Get-TenantReviewProperty -InputObject $rule -Name 'Comments'
            conditionsSummary  = ((Get-TenantReviewProperty -InputObject $rule -Name 'Conditions') -join '; ')
            actionsSummary     = ((Get-TenantReviewProperty -InputObject $rule -Name 'Actions') -join '; ')
        }
    }

    $largestMailboxes = @($items | Where-Object { $null -ne $_.totalItemSizeGB } | Sort-Object -Property totalItemSizeGB -Descending | Select-Object -First 10)
    $totalMailboxSize = ($items | Where-Object { $null -ne $_.totalItemSizeGB } | Measure-Object -Property totalItemSizeGB -Sum).Sum
    if ($null -ne $totalMailboxSize) {
        $totalMailboxSize = [decimal]::Round([decimal]$totalMailboxSize, 2)
    }

    $result = [ordered]@{
        dataset              = 'MailboxInventory'
        generatedAt          = (Get-Date).ToString('o')
        summary              = [pscustomobject]@{
            totalMailboxes                = $items.Count
            userMailboxes                 = @($items | Where-Object { $_.recipientTypeDetails -eq 'UserMailbox' }).Count
            sharedMailboxes               = @($items | Where-Object { $_.recipientTypeDetails -eq 'SharedMailbox' }).Count
            roomMailboxes                 = @($items | Where-Object { $_.recipientTypeDetails -eq 'RoomMailbox' }).Count
            equipmentMailboxes            = @($items | Where-Object { $_.recipientTypeDetails -eq 'EquipmentMailbox' }).Count
            mailboxesWithForwarding       = @($items | Where-Object { $_.forwardingEnabled }).Count
            mailboxesForwardingExternally = @($items | Where-Object { $_.externalForwardingSuspected }).Count
            transportRules                = $transportItems.Count
            enabledTransportRules         = @($transportItems | Where-Object { $_.state -eq 'Enabled' }).Count
            inboxRulesScanned             = $inboxRulesScanned
            inboxForwardingRulesFound     = $inboxForwardingRules.Count
            totalMailboxSizeGB            = $totalMailboxSize
            largestMailboxes              = @($largestMailboxes)
        }
        items                = @($items)
        transportRules       = @($transportItems)
        inboxForwardingRules = @($inboxForwardingRules)
        warnings             = @($warnings)
    }

    if ($IncludeRaw) {
        $result['raw'] = [pscustomobject]@{
            mailboxes      = @($mailboxes)
            transportRules = @($transportRules)
        }
    }

    [pscustomobject]$result
}
