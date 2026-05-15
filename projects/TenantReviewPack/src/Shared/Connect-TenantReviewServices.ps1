function Connect-TenantReviewServices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Settings
    )

    Write-Host 'Connection bootstrap placeholder.' -ForegroundColor Yellow
    Write-Host 'Next step: connect Microsoft Graph, Exchange Online, SharePoint/PnP, and Teams reporting modules.'

    [pscustomobject]@{
        GraphConnected          = $false
        ExchangeOnlineConnected = $false
        SharePointConnected     = $false
        TeamsConnected          = $false
    }
}
