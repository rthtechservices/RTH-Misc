function Export-TenantReviewJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $InputObject |
        ConvertTo-Json -Depth 20 |
        Set-Content -Path $Path -Encoding UTF8

    Write-Host "Exported JSON: $Path"
}
