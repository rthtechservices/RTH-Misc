[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

Write-Host "TenantReviewPack validation"
Write-Host "Project root: $projectRoot"

$scriptFiles = @(Get-ChildItem -Path $projectRoot -Recurse -Filter '*.ps1' -File)
foreach ($scriptFile in $scriptFiles) {
    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName, [ref]$tokens, [ref]$parseErrors) | Out-Null
    if ($parseErrors.Count -gt 0) {
        foreach ($parseError in $parseErrors) {
            $failures.Add(("Parser error in {0}: {1}" -f $scriptFile.FullName, $parseError.Message))
        }
    }
}

$safeDotSourceFiles = @(Get-ChildItem -Path (Join-Path $projectRoot 'src') -Recurse -Filter '*.ps1' -File)
foreach ($scriptFile in $safeDotSourceFiles) {
    try {
        . $scriptFile.FullName
    } catch {
        $failures.Add(("Failed to dot-source {0}: {1}" -f $scriptFile.FullName, $_.Exception.Message))
    }
}

$expectedFunctions = @(
    'Connect-TenantReviewServices',
    'Export-TenantReviewJson',
    'Get-LicenseInventory',
    'Get-UserInventory',
    'Get-LicenseUserAnalysis',
    'Invoke-AINarrative',
    'New-TenantReviewReport',
    'New-TenantReviewDeck'
)

foreach ($functionName in $expectedFunctions) {
    if (-not (Get-Command -Name $functionName -CommandType Function -ErrorAction SilentlyContinue)) {
        $failures.Add("Expected function was not loaded: $functionName")
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'Validation failed.' -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host " - $failure" -ForegroundColor Red
    }

    [pscustomobject]@{
        passed = $false
        checkedScripts = $scriptFiles.Count
        failures = @($failures)
    }
    exit 1
}

Write-Host 'Validation passed.' -ForegroundColor Green
[pscustomobject]@{
    passed = $true
    checkedScripts = $scriptFiles.Count
    functions = $expectedFunctions
}
