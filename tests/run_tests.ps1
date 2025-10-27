# FileLabeler Unit Test Runner
# Quick script to run Pester tests with formatted output

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " FileLabeler Unit Test Suite" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if Pester is available
$pesterModule = Get-Module -ListAvailable -Name Pester
if (-not $pesterModule) {
    Write-Host "ERROR: Pester module not found!" -ForegroundColor Red
    Write-Host "Pester is required to run unit tests." -ForegroundColor Yellow
    exit 1
}

Write-Host "Pester Version: $($pesterModule[0].Version)" -ForegroundColor Gray
Write-Host ""

# Run tests
Write-Host "Running tests..." -ForegroundColor Yellow
Write-Host ""

$testResults = Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1 -PassThru

# Display summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Test Results Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$passColor = if ($testResults.FailedCount -eq 0) { "Green" } else { "Yellow" }
$failColor = if ($testResults.FailedCount -gt 0) { "Red" } else { "Gray" }

Write-Host "Passed:  " -NoNewline
Write-Host "$($testResults.PassedCount)" -ForegroundColor $passColor

Write-Host "Failed:  " -NoNewline
Write-Host "$($testResults.FailedCount)" -ForegroundColor $failColor

Write-Host "Total:   $($testResults.TotalCount)" -ForegroundColor White

Write-Host "Time:    $($testResults.Time.TotalSeconds) seconds" -ForegroundColor Gray

Write-Host ""

if ($testResults.FailedCount -eq 0) {
    Write-Host "SUCCESS: All tests passed!" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAILURE: Some tests failed." -ForegroundColor Red
    exit 1
}

