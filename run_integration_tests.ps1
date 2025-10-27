#Requires -Version 5.1
<#
.SYNOPSIS
    Integration Test Runner for FileLabeler
.DESCRIPTION
    Convenient script to run integration tests with various options
    Provides formatted output and test result summaries
.EXAMPLE
    .\run_integration_tests.ps1
    Runs all integration tests
.EXAMPLE
    .\run_integration_tests.ps1 -Workflow
    Runs only workflow tests
.EXAMPLE
    .\run_integration_tests.ps1 -Detailed
    Runs tests with detailed output
#>

param(
    [switch]$Framework,      # Run only framework tests
    [switch]$Workflow,       # Run only workflow tests
    [switch]$LargeBatch,     # Run only large batch tests
    [switch]$MixedLabels,    # Run only mixed label tests
    [switch]$Protection,     # Run only protection tests
    [switch]$ErrorRecovery,  # Run only error recovery tests
    [switch]$Performance,    # Run only performance tests
    [switch]$Detailed,       # Show detailed output
    [switch]$ExportResults   # Export results to XML
)

# Script location
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$integrationTestPath = Join-Path $scriptRoot "FileLabeler.Integration.Tests.ps1"

# Check if test file exists
if (-not (Test-Path $integrationTestPath)) {
    Write-Host "ERROR: Integration test file not found at: $integrationTestPath" -ForegroundColor Red
    exit 1
}

# Display header
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "FileLabeler Integration Test Runner" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Import Pester
try {
    Import-Module Pester -ErrorAction Stop
    Write-Host "Pester module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Could not load Pester module - $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Build test parameters
$pesterParams = @{
    Path = $integrationTestPath
}

# Add verbosity if requested
if ($Detailed) {
    $pesterParams.Verbose = $true
}

# Add test name filter based on switches
if ($Framework) {
    $pesterParams.TestName = "*Test Framework*"
    Write-Host "Running: Framework Tests Only`n" -ForegroundColor Yellow
}
elseif ($Workflow) {
    $pesterParams.TestName = "*Complete Workflow*"
    Write-Host "Running: Workflow Tests Only`n" -ForegroundColor Yellow
}
elseif ($LargeBatch) {
    $pesterParams.TestName = "*Large Batch*"
    Write-Host "Running: Large Batch Tests Only`n" -ForegroundColor Yellow
}
elseif ($MixedLabels) {
    $pesterParams.TestName = "*Mixed Label*"
    Write-Host "Running: Mixed Label Tests Only`n" -ForegroundColor Yellow
}
elseif ($Protection) {
    $pesterParams.TestName = "*Protection*"
    Write-Host "Running: Protection Tests Only`n" -ForegroundColor Yellow
}
elseif ($ErrorRecovery) {
    $pesterParams.TestName = "*Error Recovery*"
    Write-Host "Running: Error Recovery Tests Only`n" -ForegroundColor Yellow
}
elseif ($Performance) {
    $pesterParams.TestName = "*Performance*"
    Write-Host "Running: Performance Benchmarks Only`n" -ForegroundColor Yellow
}
else {
    Write-Host "Running: All Integration Tests`n" -ForegroundColor Yellow
}

# Add PassThru for result analysis
$pesterParams.PassThru = $true

# Run tests
try {
    $startTime = Get-Date
    Write-Host "Test execution started at: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
    Write-Host ""
    
    $results = Invoke-Pester @pesterParams
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    # Display summary
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Test Results Summary" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    Write-Host "Total Tests:    " -NoNewline
    Write-Host $results.TotalCount -ForegroundColor White
    
    Write-Host "Passed:         " -NoNewline
    Write-Host $results.PassedCount -ForegroundColor Green
    
    Write-Host "Failed:         " -NoNewline
    if ($results.FailedCount -gt 0) {
        Write-Host $results.FailedCount -ForegroundColor Red
    } else {
        Write-Host $results.FailedCount -ForegroundColor Gray
    }
    
    Write-Host "Skipped:        " -NoNewline
    Write-Host $results.SkippedCount -ForegroundColor Yellow
    
    Write-Host "Duration:       " -NoNewline
    Write-Host "$([Math]::Round($duration.TotalSeconds, 2))s" -ForegroundColor White
    
    Write-Host "Success Rate:   " -NoNewline
    if ($results.TotalCount -gt 0) {
        $successRate = [Math]::Round(($results.PassedCount / $results.TotalCount) * 100, 1)
        if ($successRate -eq 100) {
            Write-Host "$successRate%" -ForegroundColor Green
        } elseif ($successRate -ge 90) {
            Write-Host "$successRate%" -ForegroundColor Yellow
        } else {
            Write-Host "$successRate%" -ForegroundColor Red
        }
    } else {
        Write-Host "N/A" -ForegroundColor Gray
    }
    
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Show failed test details
    if ($results.FailedCount -gt 0) {
        Write-Host "Failed Tests:" -ForegroundColor Red
        foreach ($test in $results.TestResult | Where-Object { $_.Result -eq 'Failed' }) {
            Write-Host "  - $($test.Name)" -ForegroundColor Red
            Write-Host "    Error: $($test.FailureMessage)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Show skipped test summary
    if ($results.SkippedCount -gt 0) {
        Write-Host "Skipped Tests: $($results.SkippedCount)" -ForegroundColor Yellow
        Write-Host "  (Likely due to missing PurviewInformationProtection module)`n" -ForegroundColor Gray
    }
    
    # Export results if requested
    if ($ExportResults) {
        $resultsPath = Join-Path $scriptRoot "IntegrationTestResults.xml"
        $results | Export-CliXml -Path $resultsPath
        Write-Host "Test results exported to: $resultsPath" -ForegroundColor Green
        Write-Host ""
    }
    
    # Exit with appropriate code
    if ($results.FailedCount -gt 0) {
        Write-Host "INTEGRATION TESTS FAILED" -ForegroundColor Red
        exit 1
    } else {
        Write-Host "ALL INTEGRATION TESTS PASSED" -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Host "`nERROR: Test execution failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}

