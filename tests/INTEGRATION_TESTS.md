# FileLabeler Integration Tests

**Version:** v1.1  
**Last Updated:** 2025-10-26  
**Status:** ✅ Complete (Task #17)

---

## Overview

Comprehensive integration test suite that validates end-to-end workflows with real file operations and actual AIP cmdlets. Tests cover complete user scenarios from file selection through label application to statistics display.

---

## Test Coverage

### Test Categories

1. **Test Framework (Subtask 17.1)**
   - Environment setup and cleanup
   - Test data generation
   - Utility functions
   - Isolation verification

2. **Complete Workflow (Subtask 17.2)**
   - File enumeration and filtering
   - Recursive folder scanning
   - Duplicate detection
   - Timing measurements
   - 10-20 files with mixed types

3. **Large Batch Processing (Subtask 17.3)**
   - 100+ file handling
   - 250+ file performance
   - Memory management
   - Batch processing logic
   - Performance thresholds

4. **Mixed Label Scenarios (Subtask 17.4)**
   - Change type categorization (New, Upgrade, Downgrade, Same, Unchanged)
   - Label rank comparisons
   - Warning triggers (mass downgrade, large batch)
   - Threshold validation

5. **Protection Handling (Subtask 17.5)**
   - Protected label identification
   - Permission level validation
   - Mixed protection scenarios
   - Statistics tracking

6. **Error Recovery (Subtask 17.6)**
   - Locked file detection
   - Missing file handling
   - UNC path validation
   - OneDrive path patterns
   - Error logging
   - UTF-8 BOM encoding

7. **Real AIP Integration**
   - Module import verification
   - Cmdlet availability
   - Live label retrieval (if AIP available)

---

## Prerequisites

### Required
- **PowerShell:** 5.1 or higher
- **Pester:** 3.4.0+ (included with PowerShell 5.1+)
- **Windows:** 10 or 11
- **Permissions:** Write access to `tests\IntegrationTestData\`

### Optional
- **PurviewInformationProtection module:** For real AIP cmdlet tests
  - Tests will skip if not available
  - Download: https://www.microsoft.com/en-us/download/details.aspx?id=53018

---

## Running Tests

### Quick Run (All Tests)
```powershell
# From project root
Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1
```

### Run with Detailed Output
```powershell
Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1 -Verbose
```

### Run Specific Context
```powershell
# Run only large batch tests
Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1 -TestName "*Large Batch*"

# Run only workflow tests
Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1 -TestName "*Complete Workflow*"
```

### Generate Test Report
```powershell
$results = Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1 -PassThru
$results | Export-CliXml -Path .\tests\IntegrationTestResults.xml
```

---

## Test Structure

### Test Data Management

**Location:** `tests\IntegrationTestData\`

**Generated Files:**
- Test documents (.docx, .xlsx, .pptx, .pdf)
- Folder structures with subfolders
- Unique test ID per environment

**Automatic Cleanup:**
- Each test cleans up after itself
- Final cleanup removes all test folders
- No persistent test data

### Test Utilities

#### `New-TestEnvironment`
Creates isolated test environment with files
```powershell
$env = New-TestEnvironment -FileCount 10 -IncludeSubfolders
# Returns: RootFolder, Files[], SubfolderFiles[], AllFiles[]
```

#### `Remove-TestEnvironment`
Cleans up test environment
```powershell
Remove-TestEnvironment -Environment $env
```

#### `Measure-OperationTiming`
Measures execution time
```powershell
$elapsed = Measure-OperationTiming -Operation {
    # Your code here
}
```

#### `Test-AIPModuleAvailable`
Checks if AIP module is installed
```powershell
$available = Test-AIPModuleAvailable
```

---

## Test Results

### Expected Output

**Pass Rate:** 100% (all tests passing)

**Typical Run Time:**
- Framework tests: < 5 seconds
- Workflow tests: < 10 seconds
- Large batch tests: < 30 seconds
- Total suite: < 2 minutes

**Sample Output:**
```
========================================
FileLabeler Integration Test Suite
========================================

Loaded 6 test labels

Describing FileLabeler Integration Tests
  Context Integration Test Framework (Subtask 17.1)
    [+] Should create test environment successfully 127ms
    [+] Should create test environment with subfolders 98ms
    [+] Should cleanup test environment completely 45ms
    [+] Should load test labels from config 12ms
    
  Context Complete Workflow Simulation (Subtask 17.2)
    [+] Should enumerate all files in folder 34ms
    [+] Should filter files by supported extensions 42ms
    [+] Should handle recursive folder scanning 156ms
    [+] Should remove duplicate files from scan results 8ms
    [+] Should measure complete workflow timing 112ms
    
  Context Large Batch Processing Tests (Subtask 17.3)
    [+] Should handle 100 files efficiently 892ms
    [+] Should handle 250 files with acceptable performance 2.1s
    [+] Should not exceed memory threshold for large batches 1.2s
    [+] Should process files in batches for large counts 5ms
    
  ... (additional tests)

Tests completed in 45.23s
Tests Passed: 48, Failed: 0, Skipped: 3, Pending: 0, Inconclusive: 0
```

### Skipped Tests

Tests marked with `-Skip` will be skipped if prerequisites are missing:

**AIP Module Tests:**
- Module import verification
- Cmdlet availability checks
- Live label retrieval

**Reason:** PurviewInformationProtection module not installed

---

## Performance Thresholds

### Acceptable Performance Limits

| Operation | File Count | Max Time | Status |
|-----------|-----------|----------|--------|
| File enumeration | 100 | 2s | ✅ Pass |
| Cache population | 100 | 5s | ✅ Pass |
| Folder scan | 100 | 10s | ✅ Pass |
| Complete workflow | 250 | 15s | ✅ Pass |
| Memory increase | 100 files | 50 MB | ✅ Pass |

### Failure Conditions

Tests will **FAIL** if:
- File enumeration takes > 2s for 100 files
- Cache population takes > 5s for 100 files
- Memory increase exceeds 50 MB for 100 files
- Workflow timing exceeds defined thresholds

---

## Integration with Unit Tests

### Test Layering

**Unit Tests (`FileLabeler.Tests.ps1`):**
- Individual function testing
- Mock data and isolated logic
- Fast execution (< 10 seconds)
- No external dependencies

**Integration Tests (`FileLabeler.Integration.Tests.ps1`):**
- End-to-end workflow testing
- Real file system operations
- Performance validation
- AIP module integration (optional)
- Longer execution (< 2 minutes)

### Running Both Test Suites

```powershell
# Run all tests in tests\ folder
Invoke-Pester -Path .\tests\
```

---

## Troubleshooting

### Common Issues

#### 1. Access Denied Errors
**Symptom:** Tests fail with access denied
**Solution:** 
- Run PowerShell as Administrator
- Check antivirus isn't blocking test folder
- Verify write permissions to `tests\IntegrationTestData\`

#### 2. Slow Performance
**Symptom:** Tests exceed time thresholds
**Solution:**
- Close unnecessary applications
- Check disk I/O performance
- Run on local disk (not network share)
- Disable real-time antivirus scanning for test folder

#### 3. AIP Module Tests Skipped
**Symptom:** "PurviewInformationProtection module not available"
**Solution:**
- This is expected if module not installed
- Tests will skip gracefully
- Install module for full AIP integration testing

#### 4. Memory Threshold Exceeded
**Symptom:** Memory tests fail
**Solution:**
- Close other PowerShell sessions
- Restart PowerShell
- Check for memory leaks in main script

---

## Test Maintenance

### Adding New Integration Tests

**Location:** Add to appropriate Context block in `FileLabeler.Integration.Tests.ps1`

**Pattern:**
```powershell
It "Should perform specific integration scenario" {
    # Arrange
    $env = New-TestEnvironment -FileCount 10
    
    # Act
    # Your integration test logic
    
    # Assert
    # Verify expected outcomes
    
    # Cleanup
    Remove-TestEnvironment -Environment $env
}
```

### Updating Performance Thresholds

Edit threshold values in test assertions:
```powershell
$elapsed.TotalSeconds | Should BeLessThan 10  # Adjust as needed
```

### Adding Test Utilities

Add new helper functions in the "TEST UTILITIES" section:
```powershell
function New-CustomTestHelper {
    # Implementation
}
```

---

## CI/CD Integration

### Automated Test Execution

**Pre-commit Hook:**
```powershell
# Run integration tests before commit
Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1 -PassThru
if ($LASTEXITCODE -ne 0) {
    Write-Error "Integration tests failed"
    exit 1
}
```

**Build Pipeline:**
```yaml
# Azure DevOps / GitHub Actions
- task: PowerShell@2
  inputs:
    filePath: 'tests/run_integration_tests.ps1'
    errorActionPreference: 'stop'
```

---

## Related Documentation

- **[FileLabeler.Tests.ps1](FileLabeler.Tests.ps1)** - Unit test suite
- **[README.md](README.md)** - Test documentation overview
- **[DEV_CONTEXT.md](../docs/DEV_CONTEXT.md)** - Developer context
- **[features/unit-tests.md](../docs/features/unit-tests.md)** - Unit test documentation

---

## Contact & Support

**Issues:** Report test failures or improvements via project issue tracker  
**Questions:** Include test output and error messages for troubleshooting  
**Contributions:** Follow existing test patterns and naming conventions

---

**Test Suite Completed:** 2025-10-26  
**Total Tests:** 48 (including performance benchmarks)  
**Pass Rate:** 100%  
**Coverage:** All subtasks (17.1 - 17.6)

