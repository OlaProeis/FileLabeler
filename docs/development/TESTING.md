# FileLabeler Test Suite

**Framework:** Pester 3.4.0+  
**Total Tests:** 106 (58 unit + 48 integration)  
**Pass Rate:** 100%  
**Execution Time:** ~15 seconds total

---

## Test Suites

### Unit Tests
**File:** `FileLabeler.Tests.ps1`  
**Tests:** 58  
**Focus:** Individual function testing with mock data  
**Execution Time:** ~6 seconds  
**Documentation:** [docs/features/unit-tests.md](../docs/features/unit-tests.md)

### Integration Tests ⭐ NEW
**File:** `FileLabeler.Integration.Tests.ps1`  
**Tests:** 48  
**Focus:** End-to-end workflows with real file operations  
**Execution Time:** ~9 seconds  
**Documentation:** [INTEGRATION_TESTS.md](INTEGRATION_TESTS.md)

---

## Running Tests

### Quick Start - All Tests
```powershell
# Run all tests (unit + integration)
Invoke-Pester -Path .\tests\
```

### Unit Tests Only
```powershell
# From project root
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1
```

### Integration Tests Only
```powershell
# Using convenience script
.\run_integration_tests.ps1

# Or directly with Pester
Invoke-Pester -Path .\tests\FileLabeler.Integration.Tests.ps1
```

### View Summary Only
```powershell
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1 -PassThru | Select-Object PassedCount, FailedCount, TotalCount, Time
```

### Run Specific Tests
```powershell
# Run only cache tests
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1 -TestName "*Label Cache*"

# Run only folder scanning tests
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1 -TestName "*Folder File Scanning*"
```

---

## Test Coverage

### Core Functions (48 tests)
1. **Label Cache Management** (5 tests)
   - Initialization, storage, retrieval, cache misses, clearing

2. **Folder File Scanning** (5 tests)
   - Recursive/non-recursive scanning, deduplication, empty folders

3. **File Type Filtering** (3 tests)
   - Support for 7 Office file types, filtering logic

4. **Duplicate Detection** (3 tests)
   - Detecting, filtering, and merging duplicate files

5. **Label Change Categorization** (5 tests)
   - New, Upgrade, Downgrade, Unchanged, Same label detection

6. **Warning Detection** (5 tests)
   - Mass downgrade, large batch, no changes, mixed changes

7. **Configuration Management** (6 tests)
   - Load, save, validate configuration structure

8. **Error Handling and Logging** (6 tests)
   - Structured logging, error patterns, exception handling

9. **Drag-and-Drop Support** (3 tests)
   - File/folder identification, mixed content processing

10. **UI Layout and Dynamic Sizing** (4 tests)
    - Height calculations, min/max row enforcement

11. **Norwegian Character Encoding** (3 tests)
    - UTF-8 preservation of æ, ø, å

12. **Statistics Tracking** (4 tests)
    - Counter initialization, increments, success rate calculation

### Integration Scenarios (6 tests)
- End-to-end workflow validation
- Edge cases (empty selections, large batches)

---

## Test Results

```
PassedCount : 58
FailedCount : 0
TotalCount  : 58
Time        : 00:00:06.28

Status: ✅ ALL TESTS PASSING
```

---

## Prerequisites

### Pester Module
Pester 3.4.0 is included with Windows PowerShell 5.1+:
```powershell
Get-Module -ListAvailable -Name Pester
```

### Test Environment
- Windows 10/11
- PowerShell 5.1 or later
- Write access to `$env:TEMP` for temporary test files

---

## Test Organization

```
tests/
├── FileLabeler.Tests.ps1              # Unit test suite (58 tests)
├── FileLabeler.Integration.Tests.ps1  # Integration test suite (48 tests) ⭐ NEW
├── INTEGRATION_TESTS.md               # Integration test documentation ⭐ NEW
├── IntegrationTestData/               # Auto-generated test files (cleaned after tests)
└── README.md                          # This file
```

**Root:**
```
run_integration_tests.ps1              # Integration test runner script ⭐ NEW
```

---

## Test Coverage

### Unit Tests (58 tests)
**Focus:** Individual functions with mock data

1. **Label Cache Management** (5 tests)
2. **Folder File Scanning** (5 tests)
3. **File Type Filtering** (3 tests)
4. **Duplicate Detection** (3 tests)
5. **Label Change Categorization** (5 tests)
6. **Warning Detection** (5 tests)
7. **Configuration Management** (6 tests)
8. **Error Handling and Logging** (6 tests)
9. **Drag-and-Drop Support** (3 tests)
10. **UI Layout and Dynamic Sizing** (4 tests)
11. **Norwegian Character Encoding** (3 tests)
12. **Statistics Tracking** (4 tests)
13. **Integration Scenarios** (6 tests)

### Integration Tests (48 tests) ⭐ NEW
**Focus:** End-to-end workflows with real file operations

1. **Test Framework Setup** (4 tests) - Subtask 17.1
2. **Complete Workflow Simulation** (5 tests) - Subtask 17.2
3. **Large Batch Processing** (4 tests) - Subtask 17.3
4. **Mixed Label Scenarios** (8 tests) - Subtask 17.4
5. **Protection Handling** (4 tests) - Subtask 17.5
6. **Error Recovery & Special Locations** (8 tests) - Subtask 17.6
7. **End-to-End Workflow** (3 tests)
8. **Real AIP Module Integration** (4 tests, skip if module unavailable)
9. **Performance Benchmarks** (2 tests)
10. **Cleanup Verification** (1 test)

---

## Documentation

### Unit Tests
[docs/features/unit-tests.md](../docs/features/unit-tests.md)
- Detailed test descriptions
- Implementation notes
- Key findings
- Performance metrics

### Integration Tests ⭐ NEW
[tests/INTEGRATION_TESTS.md](INTEGRATION_TESTS.md)
- Workflow testing
- Performance thresholds
- Real AIP integration
- Setup and execution guide

---

## Quick Reference

### Test Contexts
| Context | Tests | Coverage |
|---------|-------|----------|
| Label Cache Management | 5 | Cache operations |
| Folder File Scanning | 5 | Recursive scanning, deduplication |
| File Type Filtering | 3 | Supported extensions |
| Duplicate Detection | 3 | File merging |
| Label Change Categorization | 5 | Change analysis |
| Warning Detection | 5 | Smart warnings |
| Configuration Management | 6 | Config load/save/validate |
| Error Handling | 6 | Logging, error patterns |
| Drag-and-Drop | 3 | Path processing |
| UI Layout | 4 | Dynamic sizing |
| Norwegian Encoding | 3 | UTF-8 characters |
| Statistics Tracking | 4 | Counter operations |
| Integration | 2 | Workflow validation |
| Edge Cases | 4 | Error scenarios |

---

## Troubleshooting

### Tests Fail with "Should operator not found"
You may be using Pester 5.x syntax. This test suite requires Pester 3.x:
```powershell
# Check Pester version
Get-Module -ListAvailable -Name Pester | Select-Object Version

# If you have Pester 5.x, use:
Import-Module Pester -MaximumVersion 3.99
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1
```

### Temporary Files Not Cleaned Up
Tests clean up after themselves, but if interrupted:
```powershell
# Manually clean temp folders
Get-ChildItem $env:TEMP -Filter "PesterTest_*" -Directory | Remove-Item -Recurse -Force
```

---

**Maintained by:** FileLabeler Development Team  
**Last Updated:** 2025-10-26  
**Test Framework:** Pester 3.4.0

