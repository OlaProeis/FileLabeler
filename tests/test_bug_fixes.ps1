#Requires -Version 5.1
<#
.SYNOPSIS
    Test script for critical bug fixes in FileLabeler
.DESCRIPTION
    Creates test files to verify bug fixes:
    - BUG #1: Memory crashes (test with bulk operations)
    - BUG #2: Square brackets in filenames
    - BUG #3: File lock detection
.NOTES
    Run this script to prepare test environment before manual testing
#>

param(
    [string]$TestFolder = "$env:TEMP\FileLabeler_BugTest"
)

Write-Host "==================================" -ForegroundColor Cyan
Write-Host "FileLabeler Bug Fix Test Setup" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""

# Create test folder
if (Test-Path $TestFolder) {
    Write-Host "Cleaning existing test folder..." -ForegroundColor Yellow
    Remove-Item -Path $TestFolder -Recurse -Force
}

New-Item -Path $TestFolder -ItemType Directory -Force | Out-Null
Write-Host "✓ Created test folder: $TestFolder" -ForegroundColor Green
Write-Host ""

# ========================================
# BUG #2 TEST: Square Brackets in Filenames
# ========================================
Write-Host "Creating files with square brackets..." -ForegroundColor Yellow

$bracketFiles = @(
    "Report[2024].xlsx",
    "File[1].docx",
    "Data[Final].pdf",
    "Budget [Q1].xlsx",
    "Project[A-B].pptx",
    "Test [Multiple] [Brackets].xlsx"
)

foreach ($file in $bracketFiles) {
    $filePath = Join-Path $TestFolder $file
    
    # Create appropriate file based on extension
    switch ([System.IO.Path]::GetExtension($file)) {
        ".xlsx" {
            # Create minimal Excel file
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Cells.Item(1, 1) = "Test Data - $file"
            $workbook.SaveAs($filePath)
            $workbook.Close()
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        ".docx" {
            # Create minimal Word file
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Add()
            $doc.Content.Text = "Test Document - $file"
            $doc.SaveAs($filePath)
            $doc.Close()
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        ".pptx" {
            # Create minimal PowerPoint file
            $ppt = New-Object -ComObject PowerPoint.Application
            $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
            $presentation = $ppt.Presentations.Add()
            $slide = $presentation.Slides.Add(1, 12) # ppLayoutText
            $slide.Shapes.Item(1).TextFrame.TextRange.Text = "Test - $file"
            $presentation.SaveAs($filePath)
            $presentation.Close()
            $ppt.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        }
        ".pdf" {
            # Create placeholder PDF (requires actual PDF library for real PDF)
            # For now, create a text file with PDF extension
            "PDF Test File - $file" | Out-File $filePath -Encoding UTF8
        }
    }
    
    Write-Host "  ✓ Created: $file" -ForegroundColor Green
}

Write-Host ""

# ========================================
# BUG #1 TEST: Bulk Operations (100+ files)
# ========================================
Write-Host "Creating bulk test files (50 files)..." -ForegroundColor Yellow

$bulkFolder = Join-Path $TestFolder "BulkTest"
New-Item -Path $bulkFolder -ItemType Directory -Force | Out-Null

for ($i = 1; $i -le 50; $i++) {
    $fileName = "BulkTest_$($i.ToString('000')).xlsx"
    $filePath = Join-Path $bulkFolder $fileName
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Item(1, 1) = "Bulk Test File #$i"
        $workbook.SaveAs($filePath)
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    catch {
        Write-Host "  ⚠ Could not create Excel file #$i (Office not installed?)" -ForegroundColor Yellow
    }
}

Write-Host "  ✓ Created 50 bulk test files" -ForegroundColor Green
Write-Host ""

# ========================================
# BUG #3 TEST: File Lock Detection
# ========================================
Write-Host "Creating file for lock test..." -ForegroundColor Yellow

$lockTestFile = Join-Path $TestFolder "LockedFile_OpenMe.xlsx"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Cells.Item(1, 1) = "LOCK TEST - Open this file in Excel before running FileLabeler"
    $worksheet.Cells.Item(2, 1) = "This file should be SKIPPED when locked"
    $worksheet.Cells.Item(3, 1) = "You should see: 'File is locked by another process'"
    $workbook.SaveAs($lockTestFile)
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "  ✓ Created: LockedFile_OpenMe.xlsx" -ForegroundColor Green
}
catch {
    Write-Host "  ⚠ Could not create lock test file (Office not installed?)" -ForegroundColor Yellow
}

Write-Host ""

# ========================================
# CREATE README
# ========================================
$readmePath = Join-Path $TestFolder "README.txt"
$readmeContent = @"
FileLabeler Bug Fix Test Files
================================

This folder contains test files to verify bug fixes:

BUG #2 TEST: Square Brackets in Filenames
------------------------------------------
Files with brackets in their names:
- Report[2024].xlsx
- File[1].docx
- Data[Final].pdf
- Budget [Q1].xlsx
- Project[A-B].pptx
- Test [Multiple] [Brackets].xlsx

TEST: Apply any label to these files in FileLabeler
EXPECTED: All files should be labeled successfully (no crashes)


BUG #1 TEST: Bulk Operations
-----------------------------
Folder: BulkTest\
Contains: 50 Excel files

TEST: Apply label to all 50 files at once
EXPECTED: All files processed without memory crashes


BUG #3 TEST: File Lock Detection
---------------------------------
File: LockedFile_OpenMe.xlsx

TEST STEPS:
1. Open LockedFile_OpenMe.xlsx in Excel
2. Run FileLabeler and try to apply label to this file
3. Check that file is SKIPPED with message: "File is locked by another process"
4. Close Excel
5. Retry labeling - should succeed now


MANUAL TESTING CHECKLIST
=========================
[ ] Square brackets test - all files labeled successfully
[ ] Bulk test - 50+ files without crashes
[ ] Lock test - locked file skipped gracefully
[ ] Lock test - file labeled after closing Excel
[ ] Check statistics dialog shows correct counts
[ ] Check log file for detailed operation records


Expected Results (All Tests):
- No AccessViolationException crashes
- No NullReferenceException crashes
- Files with brackets process correctly
- Locked files skipped with clear message
- Statistics accurate
"@

$readmeContent | Out-File $readmePath -Encoding UTF8
Write-Host "✓ Created README.txt with test instructions" -ForegroundColor Green
Write-Host ""

# ========================================
# SUMMARY
# ========================================
Write-Host "==================================" -ForegroundColor Cyan
Write-Host "Test Setup Complete!" -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Test folder: $TestFolder" -ForegroundColor Yellow
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "1. Open FileLabeler.ps1" -ForegroundColor White
Write-Host "2. Browse to test folder: $TestFolder" -ForegroundColor White
Write-Host "3. Follow README.txt for test instructions" -ForegroundColor White
Write-Host "4. Verify all tests pass before GitHub push" -ForegroundColor White
Write-Host ""
Write-Host "Opening test folder..." -ForegroundColor Yellow
Start-Process explorer.exe -ArgumentList $TestFolder

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

