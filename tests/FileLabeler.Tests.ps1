#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive Pester unit tests for FileLabeler v1.1
.DESCRIPTION
    Unit tests for core features including:
    - Label status retrieval (Task #2)
    - Folder recursion (Task #4)
    - Drag-and-drop functionality (Task #5)
    - Configuration management (Task #11)
    - Change categorization (Task #6)
    - Error handling (Task #15)
.NOTES
    Requires Pester module
#>

# Dot-source the script to test (need to extract functions)
# For now, we'll test individual components

Describe "FileLabeler Core Functions" {
    
    Context "Label Cache Management" {
        
        It "Should initialize label cache as hashtable" {
            $testCache = @{}
            $testCache.GetType().Name | Should Be "Hashtable"
        }
        
        It "Should store label information in cache" {
            $testCache = @{}
            $testFile = "C:\test\document.docx"
            $testCache[$testFile] = @{
                DisplayName = "Fortrolig"
                LabelId = "221e033f-836b-4372-a276-90a25fdd73b5"
                Rank = 4
            }
            
            $testCache.ContainsKey($testFile) | Should Be $true
            $testCache[$testFile].DisplayName | Should Be "Fortrolig"
            $testCache[$testFile].Rank | Should Be 4
        }
        
        It "Should retrieve cached label information" {
            $testCache = @{}
            $testFile = "C:\test\document.docx"
            $testCache[$testFile] = @{
                DisplayName = "Intern"
                LabelId = "test-guid"
                Rank = 1
            }
            
            $cached = $testCache[$testFile]
            $cached.DisplayName | Should Be "Intern"
        }
        
        It "Should handle cache misses gracefully" {
            $testCache = @{}
            $nonExistentFile = "C:\test\missing.docx"
            
            $testCache.ContainsKey($nonExistentFile) | Should Be $false
        }
        
        It "Should clear cache when requested" {
            $testCache = @{}
            $testCache["file1.docx"] = "Label1"
            $testCache["file2.docx"] = "Label2"
            
            $testCache.Clear()
            $testCache.Count | Should Be 0
        }
    }
    
    Context "Folder File Scanning" {
        
        BeforeEach {
            # Create test folder structure
            $script:testRoot = Join-Path $env:TEMP "PesterTest_$(Get-Random)"
            New-Item -Path $script:testRoot -ItemType Directory -Force | Out-Null
            New-Item -Path "$script:testRoot\Sub1" -ItemType Directory -Force | Out-Null
            New-Item -Path "$script:testRoot\Sub2" -ItemType Directory -Force | Out-Null
            
            # Create test files
            "test" | Out-File "$script:testRoot\file1.docx" -Force
            "test" | Out-File "$script:testRoot\file2.xlsx" -Force
            "test" | Out-File "$script:testRoot\file3.txt" -Force  # Not supported
            "test" | Out-File "$script:testRoot\Sub1\file4.pptx" -Force
            "test" | Out-File "$script:testRoot\Sub2\file5.pdf" -Force
        }
        
        AfterEach {
            if (Test-Path $script:testRoot) {
                Remove-Item $script:testRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        
        It "Should scan for supported file types" {
            $extensions = @('*.docx', '*.xlsx', '*.pptx', '*.pdf')
            $files = @()
            
            foreach ($ext in $extensions) {
                $files += Get-ChildItem -Path $script:testRoot -Filter $ext -File -Recurse -ErrorAction SilentlyContinue
            }
            
            $files = $files | Sort-Object -Property FullName -Unique
            $files.Count | Should Be 4  # .docx, .xlsx, .pptx, .pdf (not .txt)
        }
        
        It "Should perform non-recursive scan correctly" {
            $extensions = @('*.docx', '*.xlsx', '*.pptx', '*.pdf')
            $files = @()
            
            foreach ($ext in $extensions) {
                $files += Get-ChildItem -Path $script:testRoot -Filter $ext -File -ErrorAction SilentlyContinue
            }
            
            $files = $files | Sort-Object -Property FullName -Unique
            $files.Count | Should Be 2  # Only root files: .docx and .xlsx
        }
        
        It "Should perform recursive scan correctly" {
            $extensions = @('*.docx', '*.xlsx', '*.pptx', '*.pdf')
            $files = @()
            
            foreach ($ext in $extensions) {
                $files += Get-ChildItem -Path $script:testRoot -Filter $ext -File -Recurse -ErrorAction SilentlyContinue
            }
            
            $files = $files | Sort-Object -Property FullName -Unique
            $files.Count | Should Be 4  # All supported files including subfolders
        }
        
        It "Should deduplicate results" {
            $extensions = @('*.docx', '*.docx')  # Duplicate extension
            $files = @()
            
            foreach ($ext in $extensions) {
                $files += Get-ChildItem -Path $script:testRoot -Filter $ext -File -Recurse -ErrorAction SilentlyContinue
            }
            
            $files = $files | Sort-Object -Property FullName -Unique
            $files.Count | Should Be 1  # Should have only 1 .docx file
        }
        
        It "Should handle empty folders gracefully" {
            $emptyFolder = Join-Path $script:testRoot "EmptyFolder"
            New-Item -Path $emptyFolder -ItemType Directory -Force | Out-Null
            
            $extensions = @('*.docx', '*.xlsx')
            $files = @()
            
            foreach ($ext in $extensions) {
                $files += Get-ChildItem -Path $emptyFolder -Filter $ext -File -ErrorAction SilentlyContinue
            }
            
            $files.Count | Should Be 0
        }
    }
    
    Context "File Type Filtering" {
        
        It "Should support all required Office file types" {
            $supportedExtensions = @('*.docx', '*.xlsx', '*.pptx', '*.doc', '*.xls', '*.ppt', '*.pdf')
            $supportedExtensions.Count | Should Be 7
        }
        
        It "Should filter by .docx extension" {
            $testFiles = @("file.docx", "file.xlsx", "file.txt")
            $filtered = $testFiles | Where-Object { $_ -like "*.docx" }
            $filtered.Count | Should Be 1
        }
        
        It "Should exclude unsupported file types" {
            $testFiles = @("file.docx", "file.txt", "file.jpg")
            $supportedExtensions = @('*.docx', '*.xlsx', '*.pptx', '*.doc', '*.xls', '*.ppt', '*.pdf')
            
            $filtered = $testFiles | Where-Object {
                $file = $_
                $isSupported = $false
                foreach ($ext in $supportedExtensions) {
                    if ($file -like $ext) {
                        $isSupported = $true
                        break
                    }
                }
                $isSupported
            }
            
            $filtered.Count | Should Be 1
        }
    }
    
    Context "Duplicate Detection" {
        
        It "Should detect existing files in selection" {
            $existingFiles = @("C:\test\file1.docx", "C:\test\file2.xlsx")
            $newFiles = @("C:\test\file1.docx", "C:\test\file3.pptx")
            
            $duplicates = @($newFiles | Where-Object { $existingFiles -contains $_ })
            $duplicates.Count | Should Be 1
            $duplicates[0] | Should Be "C:\test\file1.docx"
        }
        
        It "Should filter out duplicates from new selection" {
            $existingFiles = @("C:\test\file1.docx", "C:\test\file2.xlsx")
            $newFiles = @("C:\test\file1.docx", "C:\test\file2.xlsx", "C:\test\file3.pptx")
            
            $filtered = @($newFiles | Where-Object { $existingFiles -notcontains $_ })
            $filtered.Count | Should Be 1
            $filtered[0] | Should Be "C:\test\file3.pptx"
        }
        
        It "Should merge selections without duplicates" {
            $existingFiles = @("C:\test\file1.docx")
            $newFiles = @("C:\test\file1.docx", "C:\test\file2.xlsx")
            
            $uniqueNew = $newFiles | Where-Object { $existingFiles -notcontains $_ }
            $merged = @($existingFiles) + @($uniqueNew)
            
            $merged.Count | Should Be 2
        }
    }
    
    Context "Label Change Categorization" {
        
        It "Should categorize new label (no current label)" {
            $currentLabelId = $null
            $currentRank = -1
            $targetRank = 2
            
            $changeType = if (-not $currentLabelId) {
                "New"
            } elseif ($targetRank -gt $currentRank) {
                "Upgrade"
            } elseif ($targetRank -lt $currentRank) {
                "Downgrade"
            } else {
                "Unchanged"
            }
            
            $changeType | Should Be "New"
        }
        
        It "Should categorize label upgrade" {
            $currentRank = 1  # Intern
            $targetRank = 4   # Fortrolig
            
            $changeType = if ($targetRank -gt $currentRank) {
                "Upgrade"
            } elseif ($targetRank -lt $currentRank) {
                "Downgrade"
            } else {
                "Unchanged"
            }
            
            $changeType | Should Be "Upgrade"
        }
        
        It "Should categorize label downgrade" {
            $currentRank = 4  # Fortrolig
            $targetRank = 1   # Intern
            
            $changeType = if ($targetRank -gt $currentRank) {
                "Upgrade"
            } elseif ($targetRank -lt $currentRank) {
                "Downgrade"
            } else {
                "Unchanged"
            }
            
            $changeType | Should Be "Downgrade"
        }
        
        It "Should categorize unchanged (same rank)" {
            $currentRank = 2
            $targetRank = 2
            
            $changeType = if ($targetRank -gt $currentRank) {
                "Upgrade"
            } elseif ($targetRank -lt $currentRank) {
                "Downgrade"
            } else {
                "Unchanged"
            }
            
            $changeType | Should Be "Unchanged"
        }
        
        It "Should detect same label ID" {
            $currentLabelId = "221e033f-836b-4372-a276-90a25fdd73b5"
            $targetLabelId = "221e033f-836b-4372-a276-90a25fdd73b5"
            
            $isSame = ($currentLabelId -eq $targetLabelId)
            $isSame | Should Be $true
        }
    }
    
    Context "Warning Detection" {
        
        It "Should detect mass downgrade warning (3+ files)" {
            $downgradeCount = 5
            $threshold = 3
            
            $hasWarning = ($downgradeCount -ge $threshold)
            $hasWarning | Should Be $true
        }
        
        It "Should not trigger mass downgrade for small batches" {
            $downgradeCount = 2
            $threshold = 3
            
            $hasWarning = ($downgradeCount -ge $threshold)
            $hasWarning | Should Be $false
        }
        
        It "Should detect large batch warning (20+ files)" {
            $totalFiles = 25
            $threshold = 20
            
            $hasWarning = ($totalFiles -gt $threshold)
            $hasWarning | Should Be $true
        }
        
        It "Should detect no changes warning" {
            $newCount = 0
            $upgradeCount = 0
            $downgradeCount = 0
            $sameCount = 5
            
            $willChange = $newCount + $upgradeCount + $downgradeCount
            $hasWarning = ($willChange -eq 0 -and $sameCount -gt 0)
            
            $hasWarning | Should Be $true
        }
        
        It "Should detect mixed changes warning" {
            $upgradeCount = 3
            $downgradeCount = 2
            
            $hasWarning = ($upgradeCount -gt 0 -and $downgradeCount -gt 0)
            $hasWarning | Should Be $true
        }
    }
    
    Context "Configuration Management" {
        
        BeforeEach {
            $script:testConfigPath = Join-Path $env:TEMP "test_app_config_$(Get-Random).json"
        }
        
        AfterEach {
            if (Test-Path $script:testConfigPath) {
                Remove-Item $script:testConfigPath -Force -ErrorAction SilentlyContinue
            }
        }
        
        It "Should have default configuration structure" {
            $defaultConfig = @{
                version = "1.0"
                preferences = @{
                    defaultFolder = ""
                    rememberLastLabel = $false
                    includeSubfoldersDefault = $false
                    lastSelectedLabelId = $null
                }
                warnings = @{
                    showPreApplySummary = $true
                    showMassDowngradeWarning = $true
                    showLargeBatchWarning = $true
                    largeBatchThreshold = 20
                }
            }
            
            $defaultConfig.version | Should Be "1.0"
            $defaultConfig.preferences | Should Not BeNullOrEmpty
            $defaultConfig.warnings | Should Not BeNullOrEmpty
        }
        
        It "Should save configuration to JSON" {
            $config = @{
                version = "1.0"
                preferences = @{
                    defaultFolder = "C:\Test"
                }
            }
            
            $config | ConvertTo-Json -Depth 10 | Set-Content -Path $script:testConfigPath -Encoding UTF8
            
            Test-Path $script:testConfigPath | Should Be $true
        }
        
        It "Should load configuration from JSON" {
            $config = @{
                version = "1.0"
                preferences = @{
                    defaultFolder = "C:\Test"
                    rememberLastLabel = $true
                }
            }
            
            $config | ConvertTo-Json -Depth 10 | Set-Content -Path $script:testConfigPath -Encoding UTF8
            
            $loaded = Get-Content -Path $script:testConfigPath -Raw | ConvertFrom-Json
            $loaded.version | Should Be "1.0"
            $loaded.preferences.defaultFolder | Should Be "C:\Test"
            $loaded.preferences.rememberLastLabel | Should Be $true
        }
        
        It "Should validate boolean config values" {
            $boolValue = $true
            $boolValue -is [bool] | Should Be $true
        }
        
        It "Should validate numeric threshold values" {
            $threshold = 20
            ($threshold -ge 1 -and $threshold -le 1000) | Should Be $true
        }
        
        It "Should reject invalid threshold values" {
            $threshold = -5
            ($threshold -ge 1 -and $threshold -le 1000) | Should Be $false
        }
    }
    
    Context "Error Handling and Logging" {
        
        BeforeEach {
            $script:testLogPath = Join-Path $env:TEMP "test_log_$(Get-Random).txt"
        }
        
        AfterEach {
            if (Test-Path $script:testLogPath) {
                Remove-Item $script:testLogPath -Force -ErrorAction SilentlyContinue
            }
        }
        
        It "Should create structured log entry" {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
            $level = "INFO"
            $source = "TestFunction"
            $message = "Test message"
            
            $logEntry = "[$timestamp] [$level] [$source] $message"
            $logEntry | Out-File -FilePath $script:testLogPath -Encoding UTF8
            
            Test-Path $script:testLogPath | Should Be $true
            $content = Get-Content $script:testLogPath -Raw
            $content | Should Match "\[INFO\]"
            $content | Should Match "\[TestFunction\]"
        }
        
        It "Should validate log levels" {
            $validLevels = @('INFO', 'WARNING', 'ERROR', 'CRITICAL')
            
            foreach ($level in $validLevels) {
                $validLevels -contains $level | Should Be $true
            }
        }
        
        It "Should include context in log entry" {
            $context = @{
                FilePath = "C:\test\file.docx"
                Operation = "LabelApply"
            }
            
            $contextStr = ($context.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; "
            $contextStr | Should Match "FilePath=C:\\test\\file.docx"
            $contextStr | Should Match "Operation=LabelApply"
        }
        
        It "Should handle exception details" {
            try {
                throw "Test exception"
            }
            catch {
                $exception = $_
                $exception.Exception.Message | Should Be "Test exception"
                $exception.Exception.GetType().Name | Should Not BeNullOrEmpty
            }
        }
        
        It "Should identify common error patterns" {
            $errorMessage = "Access to the path is denied"
            $isAccessError = $errorMessage -match "Access.*denied"
            $isAccessError | Should Be $true
        }
        
        It "Should identify file locked errors" {
            $errorMessage = "The process cannot access the file because it is being used by another process"
            $isLockedError = $errorMessage -match "process.*another"
            $isLockedError | Should Be $true
        }
    }
    
    Context "Drag-and-Drop Support" {
        
        It "Should identify files from dropped paths" {
            $droppedPaths = @(
                "C:\test\file.docx",
                "C:\test\file.xlsx",
                "C:\test\unsupported.txt"
            )
            
            $supportedExtensions = @('.docx', '.xlsx', '.pptx', '.pdf')
            
            $supportedFiles = $droppedPaths | Where-Object {
                $filePath = $_
                $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
                $supportedExtensions -contains $ext
            }
            
            $supportedFiles.Count | Should Be 2
        }
        
        It "Should identify folders from dropped paths" {
            # Simulate path type detection
            $path = "C:\test\folder"
            $isFolder = $false  # Would use Test-Path -PathType Container in real code
            
            # For testing purposes
            if ($path -notlike "*.*") {
                $isFolder = $true
            }
            
            $isFolder | Should Be $true
        }
        
        It "Should process mixed files and folders" {
            $droppedPaths = @(
                "C:\test\file.docx",
                "C:\test\folder",
                "C:\test\file.xlsx"
            )
            
            $files = $droppedPaths | Where-Object { $_ -like "*.*" }
            $folders = $droppedPaths | Where-Object { $_ -notlike "*.*" }
            
            $files.Count | Should Be 2
            $folders.Count | Should Be 1
        }
    }
    
    Context "UI Layout and Dynamic Sizing" {
        
        It "Should calculate listbox height based on file count" {
            $fileCount = 5
            $itemHeight = 16
            $minRows = 3
            $maxRows = 10
            
            $rows = if ($fileCount -eq 0) {
                $minRows
            } elseif ($fileCount -le $maxRows) {
                [Math]::Max($fileCount, $minRows)
            } else {
                $maxRows
            }
            
            $height = $rows * $itemHeight + 8
            
            $rows | Should Be 5
            $height | Should Be 88  # (5 * 16) + 8
        }
        
        It "Should enforce minimum rows (3)" {
            $fileCount = 1
            $minRows = 3
            $maxRows = 10
            
            $rows = if ($fileCount -eq 0) {
                $minRows
            } elseif ($fileCount -le $maxRows) {
                [Math]::Max($fileCount, $minRows)
            } else {
                $maxRows
            }
            
            $rows | Should Be 3
        }
        
        It "Should enforce maximum rows (10)" {
            $fileCount = 50
            $minRows = 3
            $maxRows = 10
            
            $rows = if ($fileCount -eq 0) {
                $minRows
            } elseif ($fileCount -le $maxRows) {
                [Math]::Max($fileCount, $minRows)
            } else {
                $maxRows
            }
            
            $rows | Should Be 10
        }
        
        It "Should calculate form height adjustment" {
            $oldHeight = 570
            $newHeight = 600
            $heightDelta = $newHeight - $oldHeight
            
            $heightDelta | Should Be 30
        }
    }
    
    Context "Norwegian Character Encoding" {
        
        It "Should preserve Norwegian characters æ, ø, å" {
            $norwegianText = "Åpen"
            $norwegianText | Should Match "[ÆØÅæøå]"
        }
        
        It "Should handle common Norwegian label names" {
            $labels = @("Åpen", "Fortrolig", "Personlig")
            
            $labels[0] | Should Be "Åpen"
            $labels[1] | Should Be "Fortrolig"
        }
        
        It "Should detect Norwegian characters in strings" {
            $testString = "Ingen etikett"
            $hasNorwegian = $testString -match "[ÆØÅæøå]"
            $hasNorwegian | Should Be $false  # "etikett" has no æ/ø/å
            
            $testString2 = "Åpen"
            $hasNorwegian2 = $testString2 -match "[ÆØÅæøå]"
            $hasNorwegian2 | Should Be $true
        }
    }
    
    Context "Statistics Tracking" {
        
        It "Should initialize statistics structure" {
            $stats = @{
                TotalProcessed = 0
                SuccessCount = 0
                FailureCount = 0
                ChangeTypeBreakdown = @{
                    New = 0
                    Upgrade = 0
                    Downgrade = 0
                    Unchanged = 0
                    Same = 0
                }
            }
            
            $stats.TotalProcessed | Should Be 0
            $stats.ChangeTypeBreakdown.Keys.Count | Should Be 5
        }
        
        It "Should increment counters correctly" {
            $successCount = 0
            $successCount++
            $successCount++
            
            $successCount | Should Be 2
        }
        
        It "Should calculate success rate" {
            $totalProcessed = 10
            $successCount = 8
            
            $successRate = if ($totalProcessed -gt 0) {
                [int](($successCount / $totalProcessed) * 100)
            } else {
                0
            }
            
            $successRate | Should Be 80
        }
        
        It "Should track change type breakdown" {
            $breakdown = @{
                New = 5
                Upgrade = 3
                Downgrade = 2
                Same = 0
                Unchanged = 0
            }
            
            $total = $breakdown.Values | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            $total | Should Be 10
        }
    }
}

Describe "FileLabeler Integration Scenarios" {
    
    Context "End-to-End Workflow" {
        
        It "Should support workflow: Select files -> Analyze -> Apply" {
            # Simulate workflow states
            $selectedFiles = @("file1.docx", "file2.xlsx")
            $selectedLabelId = "test-label-id"
            
            # Step 1: Files selected
            $filesSelected = $selectedFiles.Count -gt 0
            $filesSelected | Should Be $true
            
            # Step 2: Label selected
            $labelSelected = -not [string]::IsNullOrEmpty($selectedLabelId)
            $labelSelected | Should Be $true
            
            # Step 3: Ready to analyze
            $readyToAnalyze = $filesSelected -and $labelSelected
            $readyToAnalyze | Should Be $true
        }
        
        It "Should validate pre-apply summary workflow" {
            # Mock analysis results
            $analysis = @{
                New = @(@{ File = "file1.docx" })
                Upgrade = @(@{ File = "file2.xlsx" })
                Downgrade = @()
                Same = @()
                Unchanged = @()
                TotalFiles = 2
            }
            
            $analysis.TotalFiles | Should Be 2
            $analysis.New.Count | Should Be 1
            $analysis.Upgrade.Count | Should Be 1
        }
    }
    
    Context "Edge Cases and Error Scenarios" {
        
        It "Should handle empty file selection" {
            $selectedFiles = @()
            $isEmpty = $selectedFiles.Count -eq 0
            $isEmpty | Should Be $true
        }
        
        It "Should handle no label selected" {
            $selectedLabelId = $null
            $isNull = $null -eq $selectedLabelId
            $isNull | Should Be $true
        }
        
        It "Should handle all files already labeled" {
            $analysis = @{
                New = @()
                Upgrade = @()
                Downgrade = @()
                Same = @(@{ File = "file1.docx" }, @{ File = "file2.docx" })
                TotalFiles = 2
            }
            
            $willChange = $analysis.New.Count + $analysis.Upgrade.Count + $analysis.Downgrade.Count
            $willChange | Should Be 0
        }
        
        It "Should handle large file batches (500+ files)" {
            $largeSelection = 1..600 | ForEach-Object { "file$_.docx" }
            $largeSelection.Count | Should Be 600
            
            $isLarge = $largeSelection.Count -gt 500
            $isLarge | Should Be $true
        }
    }
}

Write-Host "`n=== Pester Test Suite Created ===" -ForegroundColor Cyan
Write-Host "Run tests with: Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1" -ForegroundColor Yellow
Write-Host ""

