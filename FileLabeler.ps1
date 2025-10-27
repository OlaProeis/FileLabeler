#Requires -Version 5.1
<#
.SYNOPSIS
    Massemerking av filer - Påfør følsomhetsetiketter med datobevaring
.DESCRIPTION
    GUI-applikasjon for massepåføring av Microsoft Purview følsomhetsetiketter til Office-dokumenter og PDF-er.
    Bevarer opprinnelige fildatoer.
.NOTES
    Krever: Microsoft Purview Information Protection-klient og PurviewInformationProtection-modul
#>

# ========================================
# CRITICAL: PowerShell Transcript
# ========================================
# Captures ALL output, errors, and warnings - even crashes that escape error handlers
$transcriptDir = Join-Path $env:USERPROFILE "Documents\FileLabeler_Logs"
if (-not (Test-Path $transcriptDir)) {
    New-Item -Path $transcriptDir -ItemType Directory -Force | Out-Null
}
$transcriptPath = Join-Path $transcriptDir "TRANSCRIPT_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
Start-Transcript -Path $transcriptPath -Force | Out-Null

# Set output encoding to UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

# Import necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable visual styles
[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# GLOBAL ERROR HANDLER
# ========================================
# Trap ALL unhandled exceptions before they crash the application
trap {
    $crashTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $crashLogDir = Join-Path $env:USERPROFILE "Documents\FileLabeler_Logs"
    
    # Ensure log directory exists
    if (-not (Test-Path $crashLogDir)) {
        New-Item -Path $crashLogDir -ItemType Directory -Force | Out-Null
    }
    
    # Create detailed crash report
    $crashInfo = @{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        ExceptionMessage = $_.Exception.Message
        ExceptionType = $_.Exception.GetType().FullName
        StackTrace = if ($_.ScriptStackTrace) { $_.ScriptStackTrace } else { "No stack trace available" }
        InvocationLine = if ($_.InvocationInfo) { $_.InvocationInfo.Line } else { "Unknown" }
        InvocationPosition = if ($_.InvocationInfo) { $_.InvocationInfo.PositionMessage } else { "Unknown" }
        CategoryInfo = $_.CategoryInfo.ToString()
        FullyQualifiedErrorId = $_.FullyQualifiedErrorId
    }
    
    # Save as JSON for easy parsing
    $crashLogPath = Join-Path $crashLogDir "CRASH_$crashTimestamp.json"
    try {
        $crashInfo | ConvertTo-Json -Depth 10 | Set-Content $crashLogPath -Encoding UTF8
    } catch {
        # Fallback to text if JSON fails
        $crashLogPath = Join-Path $crashLogDir "CRASH_$crashTimestamp.txt"
        $crashInfo | Out-String | Set-Content $crashLogPath
    }
    
    # Also write to regular log if available
    if (Test-Path variable:script:logFilePath) {
        try {
            Add-Content -Path $script:logFilePath -Value "`n=== UNHANDLED EXCEPTION TRAPPED ===" -ErrorAction SilentlyContinue
            Add-Content -Path $script:logFilePath -Value "[$($crashInfo.Timestamp)] [CRITICAL] $($crashInfo.ExceptionMessage)" -ErrorAction SilentlyContinue
            Add-Content -Path $script:logFilePath -Value "Type: $($crashInfo.ExceptionType)" -ErrorAction SilentlyContinue
            Add-Content -Path $script:logFilePath -Value "Stack: $($crashInfo.StackTrace)" -ErrorAction SilentlyContinue
        } catch {}
    }
    
    # Show error dialog to user
    $errorMsg = "KRITISK FEIL - Applikasjonen krasjet!`n`n"
    $errorMsg += "Feil: $($_.Exception.Message)`n`n"
    $errorMsg += "Type: $($_.Exception.GetType().Name)`n`n"
    $errorMsg += "Detaljer lagret i:`n$crashLogPath`n`n"
    $errorMsg += "Vennligst send denne filen til support."
    
    [System.Windows.Forms.MessageBox]::Show(
        $errorMsg,
        "Kritisk feil",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    
    # Try to open crash log
    try {
        Start-Process notepad.exe -ArgumentList $crashLogPath -ErrorAction SilentlyContinue
    } catch {}
    
    # Continue execution instead of terminating
    continue
}

# ========================================
# EARLY INITIALIZATION - LOG DIRECTORY
# ========================================
# Create log directory BEFORE anything else (needed for Write-Log)
$logDirectory = Join-Path $env:USERPROFILE "Documents\FileLabeler_Logs"
if (-not (Test-Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}
$logFilePath = Join-Path $logDirectory "FileLabeler_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# ========================================
# HELPER FUNCTIONS (DEFINED EARLY FOR USE THROUGHOUT)
# ========================================

<#
.SYNOPSIS
    Enhanced logging function with structured output and log levels
.DESCRIPTION
    Writes log entries with severity levels, source information, and diagnostic context
.PARAMETER Message
    The message to log
.PARAMETER Level
    Log severity level: INFO, WARNING, ERROR, CRITICAL
.PARAMETER Source
    Source of the log entry (function/operation name)
.PARAMETER Context
    Additional context (e.g., file path, operation details)
.PARAMETER Exception
    Exception object to include stack trace and details
#>
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'CRITICAL')]
        [string]$Level = 'INFO',
        
        [Parameter(Mandatory=$false)]
        [string]$Source = '',
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Context = @{},
        
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.ErrorRecord]$Exception = $null
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        
        # Build structured log entry
        $logEntry = "[$timestamp] [$Level]"
        
        if ($Source) {
            $logEntry += " [$Source]"
        }
        
        $logEntry += " $Message"
        
        # Add context if provided
        if ($Context.Count -gt 0) {
            $contextStr = ($Context.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; "
            $logEntry += " | Context: $contextStr"
        }
        
        # Add exception details if provided
        if ($Exception) {
            $logEntry += "`n    Exception: $($Exception.Exception.Message)"
            $logEntry += "`n    Type: $($Exception.Exception.GetType().FullName)"
            
            # Include stack trace for ERROR and CRITICAL levels
            if ($Level -in @('ERROR', 'CRITICAL') -and $Exception.ScriptStackTrace) {
                $logEntry += "`n    StackTrace: $($Exception.ScriptStackTrace)"
            }
            
            # Include inner exception if present
            if ($Exception.Exception.InnerException) {
                $logEntry += "`n    InnerException: $($Exception.Exception.InnerException.Message)"
            }
        }
        
        # Write to log file
        Add-Content -Path $script:logFilePath -Value $logEntry -ErrorAction SilentlyContinue
        
        # For CRITICAL errors, also write to Windows Event Log (if possible)
        if ($Level -eq 'CRITICAL') {
            try {
                Write-EventLog -LogName Application -Source "FileLabeler" -EventId 1000 -EntryType Error -Message $Message -ErrorAction SilentlyContinue
            } catch {
                # Silently fail if event log not accessible
            }
        }
    }
    catch {
        # Fallback: write minimal log entry if structured logging fails
        try {
            $fallbackEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Logging failed: $($_.Exception.Message). Original message: $Message"
            Add-Content -Path $script:logFilePath -Value $fallbackEntry -ErrorAction SilentlyContinue
        }
        catch {
            # Ultimate fallback: do nothing to avoid infinite loops
        }
    }
}

<#
.SYNOPSIS
    Get user-friendly error message for common error scenarios
.DESCRIPTION
    Translates technical errors into actionable user messages
.PARAMETER Exception
    The exception to translate
.OUTPUTS
    Hashtable with UserMessage and TechnicalDetails
#>
function Get-FriendlyErrorMessage {
    param(
        [Parameter(Mandatory=$true)]
        $Exception
    )
    
    $errorMessage = $Exception.Exception.Message
    $errorType = $Exception.Exception.GetType().Name
    
    # Common error patterns and user-friendly messages
    $errorPatterns = @{
        # File access errors
        'UnauthorizedAccessException|Access.*denied' = @{
            UserMessage = "Ingen tilgang til filen. Sjekk at du har nødvendige tillatelser."
            Category = "FileAccess"
            Suggestion = "Høyreklikk på filen, velg Egenskaper → Sikkerhet, og kontroller dine tillatelser."
        }
        'IOException.*process.*another' = @{
            UserMessage = "Filen er i bruk av et annet program. Lukk filen og prøv igjen."
            Category = "FileLocked"
            Suggestion = "Lukk dokumentet i Word/Excel/PowerPoint og prøv på nytt."
        }
        'FileNotFoundException|Could not find.*file' = @{
            UserMessage = "Filen ble ikke funnet. Den kan være flyttet eller slettet."
            Category = "FileNotFound"
            Suggestion = "Kontroller at filen fortsatt eksisterer på angitt plassering."
        }
        # Network errors
        'IOException.*network' = @{
            UserMessage = "Nettverksfeil. Kontroller nettverkstilkoblingen og prøv igjen."
            Category = "Network"
            Suggestion = "Sjekk nettverkstilkobling. Hvis filen er på en nettverksmappe, kontroller at du er koblet til nettverket."
        }
        'DirectoryNotFoundException' = @{
            UserMessage = "Mappen ble ikke funnet. Kontroller at stien er korrekt."
            Category = "DirectoryNotFound"
            Suggestion = "Sjekk at mappen eksisterer og at stien er korrekt."
        }
        # AIP-specific errors
        'Justification' = @{
            UserMessage = "Begrunnelse kreves for å nedgradere følsomhetsetikett."
            Category = "AIPJustification"
            Suggestion = "Angi en gyldig begrunnelse for nedgraderingen."
        }
        'AdhocProtectionRequired|ad-hoc protection' = @{
            UserMessage = "Valgt etikett krever beskyttelsesinnstillinger."
            Category = "AIPProtection"
            Suggestion = "Angi hvem som skal ha tilgang til filen og hvilke rettigheter de skal ha."
        }
        # General errors
        'OutOfMemoryException' = @{
            UserMessage = "Ikke nok minne tilgjengelig. Prøv å behandle færre filer om gangen."
            Category = "Memory"
            Suggestion = "Lukk andre programmer eller reduser antall filer som behandles samtidig."
        }
        'TimeoutException' = @{
            UserMessage = "Operasjonen tok for lang tid og ble avbrutt."
            Category = "Timeout"
            Suggestion = "Prøv igjen med færre filer, eller sjekk nettverkshastigheten."
        }
    }
    
    # Find matching pattern
    foreach ($pattern in $errorPatterns.Keys) {
        if ($errorMessage -match $pattern -or $errorType -match $pattern) {
            return @{
                UserMessage = $errorPatterns[$pattern].UserMessage
                TechnicalDetails = $errorMessage
                Category = $errorPatterns[$pattern].Category
                Suggestion = $errorPatterns[$pattern].Suggestion
            }
        }
    }
    
    # Default fallback message
    return @{
        UserMessage = "En uventet feil oppstod under behandling."
        TechnicalDetails = $errorMessage
        Category = "Unknown"
        Suggestion = "Se loggfil for mer informasjon. Prøv operasjonen på nytt, eller kontakt support hvis feilen vedvarer."
    }
}

<#
.SYNOPSIS
    Show error dialog with user-friendly message and recovery options
.DESCRIPTION
    Displays error to user with actionable suggestions and optional log viewing
.PARAMETER ErrorInfo
    Hashtable from Get-FriendlyErrorMessage
.PARAMETER ShowLogOption
    Whether to show "View Log" button
#>
function Show-ErrorDialog {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$ErrorInfo,
        
        [Parameter(Mandatory=$false)]
        [bool]$ShowLogOption = $true,
        
        [Parameter(Mandatory=$false)]
        [string]$Title = "Feil"
    )
    
    # Build message
    $message = $ErrorInfo.UserMessage
    
    if ($ErrorInfo.Suggestion) {
        $message += "`n`nForslag: $($ErrorInfo.Suggestion)"
    }
    
    if ($ErrorInfo.TechnicalDetails -and $ErrorInfo.TechnicalDetails.Length -lt 200) {
        $message += "`n`nTeknisk detalj: $($ErrorInfo.TechnicalDetails)"
    }
    
    if ($ShowLogOption) {
        $message += "`n`nKlikk 'Vis logg' for mer informasjon."
    }
    
    # Show dialog
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    
    return $result
}

# ========================================
# MODULE VALIDATION
# ========================================
$moduleName = "PurviewInformationProtection"
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    [System.Windows.Forms.MessageBox]::Show(
        "ERROR: $moduleName module is not installed.`n`nPlease install the Microsoft Purview Information Protection client first.`n`nDownload from: https://www.microsoft.com/en-us/download/details.aspx?id=53018",
        "Missing Required Module",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

try {
    Import-Module $moduleName -ErrorAction Stop
    Write-Log -Message "$moduleName module imported successfully" -Level 'INFO' -Source 'ModuleValidation'
} catch {
    Write-Log -Message "Failed to import $moduleName module" -Level 'CRITICAL' -Source 'ModuleValidation' -Exception $_
    
    [System.Windows.Forms.MessageBox]::Show(
        "ERROR: Failed to import $moduleName module.`n`nError: $_",
        "Module Import Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# ========================================
# GET AVAILABLE LABELS
# ========================================
$labels = @()

# Method 1: Try to get labels from Security & Compliance Center (if module is available)
$securityModules = @("ExchangeOnlineManagement", "Microsoft.Online.SharePoint.PowerShell")
$labelsRetrieved = $false

foreach ($secModule in $securityModules) {
    if (Get-Module -ListAvailable -Name $secModule) {
        try {
            Import-Module $secModule -ErrorAction SilentlyContinue
            $labels = Get-Label -ErrorAction SilentlyContinue
            if ($labels -and $labels.Count -gt 0) {
                $labelsRetrieved = $true
                break
            }
        } catch {
            # Continue to next method
        }
    }
}

# Method 2: If no labels retrieved, use predefined list or allow manual entry
if (-not $labelsRetrieved -or $labels.Count -eq 0) {
    # Check if there's a labels configuration file
    $labelsConfigPath = Join-Path $PSScriptRoot "labels_config.json"
    
    if (Test-Path $labelsConfigPath) {
        try {
            $labelsJson = Get-Content $labelsConfigPath -Raw | ConvertFrom-Json
            $labels = $labelsJson
        } catch {
            # Will use default labels below
        }
    }
    
    # If still no labels, show a helpful message and provide default structure
    if (-not $labels -or $labels.Count -eq 0) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Could not automatically retrieve sensitivity labels.`n`nThis can happen if:`n- You're not connected to Security & Compliance Center`n- Labels need to be configured manually`n`nWould you like to:`nYES - Continue with manual label entry`nNO - Exit and configure labels first`n`nNote: You can create a 'labels_config.json' file in the script directory with your organization's labels.",
            "Label Configuration Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq 'No') {
            exit 0
        }
        
        # Provide empty list for manual entry
        $labels = @()
    }
}

# ========================================
# GLOBAL VARIABLES & CONSTANTS
# ========================================
$selectedFiles = @()
$fileLabelCache = @{}  # Cache for file label status to avoid repeated API calls

# Supported file extensions (centralized constant)
$script:SupportedExtensions = @('.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.pdf')
$script:SupportedExtensionPatterns = @('*.docx', '*.xlsx', '*.pptx', '*.doc', '*.xls', '*.ppt', '*.pdf')

# ========================================
# FILE MANAGEMENT HELPER FUNCTIONS
# ========================================

<#
.SYNOPSIS
    Scans folder for supported files
.DESCRIPTION
    Centralized function for scanning folders (sync or async) for supported file types
.PARAMETER FolderPath
    Path to folder to scan
.PARAMETER Recursive
    Whether to scan subfolders
.OUTPUTS
    Array of FileInfo objects
#>
function Get-SupportedFilesFromFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [Parameter(Mandatory=$false)]
        [bool]$Recursive = $false
    )
    
    $foundFiles = @()
    
    try {
        foreach ($ext in $script:SupportedExtensionPatterns) {
            if ($Recursive) {
                $foundFiles += Get-ChildItem -Path $FolderPath -Filter $ext -Recurse -File -ErrorAction SilentlyContinue
            } else {
                $foundFiles += Get-ChildItem -Path $FolderPath -Filter $ext -File -ErrorAction SilentlyContinue
            }
        }
        
        # Remove duplicates (important!)
        $foundFiles = $foundFiles | Sort-Object -Property FullName -Unique
        
    } catch {
        Write-Log -Message "Folder scan failed" -Level 'ERROR' -Source 'Get-SupportedFilesFromFolder' -Context @{ FolderPath = $FolderPath; Recursive = $Recursive } -Exception $_
        throw
    }
    
    return $foundFiles
}

<#
.SYNOPSIS
    Merges new files with existing selection, avoiding duplicates
.DESCRIPTION
    Centralized logic for merging file selections throughout the app
.PARAMETER NewFiles
    Array of FileInfo objects or file paths to add
.PARAMETER ExistingFiles
    Array of existing file paths (defaults to $script:selectedFiles)
.OUTPUTS
    Hashtable with MergedFiles, NewCount, DuplicateCount
#>
function Merge-FileSelection {
    param(
        [Parameter(Mandatory=$true)]
        [array]$NewFiles,
        [Parameter(Mandatory=$false)]
        [array]$ExistingFiles = $null
    )
    
    if ($null -eq $ExistingFiles) {
        $ExistingFiles = @($script:selectedFiles)
    }
    
    # Extract full paths from FileInfo objects if needed
    $newPaths = $NewFiles | ForEach-Object {
        if ($_ -is [System.IO.FileInfo]) {
            $_.FullName
        } else {
            $_
        }
    }
    
    # Filter out duplicates
    $uniqueNewPaths = $newPaths | Where-Object { $ExistingFiles -notcontains $_ }
    
    # Merge arrays
    $mergedFiles = @($ExistingFiles) + @($uniqueNewPaths)
    
    return @{
        MergedFiles = $mergedFiles
        NewCount = $uniqueNewPaths.Count
        TotalCount = $NewFiles.Count
        DuplicateCount = $NewFiles.Count - $uniqueNewPaths.Count
    }
}

# ========================================
# ASYNC RUNSPACE HELPER FUNCTIONS
# ========================================

function New-FileLabelerRunspacePool {
    <#
    .SYNOPSIS
        Creates optimized runspace pool for async operations
    .DESCRIPTION
        Sets up runspace pool with proper threading, module imports, and helper functions
    #>
    param(
        [int]$MinRunspaces = 1,
        [int]$MaxRunspaces = 4
    )
    
    # Limit max runspaces to prevent overwhelming the system
    $MaxRunspaces = [Math]::Min($MaxRunspaces, [Math]::Min([Environment]::ProcessorCount, 8))
    
    # Create initial session state
    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    # Import PurviewInformationProtection module
    $sessionState.ImportPSModule("PurviewInformationProtection")
    
    # Create runspace pool
    $pool = [RunspaceFactory]::CreateRunspacePool($MinRunspaces, $MaxRunspaces, $sessionState, $Host)
    $pool.ApartmentState = "MTA"  # MTA for background workers
    $pool.ThreadOptions = "ReuseThread"
    $pool.Open()
    
    return $pool
}

function Start-AsyncFolderScan {
    <#
    .SYNOPSIS
        Scans folder asynchronously without blocking UI
    .DESCRIPTION
        Uses runspace to scan folder recursively and returns PowerShell instance for tracking
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,
        [Parameter(Mandatory=$true)]
        [hashtable]$SharedData,
        [bool]$Recursive = $false
    )
    
    $scanScript = {
        param($Folder, $Recursive, $SharedData)
        
        $extensions = @('*.docx', '*.xlsx', '*.pptx', '*.doc', '*.xls', '*.ppt', '*.pdf')
        $foundFiles = @()
        
        try {
            foreach ($ext in $extensions) {
                if ($Recursive) {
                    $foundFiles += Get-ChildItem -Path $Folder -Filter $ext -Recurse -File -ErrorAction SilentlyContinue
                } else {
                    $foundFiles += Get-ChildItem -Path $Folder -Filter $ext -File -ErrorAction SilentlyContinue
                }
            }
            
            # Remove duplicates
            $foundFiles = $foundFiles | Sort-Object -Property FullName -Unique
            
            # Update shared data thread-safely
            $lockTaken = $false
            try {
                [System.Threading.Monitor]::Enter($SharedData.SyncRoot, [ref]$lockTaken)
                
                foreach ($file in $foundFiles) {
                    if (-not $SharedData.ScannedFiles.Contains($file.FullName)) {
                        [void]$SharedData.ScannedFiles.Add($file.FullName)
                    }
                }
                $SharedData.ScanComplete = $true
                $SharedData.ScanSuccess = $true
                
            } finally {
                if ($lockTaken) {
                    [System.Threading.Monitor]::Exit($SharedData.SyncRoot)
                }
            }
            
            return $foundFiles.Count
            
        } catch {
            # Update shared data with error
            $lockTaken = $false
            try {
                [System.Threading.Monitor]::Enter($SharedData.SyncRoot, [ref]$lockTaken)
                $SharedData.ScanComplete = $true
                $SharedData.ScanSuccess = $false
                $SharedData.ScanError = $_.Exception.Message
            } finally {
                if ($lockTaken) {
                    [System.Threading.Monitor]::Exit($SharedData.SyncRoot)
                }
            }
            
            throw
        }
    }
    
    # Create PowerShell instance
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $RunspacePool
    [void]$ps.AddScript($scanScript)
    [void]$ps.AddParameter("Folder", $FolderPath)
    [void]$ps.AddParameter("Recursive", $Recursive)
    [void]$ps.AddParameter("SharedData", $SharedData)
    
    # Start async execution
    $handle = $ps.BeginInvoke()
    
    return @{
        PowerShell = $ps
        Handle = $handle
        StartTime = Get-Date
    }
}

function Start-AsyncLabelRetrieval {
    <#
    .SYNOPSIS
        Retrieves file labels asynchronously in batches
    .DESCRIPTION
        Processes file labels in parallel using runspace pool with progress tracking
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$FilePaths,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,
        [Parameter(Mandatory=$true)]
        [hashtable]$SharedCache,
        [Parameter(Mandatory=$true)]
        [hashtable]$SharedProgress,
        [array]$AvailableLabels
    )
    
    $labelRetrievalScript = {
        param($FilePath, $SharedCache, $SharedProgress, $Labels)
        
        # === CRITICAL: RUNSPACE-LEVEL ERROR HANDLER ===
        # Prevent runspace termination from killing the entire process
        $ErrorActionPreference = 'Continue'
        
        $result = @{
            FilePath = $FilePath
            Success = $false
            DisplayName = "Ukjent"
            LabelId = $null
            Rank = -1
        }
        
        try {
            # ULTIMATE WRAPPER: Catch everything in this runspace
        try {
            # Check cache first (thread-safe read with lock)
            $lockTaken = $false
            try {
                [System.Threading.Monitor]::Enter($SharedCache.SyncRoot, [ref]$lockTaken)
                
                if ($SharedCache.ContainsKey($FilePath)) {
                    $cached = $SharedCache[$FilePath]
                    $result.Success = $true
                    $result.DisplayName = $cached.DisplayName
                    $result.LabelId = $cached.LabelId
                    $result.Rank = $cached.Rank
                    
                    # Increment progress (thread-safe)
                    [System.Threading.Interlocked]::Increment([ref]$SharedProgress.Processed)
                    return $result
                }
            } finally {
                if ($lockTaken) {
                    [System.Threading.Monitor]::Exit($SharedCache.SyncRoot)
                }
            }
            
            # Retrieve label from AIP
            # Don't use SilentlyContinue - let errors be caught by outer try/catch
            $labelStatus = Get-AIPFileStatus -Path $FilePath
            
            if ($labelStatus -and $labelStatus.MainLabelId) {
                $labelObj = $Labels | Where-Object { $_.Id -eq $labelStatus.MainLabelId }
                if ($labelObj) {
                    $result.Success = $true
                    $result.DisplayName = $labelObj.DisplayName
                    $result.LabelId = $labelStatus.MainLabelId
                    $result.Rank = if($labelObj.Rank) { $labelObj.Rank } else { 0 }
                } else {
                    # Label ID exists but not found in configuration
                    # This can happen with encrypted/protected labels not in labels_config.json
                    $result.Success = $true
                    $result.DisplayName = "Ukjent etikett (beskyttet)"
                    $result.LabelId = $labelStatus.MainLabelId
                    $result.Rank = -1
                }
            } else {
                # No label found
                $result.Success = $true
                $result.DisplayName = "Ingen etikett"
            }
            
            # Update cache (thread-safe write)
            $lockTaken = $false
            try {
                [System.Threading.Monitor]::Enter($SharedCache.SyncRoot, [ref]$lockTaken)
                $SharedCache[$FilePath] = @{
                    DisplayName = $result.DisplayName
                    LabelId = $result.LabelId
                    Rank = $result.Rank
                }
            } finally {
                if ($lockTaken) {
                    [System.Threading.Monitor]::Exit($SharedCache.SyncRoot)
                }
            }
            
        } catch {
            # Enhanced error logging for diagnostics
            # Note: Cannot use Write-Log in runspace, store error details for later analysis
            $result.Success = $false
            $result.DisplayName = "Feil ved henting"
            $result.ErrorType = $_.Exception.GetType().Name
            $result.ErrorMessage = $_.Exception.Message
        }
        
        } catch {
            # === ULTIMATE LABEL RETRIEVAL CATCH ===
            # Catches ANYTHING that escaped all other handlers
            # Prevents runspace termination during label retrieval
            
            # Ensure result has required fields
            if (-not $result.FilePath) { $result.FilePath = $FilePath }
            if (-not $result.DisplayName) { $result.DisplayName = "Runspace feil" }
            
            $result.Success = $false
            $result.ErrorType = $_.Exception.GetType().FullName
            $result.ErrorMessage = "LABEL RETRIEVAL CRASH: $($_.Exception.Message)"
        }
        
        # Increment progress (thread-safe)
        [System.Threading.Interlocked]::Increment([ref]$SharedProgress.Processed)
        
        return $result
    }
    
    # Start jobs for all files
    $jobs = @()
    
    foreach ($filePath in $FilePaths) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $RunspacePool
        [void]$ps.AddScript($labelRetrievalScript)
        [void]$ps.AddParameter("FilePath", $filePath)
        [void]$ps.AddParameter("SharedCache", $SharedCache)
        [void]$ps.AddParameter("SharedProgress", $SharedProgress)
        [void]$ps.AddParameter("Labels", $AvailableLabels)
        
        $handle = $ps.BeginInvoke()
        
        $jobs += [PSCustomObject]@{
            PowerShell = $ps
            Handle = $handle
            FilePath = $filePath
        }
    }
    
    return $jobs
}

function Update-UIThreadSafe {
    <#
    .SYNOPSIS
        Safely updates UI control from any thread
    .DESCRIPTION
        Checks if invoke is required and marshals update to UI thread
    #>
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Control]$Control,
        [Parameter(Mandatory=$true)]
        [scriptblock]$UpdateAction
    )
    
    if ($Control.InvokeRequired) {
        $Control.Invoke([Action]$UpdateAction)
    } else {
        & $UpdateAction
    }
}

function Wait-AsyncJobsWithUI {
    <#
    .SYNOPSIS
        Waits for async jobs while keeping UI responsive
    .DESCRIPTION
        Monitors job completion and updates progress bar, allowing UI to process events
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Jobs,
        [Parameter(Mandatory=$true)]
        [hashtable]$SharedProgress,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Label]$StatusLabel,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Form]$Form,
        [int]$UpdateIntervalMs = 100,
        [string]$OperationType = "Behandler"
    )
    
    $results = @()
    $totalJobs = $Jobs.Count
    $startTime = Get-Date
    
    # Monitor until all jobs complete
    while ($true) {
        $completed = 0
        
        foreach ($job in $Jobs) {
            if ($job.Handle.IsCompleted) {
                $completed++
            }
        }
        
        # Calculate percentage and elapsed time
        $percentComplete = if ($totalJobs -gt 0) { [int](($completed / $totalJobs) * 100) } else { 0 }
        $elapsedSeconds = [int]((Get-Date) - $startTime).TotalSeconds
        
        # Calculate estimated time remaining
        $estimatedTotal = if ($completed -gt 0) { 
            ($elapsedSeconds / $completed) * $totalJobs 
        } else { 
            0 
        }
        $remainingSeconds = [Math]::Max(0, [int]($estimatedTotal - $elapsedSeconds))
        
        # Update UI with detailed progress
        Update-UIThreadSafe -Control $ProgressBar -UpdateAction {
            $ProgressBar.Value = [Math]::Min(100, $percentComplete)
        }
        
        Update-UIThreadSafe -Control $StatusLabel -UpdateAction {
            # Enhanced status with percentage and time estimate
            $statusText = "$OperationType $completed av $totalJobs ($percentComplete%)"
            
            if ($remainingSeconds -gt 0 -and $completed -gt 5) {
                if ($remainingSeconds -lt 60) {
                    $statusText += " - ca. $remainingSeconds sek gjenstår"
                } else {
                    $minutes = [Math]::Ceiling($remainingSeconds / 60)
                    $statusText += " - ca. $minutes min gjenstår"
                }
            }
            
            $StatusLabel.Text = $statusText
        }
        
        # Allow UI to process events
        [System.Windows.Forms.Application]::DoEvents()
        
        # Check if all complete
        if ($completed -eq $totalJobs) {
            break
        }
        
        Start-Sleep -Milliseconds $UpdateIntervalMs
    }
    
    # Collect results
    $collectedCount = 0
    foreach ($job in $Jobs) {
        try {
            # Check if job completed successfully
            if ($job.Handle.IsCompleted) {
                try {
                    $result = $job.PowerShell.EndInvoke($job.Handle)
                    if ($result) {
                        $results += $result
                        $collectedCount++
                    }
                } catch {
                    Write-Log -Message "EndInvoke failed for job" -Level 'WARNING' -Source 'Wait-AsyncJobsWithUI' -Context @{ FilePath = $job.FilePath } -Exception $_
                    
                    # Check PowerShell streams for errors
                    if ($job.PowerShell.Streams.Error.Count -gt 0) {
                        foreach ($err in $job.PowerShell.Streams.Error) {
                            Write-Log -Message "Runspace error detected" -Level 'ERROR' -Source 'Wait-AsyncJobsWithUI' -Context @{ FilePath = $job.FilePath; ErrorMessage = $err.Exception.Message }
                        }
                    }
                }
            } else {
                Write-Log -Message "Job not completed" -Level 'WARNING' -Source 'Wait-AsyncJobsWithUI' -Context @{ FilePath = $job.FilePath }
            }
        } catch {
            Write-Log -Message "Critical error in result collection" -Level 'ERROR' -Source 'Wait-AsyncJobsWithUI' -Context @{ FilePath = $job.FilePath } -Exception $_
        } finally {
            # Always dispose, even if errors
            try {
                if ($job.PowerShell) {
                    $job.PowerShell.Dispose()
                }
            } catch {
                Write-Log -Message "Could not dispose PowerShell instance" -Level 'WARNING' -Source 'Wait-AsyncJobsWithUI' -Exception $_
            }
        }
    }
    
    Write-Log -Message "Result collection completed" -Level 'INFO' -Source 'Wait-AsyncJobsWithUI' -Context @{ CollectedCount = $collectedCount; TotalJobs = $Jobs.Count }
    
    return $results
}

function Start-AsyncBatchLabelApplication {
    <#
    .SYNOPSIS
        Applies labels to files asynchronously using runspace pool
    .DESCRIPTION
        Processes label application in parallel for improved performance on large batches
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$FilesToProcess,
        [Parameter(Mandatory=$true)]
        [string]$LabelId,
        [Parameter(Mandatory=$true)]
        [hashtable]$Analysis,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,
        [Parameter(Mandatory=$true)]
        [hashtable]$SharedStats,
        [Parameter(Mandatory=$true)]
        [object]$SelectedLabelObj,
        [string]$Justification = "Endret via massemerking",
        [hashtable]$ProtectionSettings = $null,
        [bool]$RequiresProtection = $false
    )
    
    $labelApplicationScript = {
        param(
            $FilePath,
            $LabelId,
            $ChangeType,
            $OriginalLabel,
            $SelectedLabelObj,
            $Justification,
            $ProtectionSettings,
            $RequiresProtection,
            $SharedStats,
            $NewRank
        )
        
        # === CRITICAL: RUNSPACE-LEVEL ERROR HANDLER ===
        # Runspace crashes kill the entire process
        # We MUST catch EVERYTHING here
        $ErrorActionPreference = 'Continue'  # Don't let errors terminate the runspace!
        
        $result = @{
            FilePath = $FilePath
            Success = $false
            ChangeType = $ChangeType
            OriginalLabel = $OriginalLabel
            NewLabel = $SelectedLabelObj.DisplayName
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Message = ""
            ErrorMessage = $null
        }
        
        try {
            # ULTIMATE WRAPPER: Catch everything in this runspace
        try {
            # CRITICAL: Check if file is locked before attempting label application
            # This prevents crashes when files are open in Office applications
            try {
                $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
                $fileStream.Close()
            } catch [System.IO.IOException] {
                if ($_.Exception.Message -like "*being used by another process*" -or 
                    $_.Exception.Message -like "*file is in use*") {
                    # File is locked - skip it
                    $result.Success = $false
                    $result.ErrorMessage = "Filen er åpen i et annet program. Lukk filen og prøv igjen."
                    $result.Message = "Hoppet over (låst)"
                    [System.Threading.Interlocked]::Increment([ref]$SharedStats.FailureCount)
                    [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
                    return $result
                } else {
                    throw  # Other IO errors should be handled normally
                }
            }
            
            # For protected labels with custom permissions
            if ($RequiresProtection -and $ProtectionSettings) {
                $permissionType = $ProtectionSettings.PermissionType
                
                # Get current user's email
                $currentUserEmail = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                try {
                    $currentUserEmail = ([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").mail
                    if (-not $currentUserEmail) {
                        $username = $env:USERNAME
                        $domain = $env:USERDNSDOMAIN
                        if ($domain) {
                            $currentUserEmail = "$username@$domain".ToLower()
                        }
                    }
                } catch {
                    $currentUserEmail = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                }
                
                # Determine users and permissions based on selection
                if ($permissionType -eq 4) {
                    # "Bare for meg" - Owner only
                    $userList = @($currentUserEmail)
                    $permissionLevel = "CoOwner"
                    $permDesc = "bare for meg ($currentUserEmail)"
                } else {
                    # Other options require email input
                    if (-not $ProtectionSettings.Emails) {
                        throw "Ingen brukere angitt for valgt tillatelse"
                    }
                    
                    $userList = $ProtectionSettings.Emails -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
                    
                    switch ($permissionType) {
                        0 { $permissionLevel = "Viewer"; $permDesc = "leser" }
                        1 { $permissionLevel = "Reviewer"; $permDesc = "kontrollør" }
                        2 { $permissionLevel = "CoAuthor"; $permDesc = "medforfatter" }
                        3 { $permissionLevel = "CoOwner"; $permDesc = "medeier" }
                    }
                }
                
                # Create custom permissions
                try {
                    $customPermission = New-AIPCustomPermissions -Users $userList -Permissions $permissionLevel
                    if (-not $customPermission) {
                        throw "Failed to create custom permissions"
                    }
                    
                    # Apply label with custom protection
                    Set-AIPFileLabel -Path $FilePath -LabelId $LabelId -CustomPermissions $customPermission -PreserveFileDetails
                    $result.Success = $true
                    $result.Message = "Beskyttelse: $permDesc"
                } catch {
                    # Detailed error capture for protection-related failures
                    $result.Success = $false
                    $result.ErrorMessage = $_.Exception.Message
                    $result.ErrorType = $_.Exception.GetType().Name
                    [System.Threading.Interlocked]::Increment([ref]$SharedStats.FailureCount)
                    [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
                    return $result
                }
                
            } else {
                # Normal label application without custom protection
                # Check if this file needs justification (downgrade case)
                $fileNeedsJustification = ($ChangeType -eq "Downgrade")
                
                try {
                    if ($fileNeedsJustification -and $Justification) {
                        # Apply with justification proactively for downgrades
                        Set-AIPFileLabel -Path $FilePath -LabelId $LabelId -JustificationMessage $Justification -PreserveFileDetails
                        $result.Success = $true
                        $result.Message = "Med begrunnelse"
                    } else {
                        # Apply without justification
                        Set-AIPFileLabel -Path $FilePath -LabelId $LabelId -PreserveFileDetails
                        $result.Success = $true
                    }
                } catch {
                    $errorMessage = $_.Exception.Message
                    
                    # Handle justification requirement (fallback if detection missed it)
                    if ($errorMessage -like "*Justification*" -and $Justification) {
                        try {
                            Set-AIPFileLabel -Path $FilePath -LabelId $LabelId -JustificationMessage $Justification -PreserveFileDetails
                            $result.Success = $true
                            $result.Message = "Med begrunnelse (retry)"
                        } catch {
                            # Even retry failed
                            $result.Success = $false
                            $result.ErrorMessage = "Justification retry failed: $($_.Exception.Message)"
                            $result.ErrorType = $_.Exception.GetType().Name
                            [System.Threading.Interlocked]::Increment([ref]$SharedStats.FailureCount)
                            [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
                            return $result
                        }
                    } else {
                        # Not a justification error, record and continue
                        $result.Success = $false
                        $result.ErrorMessage = $errorMessage
                        $result.ErrorType = $_.Exception.GetType().Name
                        [System.Threading.Interlocked]::Increment([ref]$SharedStats.FailureCount)
                        [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
                        return $result
                    }
                }
            }
            
            # Update statistics (thread-safe)
            [System.Threading.Interlocked]::Increment([ref]$SharedStats.SuccessCount)
            [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
            
            # Update change type breakdown
            switch ($ChangeType) {
                "New" { [System.Threading.Interlocked]::Increment([ref]$SharedStats.ChangeTypeBreakdown_New) }
                "Upgrade" { [System.Threading.Interlocked]::Increment([ref]$SharedStats.ChangeTypeBreakdown_Upgrade) }
                "Downgrade" { [System.Threading.Interlocked]::Increment([ref]$SharedStats.ChangeTypeBreakdown_Downgrade) }
                "Unchanged" { [System.Threading.Interlocked]::Increment([ref]$SharedStats.ChangeTypeBreakdown_Unchanged) }
                "Same" { [System.Threading.Interlocked]::Increment([ref]$SharedStats.ChangeTypeBreakdown_Same) }
            }
            
        } catch {
            $result.Success = $false
            $result.ErrorMessage = $_.Exception.Message
            $result.ErrorType = $_.Exception.GetType().Name
            [System.Threading.Interlocked]::Increment([ref]$SharedStats.FailureCount)
            [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
        }
        
        } catch {
            # === ULTIMATE RUNSPACE CATCH ===
            # Catches ANYTHING that escaped all other handlers
            # This prevents runspace termination which kills the entire process
            
            # Ensure result has all required fields populated
            if (-not $result.FilePath) { $result.FilePath = $FilePath }
            if (-not $result.OriginalLabel) { $result.OriginalLabel = $OriginalLabel }
            if (-not $result.NewLabel) { $result.NewLabel = $SelectedLabelObj.DisplayName }
            if (-not $result.ChangeType) { $result.ChangeType = $ChangeType }
            if (-not $result.Timestamp) { $result.Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
            
            $result.Success = $false
            $result.ErrorMessage = "RUNSPACE CRASH: $($_.Exception.Message)"
            $result.ErrorType = $_.Exception.GetType().FullName
            
            [System.Threading.Interlocked]::Increment([ref]$SharedStats.FailureCount)
            [System.Threading.Interlocked]::Increment([ref]$SharedStats.TotalProcessed)
        }
        
        return $result
    }
    
    # Start jobs for all files
    $jobs = @()
    
    foreach ($file in $FilesToProcess) {
        # Determine change type and original label for this file from analysis
        $changeType = "Unknown"
        $originalLabel = "Ukjent"
        
        $fileAnalysis = $Analysis.New | Where-Object { $_.File -eq $file }
        if ($fileAnalysis) {
            $changeType = "New"
            $originalLabel = $fileAnalysis.CurrentLabel
        }
        else {
            $fileAnalysis = $Analysis.Upgrade | Where-Object { $_.File -eq $file }
            if ($fileAnalysis) {
                $changeType = "Upgrade"
                $originalLabel = $fileAnalysis.CurrentLabel
            }
            else {
                $fileAnalysis = $Analysis.Downgrade | Where-Object { $_.File -eq $file }
                if ($fileAnalysis) {
                    $changeType = "Downgrade"
                    $originalLabel = $fileAnalysis.CurrentLabel
                }
                else {
                    $fileAnalysis = $Analysis.Unchanged | Where-Object { $_.File -eq $file }
                    if ($fileAnalysis) {
                        $changeType = "Unchanged"
                        $originalLabel = $fileAnalysis.CurrentLabel
                    }
                    else {
                        $fileAnalysis = $Analysis.Same | Where-Object { $_.File -eq $file }
                        if ($fileAnalysis) {
                            $changeType = "Same"
                            $originalLabel = $fileAnalysis.CurrentLabel
                        }
                    }
                }
            }
        }
        
        # Calculate NewRank before adding parameter
        $newRankValue = if($SelectedLabelObj.Rank) { $SelectedLabelObj.Rank } else { 0 }
        
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $RunspacePool
        [void]$ps.AddScript($labelApplicationScript)
        [void]$ps.AddParameter("FilePath", $file)
        [void]$ps.AddParameter("LabelId", $LabelId)
        [void]$ps.AddParameter("ChangeType", $changeType)
        [void]$ps.AddParameter("OriginalLabel", $originalLabel)
        [void]$ps.AddParameter("SelectedLabelObj", $SelectedLabelObj)
        [void]$ps.AddParameter("Justification", $Justification)
        [void]$ps.AddParameter("ProtectionSettings", $ProtectionSettings)
        [void]$ps.AddParameter("RequiresProtection", $RequiresProtection)
        [void]$ps.AddParameter("SharedStats", $SharedStats)
        [void]$ps.AddParameter("NewRank", $newRankValue)
        
        $handle = $ps.BeginInvoke()
        
        $jobs += [PSCustomObject]@{
            PowerShell = $ps
            Handle = $handle
            FilePath = $file
        }
    }
    
    return $jobs
}

function Get-DefaultAppConfig {
    <#
    .SYNOPSIS
        Returns default application configuration
    .DESCRIPTION
        Provides default values for all configuration options
    #>
    return @{
        version = "1.0"
        preferences = @{
            defaultFolder = ""
            rememberLastLabel = $false
            includeSubfoldersDefault = $false
            lastSelectedLabelId = $null
            language = "en-US"  # Default to English for international audience
        }
        warnings = @{
            showPreApplySummary = $true
            showMassDowngradeWarning = $true
            showLargeBatchWarning = $true
            showProtectionRequiredWarning = $true
            showNoChangesWarning = $true
            largeBatchThreshold = 20
            massDowngradeThreshold = 3
        }
        export = @{
            defaultExportFormat = "csv"
            includeTimestamps = $true
            includeSummary = $true
        }
        logging = @{
            logRetentionDays = 30
            logDirectory = ""
            enableDetailedLogging = $true
        }
        ui = @{
            windowWidth = 740
            windowHeight = 570
            rememberWindowPosition = $false
            windowPositionX = $null
            windowPositionY = $null
        }
    }
}

function Merge-AppConfig {
    <#
    .SYNOPSIS
        Merges loaded config with defaults to fill missing keys
    .DESCRIPTION
        Recursively merges loaded configuration with defaults,
        ensuring all required keys exist even if missing from file
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$LoadedConfig,
        [Parameter(Mandatory=$true)]
        [hashtable]$DefaultConfig
    )
    
    $merged = $DefaultConfig.Clone()
    
    foreach ($key in $LoadedConfig.Keys) {
        if ($merged.ContainsKey($key)) {
            $loadedValue = $LoadedConfig[$key]
            $defaultValue = $merged[$key]
            
            # Recursive merge for nested hashtables
            if ($loadedValue -is [hashtable] -and $defaultValue -is [hashtable]) {
                $merged[$key] = Merge-AppConfig -LoadedConfig $loadedValue -DefaultConfig $defaultValue
            }
            elseif ($loadedValue -is [System.Management.Automation.PSCustomObject] -and $defaultValue -is [hashtable]) {
                # Convert PSCustomObject to hashtable for merging
                $loadedHash = @{}
                $loadedValue.PSObject.Properties | ForEach-Object {
                    $loadedHash[$_.Name] = $_.Value
                }
                $merged[$key] = Merge-AppConfig -LoadedConfig $loadedHash -DefaultConfig $defaultValue
            }
            else {
                # Use loaded value if type matches or is compatible
                $merged[$key] = $loadedValue
            }
        }
    }
    
    return $merged
}

function Validate-AppConfig {
    <#
    .SYNOPSIS
        Validates application configuration structure and values
    .DESCRIPTION
        Checks for required keys, valid data types, and value ranges
    .PARAMETER Config
        Configuration hashtable to validate
    .OUTPUTS
        Hashtable with IsValid (bool) and Errors (array) keys
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    $result = @{
        IsValid = $true
        Errors = @()
        Warnings = @()
    }
    
    # Check version
    if (-not $Config.ContainsKey('version') -or [string]::IsNullOrWhiteSpace($Config.version)) {
        $result.Errors += "Missing or empty 'version' field"
        $result.IsValid = $false
    }
    
    # Validate preferences section
    if ($Config.ContainsKey('preferences')) {
        $prefs = $Config.preferences
        
        # Validate defaultFolder (string)
        if ($prefs.ContainsKey('defaultFolder') -and $prefs.defaultFolder -ne "" -and -not (Test-Path $prefs.defaultFolder -IsValid)) {
            $result.Warnings += "preferences.defaultFolder contains invalid path format"
        }
        
        # Validate boolean fields
        foreach ($boolField in @('rememberLastLabel', 'includeSubfoldersDefault')) {
            if ($prefs.ContainsKey($boolField) -and $prefs[$boolField] -isnot [bool]) {
                $result.Errors += "preferences.$boolField must be boolean (true/false)"
                $result.IsValid = $false
            }
        }
        
        # Validate lastSelectedLabelId (GUID or null)
        if ($prefs.ContainsKey('lastSelectedLabelId') -and $prefs.lastSelectedLabelId -ne $null) {
            try {
                [guid]::Parse($prefs.lastSelectedLabelId) | Out-Null
            }
            catch {
                $result.Warnings += "preferences.lastSelectedLabelId is not a valid GUID"
            }
        }
    }
    
    # Validate warnings section
    if ($Config.ContainsKey('warnings')) {
        $warnings = $Config.warnings
        
        # Validate boolean fields
        foreach ($boolField in @('showPreApplySummary', 'showMassDowngradeWarning', 'showLargeBatchWarning', 'showProtectionRequiredWarning', 'showNoChangesWarning')) {
            if ($warnings.ContainsKey($boolField) -and $warnings[$boolField] -isnot [bool]) {
                $result.Errors += "warnings.$boolField must be boolean (true/false)"
                $result.IsValid = $false
            }
        }
        
        # Validate numeric thresholds
        if ($warnings.ContainsKey('largeBatchThreshold')) {
            $val = $warnings.largeBatchThreshold
            if ($val -isnot [int] -or $val -lt 1 -or $val -gt 1000) {
                $result.Errors += "warnings.largeBatchThreshold must be integer between 1-1000"
                $result.IsValid = $false
            }
        }
        
        if ($warnings.ContainsKey('massDowngradeThreshold')) {
            $val = $warnings.massDowngradeThreshold
            if ($val -isnot [int] -or $val -lt 1 -or $val -gt 100) {
                $result.Errors += "warnings.massDowngradeThreshold must be integer between 1-100"
                $result.IsValid = $false
            }
        }
    }
    
    # Validate export section
    if ($Config.ContainsKey('export')) {
        $export = $Config.export
        
        # Validate format
        if ($export.ContainsKey('defaultExportFormat') -and $export.defaultExportFormat -ne 'csv') {
            $result.Warnings += "export.defaultExportFormat currently only supports 'csv'"
        }
        
        # Validate boolean fields
        foreach ($boolField in @('includeTimestamps', 'includeSummary')) {
            if ($export.ContainsKey($boolField) -and $export[$boolField] -isnot [bool]) {
                $result.Errors += "export.$boolField must be boolean (true/false)"
                $result.IsValid = $false
            }
        }
    }
    
    # Validate logging section
    if ($Config.ContainsKey('logging')) {
        $logging = $Config.logging
        
        # Validate logRetentionDays
        if ($logging.ContainsKey('logRetentionDays')) {
            $val = $logging.logRetentionDays
            if ($val -isnot [int] -or $val -lt 0) {
                $result.Errors += "logging.logRetentionDays must be integer >= 0 (0 = no cleanup)"
                $result.IsValid = $false
            }
        }
        
        # Validate logDirectory
        if ($logging.ContainsKey('logDirectory') -and $logging.logDirectory -ne "" -and -not (Test-Path $logging.logDirectory -IsValid)) {
            $result.Warnings += "logging.logDirectory contains invalid path format"
        }
        
        # Validate enableDetailedLogging
        if ($logging.ContainsKey('enableDetailedLogging') -and $logging.enableDetailedLogging -isnot [bool]) {
            $result.Errors += "logging.enableDetailedLogging must be boolean (true/false)"
            $result.IsValid = $false
        }
    }
    
    # Validate UI section
    if ($Config.ContainsKey('ui')) {
        $ui = $Config.ui
        
        # Validate window dimensions
        if ($ui.ContainsKey('windowWidth')) {
            $val = $ui.windowWidth
            if ($val -isnot [int] -or $val -lt 600 -or $val -gt 2000) {
                $result.Errors += "ui.windowWidth must be integer between 600-2000 pixels"
                $result.IsValid = $false
            }
        }
        
        if ($ui.ContainsKey('windowHeight')) {
            $val = $ui.windowHeight
            if ($val -isnot [int] -or $val -lt 400 -or $val -gt 1500) {
                $result.Errors += "ui.windowHeight must be integer between 400-1500 pixels"
                $result.IsValid = $false
            }
        }
        
        # Validate rememberWindowPosition
        if ($ui.ContainsKey('rememberWindowPosition') -and $ui.rememberWindowPosition -isnot [bool]) {
            $result.Errors += "ui.rememberWindowPosition must be boolean (true/false)"
            $result.IsValid = $false
        }
        
        # Validate window position (if present, must be integers)
        foreach ($posField in @('windowPositionX', 'windowPositionY')) {
            if ($ui.ContainsKey($posField) -and $ui[$posField] -ne $null) {
                if ($ui[$posField] -isnot [int]) {
                    $result.Errors += "ui.$posField must be integer or null"
                    $result.IsValid = $false
                }
            }
        }
    }
    
    return $result
}

function Load-AppConfig {
    <#
    .SYNOPSIS
        Loads application configuration from file
    .DESCRIPTION
        Attempts to load app_config.json, validates structure,
        merges with defaults, and returns configuration object
    .OUTPUTS
        Hashtable containing application configuration
    #>
    $configPath = Join-Path $PSScriptRoot "app_config.json"
    $defaultConfig = Get-DefaultAppConfig
    
    # Check if config file exists
    if (-not (Test-Path $configPath)) {
        Write-Log -Message "Config file not found, creating default configuration" -Level 'INFO' -Source 'Load-AppConfig' -Context @{ ConfigPath = $configPath }
        
        try {
            # Create new config file with defaults
            $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
            Write-Log -Message "Default config file created successfully" -Level 'INFO' -Source 'Load-AppConfig'
        }
        catch {
            Write-Log -Message "Failed to create default config file" -Level 'WARNING' -Source 'Load-AppConfig' -Exception $_
        }
        
        return $defaultConfig
    }
    
    # Try to load existing config
    try {
        $loadedJson = Get-Content -Path $configPath -Raw -ErrorAction Stop
        $loadedConfig = $loadedJson | ConvertFrom-Json -ErrorAction Stop
        
        # Convert PSCustomObject to hashtable
        $loadedHash = @{}
        $loadedConfig.PSObject.Properties | ForEach-Object {
            $loadedHash[$_.Name] = $_.Value
        }
        
        # Merge with defaults to fill any missing keys
        $mergedConfig = Merge-AppConfig -LoadedConfig $loadedHash -DefaultConfig $defaultConfig
        
        Write-Log -Message "Config loaded successfully" -Level 'INFO' -Source 'Load-AppConfig' -Context @{ ConfigPath = $configPath }
        
        # Validate configuration structure and values
        $validation = Validate-AppConfig -Config $mergedConfig
        
        if (-not $validation.IsValid) {
            Write-Log -Message "Config validation failed" -Level 'ERROR' -Source 'Load-AppConfig'
            foreach ($error in $validation.Errors) {
                Write-Log -Message "Validation error: $error" -Level 'ERROR' -Source 'Load-AppConfig'
            }
            Write-Log -Message "Using default configuration due to validation errors" -Level 'WARNING' -Source 'Load-AppConfig'
            
            # Backup invalid config
            try {
                $backupPath = "$configPath.invalid_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Copy-Item -Path $configPath -Destination $backupPath -ErrorAction Stop
                Write-Log -Message "Invalid config backed up" -Level 'INFO' -Source 'Load-AppConfig' -Context @{ BackupPath = $backupPath }
            }
            catch {
                Write-Log -Message "Could not backup invalid config" -Level 'WARNING' -Source 'Load-AppConfig' -Exception $_
            }
            
            return $defaultConfig
        }
        
        # Log warnings (non-fatal issues)
        if ($validation.Warnings.Count -gt 0) {
            foreach ($warning in $validation.Warnings) {
                Write-Log -Message "Config validation warning: $warning" -Level 'WARNING' -Source 'Load-AppConfig'
            }
        }
        
        # Validate version
        if ($mergedConfig.version -ne $defaultConfig.version) {
            Write-Log -Message "Config version mismatch" -Level 'WARNING' -Source 'Load-AppConfig' -Context @{ LoadedVersion = $mergedConfig.version; ExpectedVersion = $defaultConfig.version }
            # Could implement migration logic here in future
        }
        
        return $mergedConfig
    }
    catch {
        Write-Log -Message "Failed to load config file" -Level 'ERROR' -Source 'Load-AppConfig' -Exception $_
        Write-Log -Message "Using default configuration" -Level 'INFO' -Source 'Load-AppConfig'
        
        # Try to backup corrupted file
        try {
            $backupPath = "$configPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            Copy-Item -Path $configPath -Destination $backupPath -ErrorAction Stop
            Write-Log -Message "Corrupted config backed up" -Level 'INFO' -Source 'Load-AppConfig' -Context @{ BackupPath = $backupPath }
        }
        catch {
            Write-Log -Message "Could not backup corrupted config" -Level 'WARNING' -Source 'Load-AppConfig' -Exception $_
        }
        
        # Try to create new default config
        try {
            $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
            Write-Log -Message "New default config created" -Level 'INFO' -Source 'Load-AppConfig'
        }
        catch {
            Write-Log -Message "Failed to create new config file" -Level 'WARNING' -Source 'Load-AppConfig' -Exception $_
        }
        
        return $defaultConfig
    }
}

function Save-AppConfig {
    <#
    .SYNOPSIS
        Saves application configuration to file
    .DESCRIPTION
        Writes configuration to app_config.json using atomic write
        (write to temp file, then rename)
    .PARAMETER Config
        Hashtable containing configuration to save
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    $configPath = Join-Path $PSScriptRoot "app_config.json"
    $tempPath = "$configPath.tmp"
    
    try {
        # Write to temporary file first (atomic write)
        $Config | ConvertTo-Json -Depth 10 | Set-Content -Path $tempPath -Encoding UTF8 -ErrorAction Stop
        
        # Replace original file with temp file
        if (Test-Path $configPath) {
            Remove-Item -Path $configPath -Force -ErrorAction Stop
        }
        Move-Item -Path $tempPath -Destination $configPath -Force -ErrorAction Stop
        
        Write-Log -Message "Config saved successfully" -Level 'INFO' -Source 'Save-AppConfig' -Context @{ ConfigPath = $configPath }
        return $true
    }
    catch {
        Write-Log -Message "Failed to save config" -Level 'ERROR' -Source 'Save-AppConfig' -Exception $_
        
        # Clean up temp file if it exists
        if (Test-Path $tempPath) {
            try {
                Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
            }
            catch {}
        }
        
        return $false
    }
}

function Reset-AppConfig {
    <#
    .SYNOPSIS
        Resets application configuration to defaults
    .DESCRIPTION
        Deletes existing config file and creates new one with defaults.
        Useful for recovering from configuration corruption or resetting preferences.
    .PARAMETER Confirm
        If true, prompts for confirmation before reset
    .OUTPUTS
        Boolean indicating success
    #>
    param(
        [Parameter(Mandatory=$false)]
        [bool]$Confirm = $true
    )
    
    $configPath = Join-Path $PSScriptRoot "app_config.json"
    
    if ($Confirm) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Dette vil tilbakestille alle innstillinger til standardverdier.`n`nAlle personlige preferanser vil gå tapt.`n`nEr du sikker?",
            "Tilbakestill konfigurasjon",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            Write-Log -Message "Config reset cancelled by user" -Level 'INFO' -Source 'Reset-AppConfig'
            return $false
        }
    }
    
    try {
        # Backup existing config before reset
        if (Test-Path $configPath) {
            $backupPath = "$configPath.before_reset_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            Copy-Item -Path $configPath -Destination $backupPath -ErrorAction Stop
            Write-Log -Message "Config backed up before reset" -Level 'INFO' -Source 'Reset-AppConfig' -Context @{ BackupPath = $backupPath }
            
            # Delete existing config
            Remove-Item -Path $configPath -Force -ErrorAction Stop
            Write-Log -Message "Existing config deleted" -Level 'INFO' -Source 'Reset-AppConfig'
        }
        
        # Create new default config
        $defaultConfig = Get-DefaultAppConfig
        $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8 -ErrorAction Stop
        Write-Log -Message "Config reset to defaults successfully" -Level 'INFO' -Source 'Reset-AppConfig'
        
        [System.Windows.Forms.MessageBox]::Show(
            "Konfigurasjon tilbakestilt til standardverdier.`n`nEndringene vil tre i kraft neste gang applikasjonen startes.",
            "Tilbakestilling vellykket",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        return $true
    }
    catch {
        Write-Log -Message "Failed to reset config" -Level 'ERROR' -Source 'Reset-AppConfig' -Exception $_
        
        [System.Windows.Forms.MessageBox]::Show(
            "Kunne ikke tilbakestille konfigurasjon:`n`n$($_.Exception.Message)",
            "Feil",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $false
    }
}

function Cleanup-OldLogs {
    <#
    .SYNOPSIS
        Removes log files older than retention period
    .DESCRIPTION
        Deletes log files based on logRetentionDays setting
    .PARAMETER RetentionDays
        Number of days to keep logs (0 = no cleanup)
    .PARAMETER LogDir
        Directory containing log files
    #>
    param(
        [Parameter(Mandatory=$true)]
        [int]$RetentionDays,
        [Parameter(Mandatory=$true)]
        [string]$LogDir
    )
    
    if ($RetentionDays -le 0) {
        return  # Cleanup disabled
    }
    
    if (-not (Test-Path $LogDir)) {
        return  # Log directory doesn't exist
    }
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
        $oldLogs = Get-ChildItem -Path $LogDir -Filter "FileLabeler_Log_*.txt" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($oldLogs.Count -gt 0) {
            $oldLogs | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Log -Message "Old log files cleaned up" -Level 'INFO' -Source 'Cleanup-OldLogs' -Context @{ FilesRemoved = $oldLogs.Count; RetentionDays = $RetentionDays }
        }
    }
    catch {
        Write-Log -Message "Log cleanup failed" -Level 'WARNING' -Source 'Cleanup-OldLogs' -Exception $_
    }
}

# ========================================
# LANGUAGE RESOURCE MANAGEMENT
# ========================================

<#
.SYNOPSIS
    Get the appropriate language code based on system language or config setting
.DESCRIPTION
    Determines which language to use: auto-detect system language or use manual override
.PARAMETER ConfigLanguage
    Language setting from app_config.json (can be "auto", "nb-NO", or "en-US")
.OUTPUTS
    String - Language code ("nb-NO" or "en-US")
#>
function Get-SystemLanguage {
    param(
        [Parameter(Mandatory=$false)]
        [string]$ConfigLanguage = "auto"
    )
    
    try {
        # If manual override is set, use it
        if ($ConfigLanguage -and $ConfigLanguage -ne "auto") {
            Write-Log -Message "Using manual language override: $ConfigLanguage" -Level 'INFO' -Source 'Get-SystemLanguage'
            return $ConfigLanguage
        }
        
        # Auto-detect system language
        $systemLang = (Get-Culture).Name
        Write-Log -Message "Detected system language: $systemLang" -Level 'INFO' -Source 'Get-SystemLanguage'
        
        # Check if Norwegian (nb-NO, no, nn-NO, nb)
        if ($systemLang -like "nb-*" -or $systemLang -eq "no" -or $systemLang -like "nn-*" -or $systemLang -eq "nb" -or $systemLang -eq "nn") {
            Write-Log -Message "System language is Norwegian, using nb-NO" -Level 'INFO' -Source 'Get-SystemLanguage'
            return "nb-NO"
        }
        
        # Default to English for all other languages
        Write-Log -Message "System language is not Norwegian, using en-US" -Level 'INFO' -Source 'Get-SystemLanguage'
        return "en-US"
    }
    catch {
        Write-Log -Message "Failed to detect system language, defaulting to English" -Level 'WARNING' -Source 'Get-SystemLanguage' -Exception $_
        return "en-US"  # Default to English on error
    }
}

<#
.SYNOPSIS
    Load language resource file
.DESCRIPTION
    Loads and parses the appropriate JSON language resource file with UTF-8 BOM support
.PARAMETER LanguageCode
    Language code ("nb-NO" or "en-US")
.OUTPUTS
    PSCustomObject - Parsed language resources
#>
function Load-LanguageResources {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("nb-NO", "en-US")]
        [string]$LanguageCode
    )
    
    try {
        # Determine file name
        $fileName = if ($LanguageCode -eq "nb-NO") { "Norwegian.json" } else { "English.json" }
        $resourcePath = Join-Path $PSScriptRoot "resources\$fileName"
        
        Write-Log -Message "Loading language resources from: $resourcePath" -Level 'INFO' -Source 'Load-LanguageResources' -Context @{ LanguageCode = $LanguageCode }
        
        # Check if file exists
        if (-not (Test-Path $resourcePath)) {
            Write-Log -Message "Language file not found: $resourcePath" -Level 'ERROR' -Source 'Load-LanguageResources'
            return $null
        }
        
        # Read file with UTF-8 BOM encoding
        $jsonContent = [System.IO.File]::ReadAllText($resourcePath, [System.Text.Encoding]::UTF8)
        
        # Parse JSON
        $resources = $jsonContent | ConvertFrom-Json
        
        Write-Log -Message "Language resources loaded successfully" -Level 'INFO' -Source 'Load-LanguageResources' -Context @{ LanguageCode = $LanguageCode; FileName = $fileName }
        
        return $resources
    }
    catch {
        Write-Log -Message "Failed to load language resources" -Level 'ERROR' -Source 'Load-LanguageResources' -Exception $_ -Context @{ LanguageCode = $LanguageCode }
        return $null
    }
}

<#
.SYNOPSIS
    Get localized string by key
.DESCRIPTION
    Retrieves a localized string from loaded resources with fallback to Norwegian
.PARAMETER Key
    Dot-notation key (e.g., "buttons.apply", "dialogs.resetConfig.title")
.PARAMETER Parameters
    Optional array of parameters for string formatting ({0}, {1}, etc.)
.OUTPUTS
    String - Localized text or key name if not found
#>
function Get-LocalizedString {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Key,
        
        [Parameter(Mandatory=$false)]
        [object[]]$Parameters = @()
    )
    
    try {
        # Navigate nested structure
        $keys = $Key -split '\.'
        $value = $script:LanguageResources
        
        foreach ($k in $keys) {
            if ($value -and ($value.PSObject.Properties.Name -contains $k)) {
                $value = $value.$k
            }
            else {
                # Key not found in current language, try fallback to Norwegian
                if ($script:CurrentLanguage -ne "nb-NO" -and $script:NorwegianFallbackResources) {
                    Write-Log -Message "Key not found in $($script:CurrentLanguage), falling back to Norwegian: $Key" -Level 'WARNING' -Source 'Get-LocalizedString'
                    
                    $fallbackValue = $script:NorwegianFallbackResources
                    foreach ($fk in $keys) {
                        if ($fallbackValue -and ($fallbackValue.PSObject.Properties.Name -contains $fk)) {
                            $fallbackValue = $fallbackValue.$fk
                        }
                        else {
                            $fallbackValue = $null
                            break
                        }
                    }
                    
                    if ($fallbackValue) {
                        $value = $fallbackValue
                        break
                    }
                }
                
                # Key not found in either language
                Write-Log -Message "Localization key not found: $Key" -Level 'WARNING' -Source 'Get-LocalizedString'
                return "[$Key]"  # Return key in brackets for debugging
            }
        }
        
        # If value is still an object (not a leaf string), return key
        if ($value -is [PSCustomObject]) {
            Write-Log -Message "Key points to object, not string: $Key" -Level 'WARNING' -Source 'Get-LocalizedString'
            return "[$Key]"
        }
        
        # Apply parameters if provided
        if ($Parameters -and $Parameters.Count -gt 0) {
            try {
                $value = $value -f $Parameters
            }
            catch {
                Write-Log -Message "Failed to format string with parameters" -Level 'WARNING' -Source 'Get-LocalizedString' -Context @{ Key = $Key; ParameterCount = $Parameters.Count } -Exception $_
            }
        }
        
        return $value
    }
    catch {
        Write-Log -Message "Error retrieving localized string" -Level 'ERROR' -Source 'Get-LocalizedString' -Exception $_ -Context @{ Key = $Key }
        return "[$Key]"
    }
}

<#
.SYNOPSIS
    Initialize language resources
.DESCRIPTION
    Detects system language, loads appropriate resources, and sets up fallback
.PARAMETER Config
    Application configuration hashtable
.OUTPUTS
    Boolean - True if successful, False otherwise
#>
function Initialize-LanguageResources {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    try {
        # Determine language
        $configLang = if ($Config.preferences.language) { $Config.preferences.language } else { "auto" }
        $script:CurrentLanguage = Get-SystemLanguage -ConfigLanguage $configLang
        
        Write-Log -Message "Initializing language resources" -Level 'INFO' -Source 'Initialize-LanguageResources' -Context @{ CurrentLanguage = $script:CurrentLanguage; ConfigLanguage = $configLang }
        
        # Load main language resources
        $script:LanguageResources = Load-LanguageResources -LanguageCode $script:CurrentLanguage
        
        if (-not $script:LanguageResources) {
            Write-Log -Message "Failed to load main language resources, attempting Norwegian fallback" -Level 'ERROR' -Source 'Initialize-LanguageResources'
            
            # Try Norwegian as fallback
            $script:LanguageResources = Load-LanguageResources -LanguageCode "nb-NO"
            $script:CurrentLanguage = "nb-NO"
            
            if (-not $script:LanguageResources) {
                Write-Log -Message "Failed to load any language resources" -Level 'CRITICAL' -Source 'Initialize-LanguageResources'
                return $false
            }
        }
        
        # Always load Norwegian as fallback (for missing keys in other languages)
        if ($script:CurrentLanguage -ne "nb-NO") {
            $script:NorwegianFallbackResources = Load-LanguageResources -LanguageCode "nb-NO"
            if (-not $script:NorwegianFallbackResources) {
                Write-Log -Message "Failed to load Norwegian fallback resources" -Level 'WARNING' -Source 'Initialize-LanguageResources'
            }
        }
        else {
            $script:NorwegianFallbackResources = $null  # No fallback needed if Norwegian is primary
        }
        
        Write-Log -Message "Language resources initialized successfully" -Level 'INFO' -Source 'Initialize-LanguageResources' -Context @{ CurrentLanguage = $script:CurrentLanguage }
        return $true
    }
    catch {
        Write-Log -Message "Critical error initializing language resources" -Level 'CRITICAL' -Source 'Initialize-LanguageResources' -Exception $_
        return $false
    }
}

# Global variables for language resources
$script:LanguageResources = $null
$script:NorwegianFallbackResources = $null
$script:CurrentLanguage = "en-US"  # Default to English for international audience

function Show-SettingsDialog {
    <#
    .SYNOPSIS
        Displays configuration settings dialog
    .DESCRIPTION
        Modal dialog for editing application configuration
    .PARAMETER Config
        Current configuration hashtable
    .OUTPUTS
        Boolean indicating if settings were changed
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    # Create settings form
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = Get-LocalizedString -Key "settings.title"
    $settingsForm.Size = New-Object System.Drawing.Size(600, 650)
    $settingsForm.StartPosition = 'CenterParent'
    $settingsForm.FormBorderStyle = 'FixedDialog'
    $settingsForm.MaximizeBox = $false
    $settingsForm.MinimizeBox = $false
    $settingsForm.TopMost = $true
    $settingsForm.Owner = $form
    $settingsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create tab control for organized settings
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(565, 540)
    $settingsForm.Controls.Add($tabControl)
    
    # ====================
    # TAB 1: PREFERENCES
    # ====================
    $prefTab = New-Object System.Windows.Forms.TabPage
    $prefTab.Text = Get-LocalizedString -Key "settings.tabs.preferences"
    $tabControl.TabPages.Add($prefTab)
    
    $yPos = 20
    
    # Remember last label checkbox
    $rememberLabelCheck = New-Object System.Windows.Forms.CheckBox
    $rememberLabelCheck.Text = Get-LocalizedString -Key "settings.preferences.rememberLabel"
    $rememberLabelCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $rememberLabelCheck.Size = New-Object System.Drawing.Size(520, 20)
    $rememberLabelCheck.Checked = $Config.preferences.rememberLastLabel
    $prefTab.Controls.Add($rememberLabelCheck)
    
    $yPos += 30
    
    # Include subfolders default checkbox
    $subfoldersCheck = New-Object System.Windows.Forms.CheckBox
    $subfoldersCheck.Text = Get-LocalizedString -Key "settings.preferences.includeSubfolders"
    $subfoldersCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $subfoldersCheck.Size = New-Object System.Drawing.Size(520, 20)
    $subfoldersCheck.Checked = $Config.preferences.includeSubfoldersDefault
    $prefTab.Controls.Add($subfoldersCheck)
    
    $yPos += 40
    
    # Default folder label
    $defaultFolderLabel = New-Object System.Windows.Forms.Label
    $defaultFolderLabel.Text = Get-LocalizedString -Key "settings.preferences.defaultFolder"
    $defaultFolderLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $defaultFolderLabel.Size = New-Object System.Drawing.Size(520, 20)
    $prefTab.Controls.Add($defaultFolderLabel)
    
    $yPos += 25
    
    # Default folder textbox
    $defaultFolderText = New-Object System.Windows.Forms.TextBox
    $defaultFolderText.Location = New-Object System.Drawing.Point(20, $yPos)
    $defaultFolderText.Size = New-Object System.Drawing.Size(420, 25)
    $defaultFolderText.Text = $Config.preferences.defaultFolder
    $prefTab.Controls.Add($defaultFolderText)
    
    # Browse button for default folder
    $browseFolderBtn = New-Object System.Windows.Forms.Button
    $browseFolderBtn.Text = '...'
    $browseFolderBtn.Location = New-Object System.Drawing.Point(450, $yPos)
    $browseFolderBtn.Size = New-Object System.Drawing.Size(40, 25)
    $browseFolderBtn.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = Get-LocalizedString -Key "dialogs.selectFolder.description"
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $defaultFolderText.Text = $folderBrowser.SelectedPath
        }
    })
    $prefTab.Controls.Add($browseFolderBtn)
    
    # ====================
    # TAB 2: WARNINGS
    # ====================
    $warningsTab = New-Object System.Windows.Forms.TabPage
    $warningsTab.Text = Get-LocalizedString -Key "settings.tabs.warnings"
    $tabControl.TabPages.Add($warningsTab)
    
    $yPos = 20
    
    # Show pre-apply summary checkbox
    $showSummaryCheck = New-Object System.Windows.Forms.CheckBox
    $showSummaryCheck.Text = Get-LocalizedString -Key "settings.warnings.showSummary"
    $showSummaryCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $showSummaryCheck.Size = New-Object System.Drawing.Size(520, 20)
    $showSummaryCheck.Checked = $Config.warnings.showPreApplySummary
    $warningsTab.Controls.Add($showSummaryCheck)
    
    $yPos += 30
    
    # Mass downgrade warning checkbox
    $massDowngradeCheck = New-Object System.Windows.Forms.CheckBox
    $massDowngradeCheck.Text = Get-LocalizedString -Key "settings.warnings.massDowngrade"
    $massDowngradeCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $massDowngradeCheck.Size = New-Object System.Drawing.Size(520, 20)
    $massDowngradeCheck.Checked = $Config.warnings.showMassDowngradeWarning
    $warningsTab.Controls.Add($massDowngradeCheck)
    
    $yPos += 30
    
    # Large batch warning checkbox
    $largeBatchCheck = New-Object System.Windows.Forms.CheckBox
    $largeBatchCheck.Text = Get-LocalizedString -Key "settings.warnings.largeBatch"
    $largeBatchCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $largeBatchCheck.Size = New-Object System.Drawing.Size(520, 20)
    $largeBatchCheck.Checked = $Config.warnings.showLargeBatchWarning
    $warningsTab.Controls.Add($largeBatchCheck)
    
    $yPos += 30
    
    # Protection required warning checkbox
    $protectionCheck = New-Object System.Windows.Forms.CheckBox
    $protectionCheck.Text = Get-LocalizedString -Key "settings.warnings.protectionRequired"
    $protectionCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $protectionCheck.Size = New-Object System.Drawing.Size(520, 20)
    $protectionCheck.Checked = $Config.warnings.showProtectionRequiredWarning
    $warningsTab.Controls.Add($protectionCheck)
    
    $yPos += 30
    
    # No changes warning checkbox
    $noChangesCheck = New-Object System.Windows.Forms.CheckBox
    $noChangesCheck.Text = Get-LocalizedString -Key "settings.warnings.noChanges"
    $noChangesCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $noChangesCheck.Size = New-Object System.Drawing.Size(520, 20)
    $noChangesCheck.Checked = $Config.warnings.showNoChangesWarning
    $warningsTab.Controls.Add($noChangesCheck)
    
    $yPos += 50
    
    # Large batch threshold
    $largeBatchLabel = New-Object System.Windows.Forms.Label
    $largeBatchLabel.Text = Get-LocalizedString -Key "settings.warnings.largeBatchThreshold"
    $largeBatchLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $largeBatchLabel.Size = New-Object System.Drawing.Size(300, 20)
    $warningsTab.Controls.Add($largeBatchLabel)
    
    $largeBatchNumeric = New-Object System.Windows.Forms.NumericUpDown
    $largeBatchNumeric.Location = New-Object System.Drawing.Point(330, $yPos)
    $largeBatchNumeric.Size = New-Object System.Drawing.Size(100, 25)
    $largeBatchNumeric.Minimum = 1
    $largeBatchNumeric.Maximum = 1000
    $largeBatchNumeric.Value = $Config.warnings.largeBatchThreshold
    $warningsTab.Controls.Add($largeBatchNumeric)
    
    $yPos += 35
    
    # Mass downgrade threshold
    $massDowngradeLabel = New-Object System.Windows.Forms.Label
    $massDowngradeLabel.Text = Get-LocalizedString -Key "settings.warnings.massDowngradeThreshold"
    $massDowngradeLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $massDowngradeLabel.Size = New-Object System.Drawing.Size(300, 20)
    $warningsTab.Controls.Add($massDowngradeLabel)
    
    $massDowngradeNumeric = New-Object System.Windows.Forms.NumericUpDown
    $massDowngradeNumeric.Location = New-Object System.Drawing.Point(330, $yPos)
    $massDowngradeNumeric.Size = New-Object System.Drawing.Size(100, 25)
    $massDowngradeNumeric.Minimum = 1
    $massDowngradeNumeric.Maximum = 100
    $massDowngradeNumeric.Value = $Config.warnings.massDowngradeThreshold
    $warningsTab.Controls.Add($massDowngradeNumeric)
    
    # ====================
    # TAB 3: LOGGING
    # ====================
    $loggingTab = New-Object System.Windows.Forms.TabPage
    $loggingTab.Text = Get-LocalizedString -Key "settings.tabs.logging"
    $tabControl.TabPages.Add($loggingTab)
    
    $yPos = 20
    
    # Log retention days
    $retentionLabel = New-Object System.Windows.Forms.Label
    $retentionLabel.Text = Get-LocalizedString -Key "settings.logging.retention"
    $retentionLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $retentionLabel.Size = New-Object System.Drawing.Size(400, 20)
    $loggingTab.Controls.Add($retentionLabel)
    
    $retentionNumeric = New-Object System.Windows.Forms.NumericUpDown
    $retentionNumeric.Location = New-Object System.Drawing.Point(430, $yPos)
    $retentionNumeric.Size = New-Object System.Drawing.Size(100, 25)
    $retentionNumeric.Minimum = 0
    $retentionNumeric.Maximum = 365
    $retentionNumeric.Value = $Config.logging.logRetentionDays
    $loggingTab.Controls.Add($retentionNumeric)
    
    $yPos += 40
    
    # Detailed logging checkbox
    $detailedLoggingCheck = New-Object System.Windows.Forms.CheckBox
    $detailedLoggingCheck.Text = Get-LocalizedString -Key "settings.logging.detailedLogging"
    $detailedLoggingCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $detailedLoggingCheck.Size = New-Object System.Drawing.Size(520, 20)
    $detailedLoggingCheck.Checked = $Config.logging.enableDetailedLogging
    $loggingTab.Controls.Add($detailedLoggingCheck)
    
    $yPos += 40
    
    # Log directory label
    $logDirLabel = New-Object System.Windows.Forms.Label
    $logDirLabel.Text = Get-LocalizedString -Key "settings.logging.customDirectory"
    $logDirLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $logDirLabel.Size = New-Object System.Drawing.Size(520, 20)
    $loggingTab.Controls.Add($logDirLabel)
    
    $yPos += 25
    
    # Log directory textbox
    $logDirText = New-Object System.Windows.Forms.TextBox
    $logDirText.Location = New-Object System.Drawing.Point(20, $yPos)
    $logDirText.Size = New-Object System.Drawing.Size(420, 25)
    $logDirText.Text = $Config.logging.logDirectory
    $loggingTab.Controls.Add($logDirText)
    
    # Browse button for log directory
    $browseLogBtn = New-Object System.Windows.Forms.Button
    $browseLogBtn.Text = '...'
    $browseLogBtn.Location = New-Object System.Drawing.Point(450, $yPos)
    $browseLogBtn.Size = New-Object System.Drawing.Size(40, 25)
    $browseLogBtn.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = Get-LocalizedString -Key "dialogs.selectFolder.description"
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $logDirText.Text = $folderBrowser.SelectedPath
        }
    })
    $loggingTab.Controls.Add($browseLogBtn)
    
    # ====================
    # TAB 4: UI
    # ====================
    $uiTab = New-Object System.Windows.Forms.TabPage
    $uiTab.Text = Get-LocalizedString -Key "settings.tabs.ui"
    $tabControl.TabPages.Add($uiTab)
    
    $yPos = 20
    
    # Language selection
    $languageLabel = New-Object System.Windows.Forms.Label
    $languageLabel.Text = Get-LocalizedString -Key "settings.ui.language"
    $languageLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $languageLabel.Size = New-Object System.Drawing.Size(200, 20)
    $uiTab.Controls.Add($languageLabel)
    
    $languageDropdown = New-Object System.Windows.Forms.ComboBox
    $languageDropdown.Location = New-Object System.Drawing.Point(230, $yPos)
    $languageDropdown.Size = New-Object System.Drawing.Size(290, 25)
    $languageDropdown.DropDownStyle = 'DropDownList'
    $languageDropdown.Items.AddRange(@(
        "Auto-detect (System)",
        "Norwegian (Norsk)",
        "English"
    ))
    
    # Set current selection based on config
    $currentLang = if ($Config.preferences.language) { $Config.preferences.language } else { "auto" }
    switch ($currentLang) {
        "nb-NO" { $languageDropdown.SelectedIndex = 1 }
        "en-US" { $languageDropdown.SelectedIndex = 2 }
        default { $languageDropdown.SelectedIndex = 0 }
    }
    $uiTab.Controls.Add($languageDropdown)
    
    $yPos += 35
    
    # Language note
    $langNoteLabel = New-Object System.Windows.Forms.Label
    $langNoteLabel.Text = Get-LocalizedString -Key "settings.ui.languageNote"
    $langNoteLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $langNoteLabel.Size = New-Object System.Drawing.Size(520, 30)
    $langNoteLabel.ForeColor = [System.Drawing.Color]::Gray
    $langNoteLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $uiTab.Controls.Add($langNoteLabel)
    
    $yPos += 45
    
    # Remember window position checkbox
    $rememberPosCheck = New-Object System.Windows.Forms.CheckBox
    $rememberPosCheck.Text = Get-LocalizedString -Key "settings.ui.rememberPosition"
    $rememberPosCheck.Location = New-Object System.Drawing.Point(20, $yPos)
    $rememberPosCheck.Size = New-Object System.Drawing.Size(520, 20)
    $rememberPosCheck.Checked = $Config.ui.rememberWindowPosition
    $uiTab.Controls.Add($rememberPosCheck)
    
    $yPos += 35
    
    # Info label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = Get-LocalizedString -Key "settings.ui.windowSizeNote"
    $infoLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $infoLabel.Size = New-Object System.Drawing.Size(520, 40)
    $infoLabel.ForeColor = [System.Drawing.Color]::Gray
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $uiTab.Controls.Add($infoLabel)
    
    # ====================
    # BUTTONS
    # ====================
    
    # Reset to defaults button
    $resetBtn = New-Object System.Windows.Forms.Button
    $resetBtn.Text = Get-LocalizedString -Key "buttons.reset"
    $resetBtn.Location = New-Object System.Drawing.Point(20, 560)
    $resetBtn.Size = New-Object System.Drawing.Size(180, 35)
    $resetBtn.Add_Click({
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Dette vil tilbakestille alle innstillinger til standardverdier.`n`nEr du sikker?",
            "Bekreft tilbakestilling",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $resetSuccess = Reset-AppConfig -Confirm $false
            if ($resetSuccess) {
                $settingsForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $settingsForm.Tag = 'reset'
                $settingsForm.Close()
            }
        }
    })
    $settingsForm.Controls.Add($resetBtn)
    
    # Save button
    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = Get-LocalizedString -Key "buttons.save"
    $saveBtn.Location = New-Object System.Drawing.Point(365, 560)
    $saveBtn.Size = New-Object System.Drawing.Size(100, 35)
    $saveBtn.BackColor = [System.Drawing.Color]::LightGreen
    $saveBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $saveBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $settingsForm.Controls.Add($saveBtn)
    $settingsForm.AcceptButton = $saveBtn
    
    # Cancel button
    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = Get-LocalizedString -Key "buttons.cancel"
    $cancelBtn.Location = New-Object System.Drawing.Point(475, 560)
    $cancelBtn.Size = New-Object System.Drawing.Size(100, 35)
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $settingsForm.Controls.Add($cancelBtn)
    $settingsForm.CancelButton = $cancelBtn
    
    # Show dialog
    $result = $settingsForm.ShowDialog()
    
    # Check if reset was triggered
    if ($settingsForm.Tag -eq 'reset') {
        Write-Log -Message "Settings reset to defaults via settings dialog" -Level 'INFO' -Source 'Show-SettingsDialog'
        return $true  # Config was changed (reset)
    }
    
    # If user clicked Save, apply changes
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Update config with UI values
            $Config.preferences.rememberLastLabel = $rememberLabelCheck.Checked
            $Config.preferences.includeSubfoldersDefault = $subfoldersCheck.Checked
            $Config.preferences.defaultFolder = $defaultFolderText.Text.Trim()
            
            # Update language preference
            switch ($languageDropdown.SelectedIndex) {
                0 { $Config.preferences.language = "auto" }
                1 { $Config.preferences.language = "nb-NO" }
                2 { $Config.preferences.language = "en-US" }
            }
            
            $Config.warnings.showPreApplySummary = $showSummaryCheck.Checked
            $Config.warnings.showMassDowngradeWarning = $massDowngradeCheck.Checked
            $Config.warnings.showLargeBatchWarning = $largeBatchCheck.Checked
            $Config.warnings.showProtectionRequiredWarning = $protectionCheck.Checked
            $Config.warnings.showNoChangesWarning = $noChangesCheck.Checked
            $Config.warnings.largeBatchThreshold = [int]$largeBatchNumeric.Value
            $Config.warnings.massDowngradeThreshold = [int]$massDowngradeNumeric.Value
            
            $Config.logging.logRetentionDays = [int]$retentionNumeric.Value
            $Config.logging.enableDetailedLogging = $detailedLoggingCheck.Checked
            $Config.logging.logDirectory = $logDirText.Text.Trim()
            
            $Config.ui.rememberWindowPosition = $rememberPosCheck.Checked
            
            # Validate updated config
            $validation = Validate-AppConfig -Config $Config
            
            if (-not $validation.IsValid) {
                $errorMsg = "Ugyldige innstillinger:`n`n" + ($validation.Errors -join "`n")
                [System.Windows.Forms.MessageBox]::Show(
                    $errorMsg,
                    "Valideringsfeil",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return $false
            }
            
            # Save to file
            $saveResult = Save-AppConfig -Config $Config
            
            if ($saveResult) {
                Write-Log -Message "Settings updated successfully" -Level 'INFO' -Source 'Show-SettingsDialog'
                [System.Windows.Forms.MessageBox]::Show(
                    "Innstillinger lagret.`n`nNoen endringer trer i kraft neste gang applikasjonen startes.",
                    "Lagret",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                return $true
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Kunne ikke lagre innstillinger. Se loggfil for detaljer.",
                    "Lagringsfeil",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return $false
            }
        }
        catch {
            Write-Log -Message "Failed to save settings from dialog" -Level 'ERROR' -Source 'Show-SettingsDialog' -Exception $_
            [System.Windows.Forms.MessageBox]::Show(
                "Feil under lagring: $($_.Exception.Message)",
                "Feil",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
    }
    
    return $false  # No changes made
}

function Update-FileCount {
    if ($selectedFiles.Count -eq 0) {
        $fileCountLabel.Text = Get-LocalizedString -Key "labels.fileCount_none"
        $fileCountLabel.ForeColor = [System.Drawing.Color]::Gray
    } elseif ($selectedFiles.Count -eq 1) {
        $fileCountLabel.Text = Get-LocalizedString -Key "labels.fileCount_single"
        $fileCountLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $fileCountLabel.Text = Get-LocalizedString -Key "labels.fileCount_multiple" -Parameters @($selectedFiles.Count)
        $fileCountLabel.ForeColor = [System.Drawing.Color]::Green
    }
}

function Remove-FileFromSelection {
    <#
    .SYNOPSIS
        Removes a single file from the selected files list
    .DESCRIPTION
        Removes file from selection array, label cache, updates UI display and layout
    .PARAMETER FilePath
        Full path of the file to remove
    .PARAMETER Index
        Optional: Index in the selectedFiles array (for performance)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [int]$Index = -1
    )
    
    try {
        # Validate file exists in selection
        if ($Index -eq -1) {
            $Index = [Array]::IndexOf($script:selectedFiles, $FilePath)
        }
        
        if ($Index -eq -1) {
            Write-Log -Message "File not found in selection: $FilePath" -Level 'WARNING' -Source 'Remove-FileFromSelection'
            return $false
        }
        
        # Remove from selectedFiles array
        # PowerShell arrays are immutable, so we create a new array without the item
        $newArray = @()
        for ($i = 0; $i -lt $script:selectedFiles.Count; $i++) {
            if ($i -ne $Index) {
                $newArray += $script:selectedFiles[$i]
            }
        }
        $script:selectedFiles = $newArray
        
        # Remove from label cache
        if ($script:fileLabelCache.ContainsKey($FilePath)) {
            $script:fileLabelCache.Remove($FilePath)
        }
        
        # Log removal
        Write-Log -Message "File removed from selection: $FilePath" -Level 'INFO' -Source 'Remove-FileFromSelection' -Context @{
            RemainingFiles = $script:selectedFiles.Count
        }
        
        # Update UI
        Update-FileListDisplay
        Update-FileCount
        Adjust-UILayout
        
        return $true
        
    } catch {
        Write-Log -Message "Failed to remove file from selection" -Level 'ERROR' -Source 'Remove-FileFromSelection' -Exception $_ -Context @{
            FilePath = $FilePath
        }
        return $false
    }
}

function Get-FileLabelDisplayName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    # Check cache first
    if ($script:fileLabelCache.ContainsKey($FilePath)) {
        $cached = $script:fileLabelCache[$FilePath]
        # Return display name (could be string or hashtable)
        if ($cached -is [hashtable]) {
            return $cached.DisplayName
        }
        return $cached
    }
    
    # Retrieve label status
    try {
        $labelStatus = Get-AIPFileStatus -Path $FilePath -ErrorAction SilentlyContinue
        
        if ($labelStatus -and $labelStatus.MainLabelId) {
            # Find matching label from config
            $labelObj = $script:labels | Where-Object { $_.Id -eq $labelStatus.MainLabelId }
            if ($labelObj) {
                $displayText = $labelObj.DisplayName
                # Cache full label info (not just display name)
                $script:fileLabelCache[$FilePath] = @{
                    DisplayName = $displayText
                    LabelId = $labelStatus.MainLabelId
                    Rank = if($labelObj.Rank) { $labelObj.Rank } else { 0 }
                }
                return $displayText
            } else {
                # Label ID exists but not found in configuration
                # This can happen with encrypted/protected labels not in labels_config.json
                Write-Log -Message "Label ID not found in configuration" -Level 'WARNING' -Source 'Get-FileLabelDisplayName' -Context @{ 
                    FilePath = $FilePath
                    MainLabelId = $labelStatus.MainLabelId
                    IsProtected = $labelStatus.IsProtected
                }
                
                $script:fileLabelCache[$FilePath] = @{
                    DisplayName = "Ukjent etikett (beskyttet)"
                    LabelId = $labelStatus.MainLabelId
                    Rank = -1
                    IsProtected = $labelStatus.IsProtected
                }
                return "Ukjent etikett (beskyttet)"
            }
        }
        
        # No label found
        $script:fileLabelCache[$FilePath] = @{
            DisplayName = "Ingen etikett"
            LabelId = $null
            Rank = -1
        }
        return "Ingen etikett"
        
    } catch {
        # Enhanced error logging with full exception details
        Write-Log -Message "Label retrieval failed" -Level 'ERROR' -Source 'Get-FileLabelDisplayName' -Context @{ 
            FilePath = $FilePath
            ErrorType = $_.Exception.GetType().Name
        } -Exception $_
        
        $script:fileLabelCache[$FilePath] = @{
            DisplayName = "Feil ved henting"
            LabelId = $null
            Rank = -1
            Error = $_.Exception.Message
        }
        return "Feil ved henting"
    }
}

function Update-FileListDisplay {
    <#
    .SYNOPSIS
        Refreshes file list box with current labels
    .DESCRIPTION
        Uses async label retrieval for large file lists (>50 files) to prevent UI freeze
    #>
    param(
        [bool]$ForceAsync = $false
    )
    
    # Refresh the file list box with current labels
    if ($script:selectedFiles.Count -eq 0) {
        return
    }
    
    $fileCount = $script:selectedFiles.Count
    
    # === ASYNC vs SYNC DECISION ===
    # ALWAYS use async if runspace pool available (prevents crashes from locked files)
    # - Async: Parallel processing, handles errors gracefully, non-blocking
    # - Sync: ONLY as fallback if runspace pool not available
    # - Performance: Async is 4x faster for large sets, safe for all sizes
    $useAsyncRetrieval = ($script:runspacePool -ne $null -or $ForceAsync)
    
    if ($useAsyncRetrieval) {
        # ASYNC MODE (prevents crashes and improves performance)
        Write-Log -Message "Using async label retrieval" -Level 'INFO' -Source 'Update-FileListDisplay' -Context @{ FileCount = $fileCount }
        
        # Clear listbox and show placeholder
        $fileListBox.Items.Clear()
        foreach ($file in $script:selectedFiles) {
            $fileName = [System.IO.Path]::GetFileName($file)
            $fileListBox.Items.Add("$fileName [...]")
        }
        
        # Reset progress
        $script:asyncSharedProgress.Processed = 0
        $script:asyncSharedProgress.Total = $fileCount
        
        # Show progress UI
        $progressBar.Style = 'Continuous'
        $progressBar.Value = 0
        $statusLabel.Text = "Henter etikettinformasjon asynkront..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        # Clear async cache before starting to avoid stale data
        $script:asyncSharedCache.Clear()
        
        try {
            # Start async label retrieval
            $labelJobs = Start-AsyncLabelRetrieval -FilePaths $script:selectedFiles `
                                                    -RunspacePool $script:runspacePool `
                                                    -SharedCache $script:asyncSharedCache `
                                                    -SharedProgress $script:asyncSharedProgress `
                                                    -AvailableLabels $script:labels
            
            # Wait for completion with UI updates
            try {
                $labelResults = Wait-AsyncJobsWithUI -Jobs $labelJobs `
                                                      -SharedProgress $script:asyncSharedProgress `
                                                      -ProgressBar $progressBar `
                                                      -StatusLabel $statusLabel `
                                                      -Form $form `
                                                      -OperationType "Henter etiketter"
                
                Write-Log -Message "Async label retrieval completed" -Level 'INFO' -Source 'Update-FileListDisplay' -Context @{ ResultsReceived = $labelResults.Count }
                
                # CRITICAL: Merge asyncSharedCache into main fileLabelCache
                # Async operations populate asyncSharedCache, but display reads from fileLabelCache
                $lockTaken = $false
                try {
                    [System.Threading.Monitor]::Enter($script:asyncSharedCache.SyncRoot, [ref]$lockTaken)
                    
                    foreach ($key in $script:asyncSharedCache.Keys) {
                        $script:fileLabelCache[$key] = $script:asyncSharedCache[$key]
                    }
                    
                    Write-Log -Message "Merged async cache to main cache" -Level 'INFO' -Source 'Update-FileListDisplay' -Context @{ MergedCount = $script:asyncSharedCache.Count }
                } finally {
                    if ($lockTaken) {
                        [System.Threading.Monitor]::Exit($script:asyncSharedCache.SyncRoot)
                    }
                }
                
            } catch {
                Write-Log -Message "Wait-AsyncJobsWithUI failed" -Level 'ERROR' -Source 'Update-FileListDisplay' -Context @{ FileCount = $fileCount } -Exception $_
                throw
            }
            
            # Now update display with retrieved labels
            try {
                $statusLabel.Text = "Oppdaterer visning..."
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                $form.Refresh()
                
                # Small delay to let UI settle
                Start-Sleep -Milliseconds 50
                [System.Windows.Forms.Application]::DoEvents()
                
                # Clear listbox with BeginUpdate/EndUpdate for performance
                $fileListBox.BeginUpdate()
                $fileListBox.Items.Clear()
                $itemsAdded = 0
                
                foreach ($file in $script:selectedFiles) {
                    try {
                        $fileName = [System.IO.Path]::GetFileName($file)
                        
                        # Get from cache (should be populated now)
                        if ($script:fileLabelCache.ContainsKey($file)) {
                            $cached = $script:fileLabelCache[$file]
                            $labelName = if ($cached -is [hashtable]) { $cached.DisplayName } else { $cached }
                        } else {
                            $labelName = "Ukjent"
                            Write-Log -Message "File not in cache after async retrieval" -Level 'WARNING' -Source 'Update-FileListDisplay' -Context @{ FilePath = $file }
                        }
                        
                        $displayText = "$fileName [$labelName]"
                        [void]$fileListBox.Items.Add($displayText)
                        $itemsAdded++
                        
                    } catch {
                        Write-Log -Message "Failed to add file to display" -Level 'ERROR' -Source 'Update-FileListDisplay' -Context @{ FilePath = $file } -Exception $_
                        # Try to add with error indicator
                        try {
                            $fileName = [System.IO.Path]::GetFileName($file)
                            [void]$fileListBox.Items.Add("$fileName [FEIL]")
                        } catch {
                            Write-Log -Message "Failed to add file with error indicator" -Level 'ERROR' -Source 'Update-FileListDisplay' -Exception $_
                        }
                    }
                }
                
                Write-Log -Message "Display updated with files" -Level 'INFO' -Source 'Update-FileListDisplay' -Context @{ ItemsAdded = $itemsAdded; TotalFiles = $script:selectedFiles.Count }
                
                # End batch update
                $fileListBox.EndUpdate()
                
                $statusLabel.Text = "Etiketter hentet ($itemsAdded filer)."
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                $progressBar.Value = 0
                
            } catch {
                Write-Log -Message "Display update failed" -Level 'ERROR' -Source 'Update-FileListDisplay' -Context @{ FileCount = $itemsAdded } -Exception $_
                
                # Ensure EndUpdate is called even on error
                try {
                    $fileListBox.EndUpdate()
                } catch {
                    Write-Log -Message "EndUpdate failed" -Level 'WARNING' -Source 'Update-FileListDisplay' -Exception $_
                }
                
                $errorInfo = Get-FriendlyErrorMessage -Exception $_
                Show-ErrorDialog -ErrorInfo $errorInfo -Title "Visningsfeil"
                
                throw
            }
            
        } catch {
            Write-Log -Message "Async label retrieval failed" -Level 'ERROR' -Source 'Update-FileListDisplay' -Exception $_
            
            # Fall back to sync mode
            $statusLabel.Text = "Async feilet, prøver synkron modus..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            $progressBar.Value = 0
            $form.Refresh()
            
            $useAsyncRetrieval = $false
        }
    }
    
    if (-not $useAsyncRetrieval) {
        # SYNC MODE for small file lists or fallback
        Write-Log -Message "Using sync label retrieval" -Level 'INFO' -Source 'Update-FileListDisplay' -Context @{ FileCount = $fileCount }
        
        $fileListBox.Items.Clear()
        foreach ($file in $script:selectedFiles) {
            $fileName = [System.IO.Path]::GetFileName($file)
            $labelName = Get-FileLabelDisplayName -FilePath $file
            $displayText = "$fileName [$labelName]"
            $fileListBox.Items.Add($displayText)
        }
    }
    
    # Adjust UI layout based on file count
    Adjust-UILayout
}

function Adjust-UILayout {
    # Calculate optimal listbox height based on file count
    $fileCount = $script:selectedFiles.Count
    
    # Calculate height: min 3 rows, max 10 rows, default shows all files if <= 10
    $itemHeight = 16  # Height per item in pixels (accounting for font and spacing)
    $minRows = 3
    $maxRows = 10
    
    if ($fileCount -eq 0) {
        $rows = $minRows
    } elseif ($fileCount -le $maxRows) {
        $rows = [Math]::Max($fileCount, $minRows)
    } else {
        $rows = $maxRows
    }
    
    $newListBoxHeight = $rows * $itemHeight + 8  # +8 for border/padding
    $oldListBoxHeight = $fileListBox.Height
    $heightDelta = $newListBoxHeight - $oldListBoxHeight
    
    # Update listbox size
    $fileListBox.Height = $newListBoxHeight
    
    # Update file count label position
    $fileCountLabel.Top = 20 + $newListBoxHeight + 10
    
    # Update file group box height
    # Minimum height must be 170 to accommodate all buttons (Browse, Folder, Checkbox, Clear)
    $calculatedHeight = $fileCountLabel.Top + $fileCountLabel.Height + 10
    $newFileGroupHeight = [Math]::Max($calculatedHeight, 170)
    $oldFileGroupHeight = $fileGroupBox.Height
    $fileGroupBox.Height = $newFileGroupHeight
    
    $groupHeightDelta = $newFileGroupHeight - $oldFileGroupHeight
    
    # Adjust dependent controls (shift down by height delta)
    $labelGroupBox.Top = 10 + $fileGroupBox.Height + 10
    $progressGroupBox.Top = $labelGroupBox.Top + $labelGroupBox.Height + 10
    
    # Position buttons with better spacing and centering
    $buttonY = $progressGroupBox.Top + $progressGroupBox.Height + 20  # Increased spacing from 15 to 20
    
    # Center the three buttons horizontally
    # Total button width: 200 + 100 + 100 = 400px
    # Total spacing: 10px between each = 20px
    # Total needed: 420px
    # Form content width: 720px (740 - 20 for borders)
    # Left margin for centering: (720 - 420) / 2 = 150px
    $centerStartX = 150
    
    $applyBtn.Location = New-Object System.Drawing.Point($centerStartX, $buttonY)
    $viewLogBtn.Location = New-Object System.Drawing.Point(($centerStartX + 200 + 10), $buttonY)
    $settingsBtn.Location = New-Object System.Drawing.Point(($centerStartX + 200 + 10 + 100 + 10), $buttonY)
    
    # Adjust footer label position with more spacing
    $footerLabel.Top = $applyBtn.Top + $applyBtn.Height + 20  # Increased spacing from 10 to 20
    
    # Adjust form height with extra bottom padding
    $newFormHeight = $footerLabel.Top + $footerLabel.Height + 15  # Increased from 10 to 15
    $form.Height = $newFormHeight
    
    $form.Refresh()
}

# ========================================
# LOAD APPLICATION CONFIGURATION
# ========================================
$appConfig = Load-AppConfig

# Override log directory if configured
if ($appConfig.logging.logDirectory -and $appConfig.logging.logDirectory -ne "") {
    $customLogDir = $appConfig.logging.logDirectory
    if (Test-Path $customLogDir -PathType Container) {
        $logDirectory = $customLogDir
        $logFilePath = Join-Path $logDirectory "FileLabeler_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        Write-Log -Message "Using custom log directory" -Level 'INFO' -Source 'Startup' -Context @{ LogDirectory = $logDirectory }
    }
    else {
        Write-Log -Message "Configured log directory not found, using default" -Level 'WARNING' -Source 'Startup' -Context @{ ConfiguredPath = $customLogDir; DefaultPath = $logDirectory }
    }
}

# Cleanup old logs based on retention setting
Cleanup-OldLogs -RetentionDays $appConfig.logging.logRetentionDays -LogDir $logDirectory

# ========================================
# INITIALIZE LANGUAGE RESOURCES
# ========================================
Write-Log -Message "Initializing language resources..." -Level 'INFO' -Source 'Startup'
$languageInitResult = Initialize-LanguageResources -Config $appConfig

if (-not $languageInitResult) {
    Write-Log -Message "Failed to initialize language resources, using hardcoded Norwegian strings" -Level 'WARNING' -Source 'Startup'
    # Application will continue but without localization support
}
else {
    Write-Log -Message "Language resources initialized: $($script:CurrentLanguage)" -Level 'INFO' -Source 'Startup'
}

Write-Log -Message "========================================" -Level 'INFO'
Write-Log -Message "FileLabeler started - v1.1" -Level 'INFO' -Source 'Startup'
Write-Log -Message "Config version: $($appConfig.version)" -Level 'INFO' -Source 'Startup'
Write-Log -Message "Log directory: $logDirectory" -Level 'INFO' -Source 'Startup'
Write-Log -Message "Async operations: $(if($runspacePool) { 'Enabled' } else { 'Disabled' })" -Level 'INFO' -Source 'Startup'
Write-Log -Message "========================================" -Level 'INFO'

# ========================================
# BUILD FORM
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = Get-LocalizedString -Key "form.title"
# Ensure minimum height for proper display
$formHeight = [Math]::Max($appConfig.ui.windowHeight, 600)
$form.Size = New-Object System.Drawing.Size($appConfig.ui.windowWidth, $formHeight)
$form.StartPosition = 'CenterScreen'

# Apply saved window position if configured
if ($appConfig.ui.rememberWindowPosition -and $appConfig.ui.windowPositionX -and $appConfig.ui.windowPositionY) {
    # Validate position is within screen bounds
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $validPosition = $false
    
    foreach ($screen in $screens) {
        if ($appConfig.ui.windowPositionX -ge $screen.WorkingArea.Left -and
            $appConfig.ui.windowPositionX -lt $screen.WorkingArea.Right -and
            $appConfig.ui.windowPositionY -ge $screen.WorkingArea.Top -and
            $appConfig.ui.windowPositionY -lt $screen.WorkingArea.Bottom) {
            $validPosition = $true
            break
        }
    }
    
    if ($validPosition) {
        $form.StartPosition = 'Manual'
        $form.Location = New-Object System.Drawing.Point($appConfig.ui.windowPositionX, $appConfig.ui.windowPositionY)
    }
}

$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AllowDrop = $true

# ========================================
# INITIALIZE RUNSPACE POOL FOR ASYNC OPERATIONS
# ========================================
Write-Log -Message "Initializing runspace pool for async operations" -Level 'INFO' -Source 'Startup'
$runspacePool = $null
$asyncJobs = @()

try {
    $runspacePool = New-FileLabelerRunspacePool -MinRunspaces 1 -MaxRunspaces 4
    Write-Log -Message "Runspace pool created successfully" -Level 'INFO' -Source 'AsyncInitialization'
} catch {
    Write-Log -Message "Failed to create runspace pool - async operations will be disabled" -Level 'WARNING' -Source 'AsyncInitialization' -Exception $_
}

# Shared data structures for async operations (synchronized)
$script:asyncSharedCache = [Hashtable]::Synchronized(@{})
$script:asyncSharedCache.SyncRoot = New-Object System.Object

$script:asyncSharedProgress = @{
    Processed = 0
    Total = 0
}

# ========================================
# FILE SELECTION SECTION
# ========================================
$fileGroupBox = New-Object System.Windows.Forms.GroupBox
$fileGroupBox.Text = Get-LocalizedString -Key "groupBoxes.fileSelection"
$fileGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$fileGroupBox.Size = New-Object System.Drawing.Size(700, 170)  # Increased height from 150 to 170 to fit all buttons
$form.Controls.Add($fileGroupBox)

# File list box
$fileListBox = New-Object System.Windows.Forms.ListBox
$fileListBox.Location = New-Object System.Drawing.Point(10, 20)
$fileListBox.Size = New-Object System.Drawing.Size(580, 60)
$fileListBox.HorizontalScrollbar = $true
$fileListBox.AllowDrop = $true
$fileGroupBox.Controls.Add($fileListBox)

# Context menu for file removal
$fileContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$removeMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$removeMenuItem.Text = Get-LocalizedString -Key "buttons.removeFile"
$removeMenuItem.Add_Click({
    # Get selected index
    $selectedIndex = $fileListBox.SelectedIndex
    
    # Validate selection
    if ($selectedIndex -eq -1) {
        Write-Log -Message "No file selected for removal" -Level 'WARNING' -Source 'ContextMenu'
        return
    }
    
    # Get file path from selection
    if ($selectedIndex -ge 0 -and $selectedIndex -lt $script:selectedFiles.Count) {
        $filePath = $script:selectedFiles[$selectedIndex]
        
        # Remove file from selection
        $result = Remove-FileFromSelection -FilePath $filePath -Index $selectedIndex
        
        if ($result) {
            Write-Log -Message "File successfully removed via context menu: $filePath" -Level 'INFO' -Source 'ContextMenu'
        }
    } else {
        Write-Log -Message "Invalid index for file removal: $selectedIndex" -Level 'ERROR' -Source 'ContextMenu'
    }
})
$fileContextMenu.Items.Add($removeMenuItem)
$fileListBox.ContextMenuStrip = $fileContextMenu

# Delete key support for quick file removal
$fileListBox.Add_KeyDown({
    param($sender, $e)
    
    # Check for Delete key
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
        # Get selected index
        $selectedIndex = $fileListBox.SelectedIndex
        
        # Validate selection
        if ($selectedIndex -eq -1) {
            Write-Log -Message "No file selected for removal (Delete key)" -Level 'WARNING' -Source 'DeleteKey'
            return
        }
        
        # Get file path from selection
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $script:selectedFiles.Count) {
            $filePath = $script:selectedFiles[$selectedIndex]
            
            # Remove file from selection
            $result = Remove-FileFromSelection -FilePath $filePath -Index $selectedIndex
            
            if ($result) {
                Write-Log -Message "File successfully removed via Delete key: $filePath" -Level 'INFO' -Source 'DeleteKey'
            }
            
            # Mark event as handled
            $e.Handled = $true
        } else {
            Write-Log -Message "Invalid index for file removal (Delete key): $selectedIndex" -Level 'ERROR' -Source 'DeleteKey'
        }
    }
})

# Browse button
$browseBtn = New-Object System.Windows.Forms.Button
$browseBtn.Text = Get-LocalizedString -Key "buttons.browse"
$browseBtn.Location = New-Object System.Drawing.Point(600, 20)
$browseBtn.Size = New-Object System.Drawing.Size(90, 28)
$browseBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Multiselect = $true
    $dialog.Title = 'Velg filer å merke'
    $dialog.Filter = 'Office & PDF-filer|*.docx;*.xlsx;*.pptx;*.doc;*.xls;*.ppt;*.pdf|Word-dokumenter|*.docx;*.doc|Excel-arbeidsbøker|*.xlsx;*.xls|PowerPoint-presentasjoner|*.pptx;*.ppt|PDF-filer|*.pdf|Alle filer|*.*'
    $dialog.FilterIndex = 1
    
    if ($dialog.ShowDialog() -eq 'OK') {
        # Merge with existing selection (don't replace)
        $newFiles = $dialog.FileNames
        $mergeResult = Merge-FileSelection -NewFiles $newFiles
        
        if ($mergeResult.NewCount -gt 0) {
            $script:selectedFiles = $mergeResult.MergedFiles
            
            # Show loading message
            $statusLabel.Text = Get-LocalizedString -Key "status.retrievingLabels"
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            # Update display with labels
            Update-FileListDisplay
            
            Update-FileCount
            $statusLabel.Text = Get-LocalizedString -Key "status.filesLoaded" -Parameters @($mergeResult.TotalCount, $mergeResult.NewCount)
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
        } else {
            $statusLabel.Text = Get-LocalizedString -Key "status.noNewFiles"
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        }
    }
})
$fileGroupBox.Controls.Add($browseBtn)

# Include subfolders checkbox (MUST be defined before folder button that references it)
$includeSubfoldersCheckbox = New-Object System.Windows.Forms.CheckBox
$includeSubfoldersCheckbox.Text = Get-LocalizedString -Key "labels.includeSubfolders"
$includeSubfoldersCheckbox.Location = New-Object System.Drawing.Point(600, 86)
$includeSubfoldersCheckbox.Size = New-Object System.Drawing.Size(90, 35)  # Increased height to allow text wrap
$includeSubfoldersCheckbox.Checked = $appConfig.preferences.includeSubfoldersDefault
$includeSubfoldersCheckbox.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$fileGroupBox.Controls.Add($includeSubfoldersCheckbox)

# Folder button
$folderBtn = New-Object System.Windows.Forms.Button
$folderBtn.Text = Get-LocalizedString -Key "buttons.folder"
$folderBtn.Location = New-Object System.Drawing.Point(600, 53)
$folderBtn.Size = New-Object System.Drawing.Size(90, 28)
$folderBtn.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = 'Velg mappe å skanne for filer'
    $folderDialog.ShowNewFolderButton = $false
    
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $selectedFolder = $folderDialog.SelectedPath
        
        # Check if runspace pool is available
        if (-not $script:runspacePool) {
            # Fallback to synchronous mode
            Write-Log -Message "Running folder scan in synchronous mode (runspace pool not available)" -Level 'INFO' -Source 'FolderButton'
            
            # Show loading message
            $statusLabel.Text = Get-LocalizedString -Key "status.scanningFolder"
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            # Scan folder using centralized function
            try {
                $foundFiles = Get-SupportedFilesFromFolder -FolderPath $selectedFolder -Recursive $includeSubfoldersCheckbox.Checked
                
                if ($foundFiles.Count -gt 0) {
                    # Merge with existing selection using centralized function
                    $mergeResult = Merge-FileSelection -NewFiles $foundFiles
                    
                    if ($mergeResult.NewCount -gt 0) {
                        $script:selectedFiles = $mergeResult.MergedFiles
                        
                        # Update display
                        $statusLabel.Text = Get-LocalizedString -Key "status.retrievingLabels"
                        $form.Refresh()
                        Update-FileListDisplay
                        
                        Update-FileCount
                        $statusLabel.Text = Get-LocalizedString -Key "status.folderScanned" -Parameters @($mergeResult.TotalCount, $mergeResult.NewCount)
                        $statusLabel.ForeColor = [System.Drawing.Color]::Green
                        Write-Log -Message "Folder scanned successfully" -Level 'INFO' -Source 'FolderScan' -Context @{ FolderPath = $selectedFolder; FilesFound = $mergeResult.TotalCount; NewFiles = $mergeResult.NewCount }
                    } else {
                        $statusLabel.Text = Get-LocalizedString -Key "status.noNewFilesInFolder"
                        $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                    }
                } else {
                    $statusLabel.Text = Get-LocalizedString -Key "status.noSupportedFiles"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                    [System.Windows.Forms.MessageBox]::Show(
                        "Ingen støttede Office- eller PDF-filer ble funnet i valgt mappe.`n`nStøttede filtyper: .docx, .xlsx, .pptx, .doc, .xls, .ppt, .pdf",
                        "Ingen filer funnet",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                }
            }
            catch {
                $statusLabel.Text = Get-LocalizedString -Key "errors.couldNotScanFolder"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                
                Write-Log -Message "Folder scanning failed" -Level 'ERROR' -Source 'FolderScan' -Context @{ FolderPath = $selectedFolder; IsRecursive = $includeSubfoldersCheckbox.Checked } -Exception $_
                
                $errorInfo = Get-FriendlyErrorMessage -Exception $_
                Show-ErrorDialog -ErrorInfo $errorInfo -Title "Mappeskanning feilet"
            }
            
            return
        }
        
        # ASYNC MODE - Use runspaces
        Write-Log -Message "Starting async folder scan" -Level 'INFO' -Source 'FolderButton' -Context @{ FolderPath = $selectedFolder }
        
        # Disable button during scan
        $folderBtn.Enabled = $false
        $browseBtn.Enabled = $false
        
        # Show loading message
        $statusLabel.Text = Get-LocalizedString -Key "status.scanningFolderAsync"
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $progressBar.Style = 'Marquee'  # Indeterminate progress
        $form.Refresh()
        
        # Prepare shared data
        $scanData = [Hashtable]::Synchronized(@{})
        $scanData.SyncRoot = New-Object System.Object
        $scanData.ScannedFiles = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        $scanData.ScanComplete = $false
        $scanData.ScanSuccess = $false
        $scanData.ScanError = $null
        
        try {
            # Start async folder scan
            $scanJob = Start-AsyncFolderScan -FolderPath $selectedFolder `
                                             -RunspacePool $script:runspacePool `
                                             -SharedData $scanData `
                                             -Recursive $includeSubfoldersCheckbox.Checked
            
            # Monitor scan progress
            $monitorTimer = New-Object System.Windows.Forms.Timer
            $monitorTimer.Interval = 200  # Check every 200ms
            
            # Store references in timer's Tag for access in callback
            $monitorTimer.Tag = @{
                ScanData = $scanData
                ScanJob = $scanJob
                StartTime = Get-Date
                TimeoutSeconds = 300  # 5 minute timeout
            }
            
            $monitorTimer.Add_Tick({
                param($sender, $e)
                
                try {
                    $timer = $sender
                    $data = $timer.Tag.ScanData
                    $startTime = $timer.Tag.StartTime
                    $timeout = $timer.Tag.TimeoutSeconds
                    
                    # Check for timeout
                $elapsed = ((Get-Date) - $startTime).TotalSeconds
                if ($elapsed -gt $timeout) {
                    $timer.Stop()
                    $timer.Dispose()
                    
                    # Reset UI
                    $progressBar.Style = 'Continuous'
                    $progressBar.Value = 0
                    $folderBtn.Enabled = $true
                    $browseBtn.Enabled = $true
                    
                    $statusLabel.Text = Get-LocalizedString -Key "status.timeout" -Parameters @([math]::Round($elapsed))
                    $statusLabel.ForeColor = [System.Drawing.Color]::Red
                    Write-Log -Message "Folder scan timeout" -Level 'ERROR' -Source 'FolderButton' -Context @{ ElapsedSeconds = $elapsed }
                    
                    # Cleanup job
                    try {
                        $job = $timer.Tag.ScanJob
                        $job.PowerShell.Stop()
                        $job.PowerShell.Dispose()
                    } catch {
                        Write-Log -Message "Error stopping timed-out scan job" -Level 'WARNING' -Source 'FolderButton' -Exception $_
                    }
                    
                    return
                }
                
                if ($data.ScanComplete) {
                    $timer.Stop()
                    $timer.Dispose()
                    
                    # Reset progress bar
                    $progressBar.Style = 'Continuous'
                    $progressBar.Value = 0
                    
                    # Re-enable buttons
                    $folderBtn.Enabled = $true
                    $browseBtn.Enabled = $true
                    
                    if ($data.ScanSuccess) {
                        # Process results
                        $foundFiles = $data.ScannedFiles
                        $fileCount = $foundFiles.Count
                        
                        Write-Log -Message "Async folder scan complete" -Level 'INFO' -Source 'FolderButton' -Context @{ FilesFound = $fileCount }
                        
                        if ($fileCount -gt 0) {
                            # Merge with existing selection using centralized function
                            $mergeResult = Merge-FileSelection -NewFiles $foundFiles
                            
                            if ($mergeResult.NewCount -gt 0) {
                                $script:selectedFiles = $mergeResult.MergedFiles
                                
                                # Update display (always call to trigger UI refresh)
                                try {
                                    $statusLabel.Text = Get-LocalizedString -Key "status.updatingDisplay"
                                    $form.Refresh()
                                    Update-FileListDisplay
                                    
                                    Update-FileCount
                                    $statusLabel.Text = Get-LocalizedString -Key "status.folderScanned" -Parameters @($fileCount, $mergeResult.NewCount)
                                    $statusLabel.ForeColor = [System.Drawing.Color]::Green
                                } catch {
                                    Write-Log -Message "Error updating display after folder scan" -Level 'ERROR' -Source 'FolderButton' -Exception $_
                                    
                                    # Fallback: just update count
                                    Update-FileCount
                                    $statusLabel.Text = Get-LocalizedString -Key "status.folderScanned" -Parameters @($fileCount, $mergeResult.NewCount)
                                    $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                                }
                                
                                Write-Log -Message "Async folder scan completed" -Level 'INFO' -Source 'FolderScan' -Context @{ FolderPath = $selectedFolder; FilesFound = $fileCount; NewFiles = $mergeResult.NewCount }
                                
                                <# OLD CODE - Removed to prevent nested async operations which caused crashes
                                The label retrieval now happens automatically in Update-FileListDisplay #>
                            } else {
                                $statusLabel.Text = Get-LocalizedString -Key "status.noNewFilesInFolder"
                                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                            }
                        } else {
                            $statusLabel.Text = Get-LocalizedString -Key "status.noSupportedFiles"
                            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                            [System.Windows.Forms.MessageBox]::Show(
                                "Ingen støttede Office- eller PDF-filer ble funnet i valgt mappe.`n`nStøttede filtyper: .docx, .xlsx, .pptx, .doc, .xls, .ppt, .pdf",
                                "Ingen filer funnet",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Information
                            )
                        }
                    } else {
                        # Scan failed
                        $statusLabel.Text = Get-LocalizedString -Key "errors.couldNotScanFolder"
                        $statusLabel.ForeColor = [System.Drawing.Color]::Red
                        Write-Log -Message "Async folder scan failed" -Level 'ERROR' -Source 'FolderButton' -Context @{ FolderPath = $selectedFolder; ErrorMessage = $data.ScanError }
                        [System.Windows.Forms.MessageBox]::Show(
                            "Kunne ikke skanne mappen:`n`n$($data.ScanError)",
                            "Feil",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Error
                        )
                    }
                    
                    # Cleanup
                    $job = $timer.Tag.ScanJob
                    try {
                        $job.PowerShell.EndInvoke($job.Handle) | Out-Null
                        $job.PowerShell.Dispose()
                    } catch {
                        Write-Log -Message "Error cleaning up scan job" -Level 'WARNING' -Source 'FolderButton' -Exception $_
                    }
                }
                
                # Allow UI to process events
                [System.Windows.Forms.Application]::DoEvents()
                    
                } catch {
                    # Error in timer callback - stop timer and recover
                    Write-Log -Message "Error in folder scan timer callback" -Level 'ERROR' -Source 'FolderButton' -Exception $_
                    
                    try {
                        $timer.Stop()
                        $timer.Dispose()
                    } catch { }
                    
            # Reset UI
            $progressBar.Style = 'Continuous'
            $progressBar.Value = 0
            $folderBtn.Enabled = $true
            $browseBtn.Enabled = $true
            
            $statusLabel.Text = Get-LocalizedString -Key "errors.errorDuringFolderScan"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        "En feil oppstod under mappeskanning:`n`n$($_.Exception.Message)",
                        "Feil",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
            })
            
            $monitorTimer.Start()
            
        } catch {
            # Error starting async operation
            $progressBar.Style = 'Continuous'
            $progressBar.Value = 0
            $folderBtn.Enabled = $true
            $browseBtn.Enabled = $true
            
            $statusLabel.Text = Get-LocalizedString -Key "errors.couldNotStartAsyncScan"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            Write-Log -Message "Could not start async folder scan" -Level 'ERROR' -Source 'FolderButton' -Exception $_
            [System.Windows.Forms.MessageBox]::Show(
                "Kunne ikke starte mappeskanning:`n`n$($_.Exception.Message)",
                "Feil",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})
$fileGroupBox.Controls.Add($folderBtn)

# Clear button
$clearBtn = New-Object System.Windows.Forms.Button
$clearBtn.Text = Get-LocalizedString -Key "buttons.clear"
$clearBtn.Location = New-Object System.Drawing.Point(600, 125)
$clearBtn.Size = New-Object System.Drawing.Size(90, 28)
$clearBtn.Add_Click({
    $script:selectedFiles = @()
    $script:fileLabelCache = @{}  # Clear label cache (hashtable, not array!)
    $fileListBox.Items.Clear()
    Update-FileCount
    Adjust-UILayout  # Reset UI to default size
    $statusLabel.Text = Get-LocalizedString -Key "status.selectionCleared"
    $statusLabel.ForeColor = [System.Drawing.Color]::Gray
})
$fileGroupBox.Controls.Add($clearBtn)

# File count label
$fileCountLabel = New-Object System.Windows.Forms.Label
$fileCountLabel.Location = New-Object System.Drawing.Point(10, 90)
$fileCountLabel.Size = New-Object System.Drawing.Size(300, 20)
$fileCountLabel.Text = Get-LocalizedString -Key "labels.fileCount_none"
$fileCountLabel.ForeColor = [System.Drawing.Color]::Gray
$fileCountLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fileGroupBox.Controls.Add($fileCountLabel)

# ========================================
# DRAG-AND-DROP HANDLERS
# ========================================

# Helper function to process dropped items (files and folders)
function Process-DroppedItems {
    param([string[]]$DroppedPaths)
    
    $foundFiles = @()
    
    foreach ($path in $DroppedPaths) {
        if (Test-Path $path -PathType Container) {
            # It's a folder - scan it using centralized function
            try {
                $folderFiles = Get-SupportedFilesFromFolder -FolderPath $path -Recursive $includeSubfoldersCheckbox.Checked
                $foundFiles += $folderFiles
            }
            catch {
                Write-Log -Message "Could not scan dropped folder" -Level 'ERROR' -Source 'Process-DroppedItems' -Context @{ FolderPath = $path } -Exception $_
            }
        }
        elseif (Test-Path $path -PathType Leaf) {
            # It's a file - check if supported
            $ext = [System.IO.Path]::GetExtension($path).ToLower()
            if ($script:SupportedExtensions -contains $ext) {
                $foundFiles += Get-Item $path
            }
        }
    }
    
    # Remove duplicates (in case multiple folders or mixed files/folders)
    $foundFiles = $foundFiles | Sort-Object -Property FullName -Unique
    
    return $foundFiles
}

# Form DragEnter event - visual feedback
$form.Add_DragEnter({
    param($sender, $e)
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        
        # Visual feedback - change border color
        $form.BackColor = [System.Drawing.Color]::FromArgb(220, 240, 255)  # Light blue tint
    }
    else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

# Form DragDrop event - process dropped items
$form.Add_DragDrop({
    param($sender, $e)
    
    # Reset visual feedback
    $form.BackColor = [System.Drawing.SystemColors]::Control
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $droppedPaths = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        
        # Show processing message
        $statusLabel.Text = Get-LocalizedString -Key "status.processingDrop"
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        # Process dropped items
        $foundFiles = Process-DroppedItems -DroppedPaths $droppedPaths
        
        if ($foundFiles.Count -gt 0) {
            # Merge with existing selection using centralized function
            $mergeResult = Merge-FileSelection -NewFiles $foundFiles
            
            if ($mergeResult.NewCount -gt 0) {
                $script:selectedFiles = $mergeResult.MergedFiles
                
                # Update display
                $statusLabel.Text = Get-LocalizedString -Key "status.retrievingLabels"
                $form.Refresh()
                Update-FileListDisplay
                
                Update-FileCount
                $statusLabel.Text = Get-LocalizedString -Key "status.dragDropComplete" -Parameters @($mergeResult.TotalCount, $mergeResult.NewCount)
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                Write-Log -Message "Drag-and-drop completed (Form)" -Level 'INFO' -Source 'DragDrop' -Context @{ FilesFound = $mergeResult.TotalCount; NewFiles = $mergeResult.NewCount }
            } else {
                $statusLabel.Text = Get-LocalizedString -Key "status.noNewFiles"
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            }
        } else {
            $statusLabel.Text = Get-LocalizedString -Key "status.noSupportedFiles"
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        }
    }
})

# Form DragLeave event - reset visual feedback
$form.Add_DragLeave({
    param($sender, $e)
    $form.BackColor = [System.Drawing.SystemColors]::Control
})

# FileListBox DragEnter event - visual feedback
$fileListBox.Add_DragEnter({
    param($sender, $e)
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        
        # Visual feedback - change border color
        $fileListBox.BackColor = [System.Drawing.Color]::FromArgb(220, 240, 255)  # Light blue tint
    }
    else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

# FileListBox DragDrop event - process dropped items
$fileListBox.Add_DragDrop({
    param($sender, $e)
    
    # Reset visual feedback
    $fileListBox.BackColor = [System.Drawing.Color]::White
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $droppedPaths = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        
        # Show processing message
        $statusLabel.Text = Get-LocalizedString -Key "status.processingDrop"
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        # Process dropped items
        $foundFiles = Process-DroppedItems -DroppedPaths $droppedPaths
        
        if ($foundFiles.Count -gt 0) {
            # Merge with existing selection using centralized function
            $mergeResult = Merge-FileSelection -NewFiles $foundFiles
            
            if ($mergeResult.NewCount -gt 0) {
                $script:selectedFiles = $mergeResult.MergedFiles
                
                # Update display
                $statusLabel.Text = Get-LocalizedString -Key "status.retrievingLabels"
                $form.Refresh()
                Update-FileListDisplay
                
                Update-FileCount
                $statusLabel.Text = Get-LocalizedString -Key "status.dragDropComplete" -Parameters @($mergeResult.TotalCount, $mergeResult.NewCount)
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                Write-Log -Message "Drag-and-drop completed (ListBox)" -Level 'INFO' -Source 'DragDrop' -Context @{ FilesFound = $mergeResult.TotalCount; NewFiles = $mergeResult.NewCount }
            } else {
                $statusLabel.Text = Get-LocalizedString -Key "status.noNewFiles"
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            }
        } else {
            $statusLabel.Text = Get-LocalizedString -Key "status.noSupportedFiles"
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        }
    }
})

# FileListBox DragLeave event - reset visual feedback
$fileListBox.Add_DragLeave({
    param($sender, $e)
    $fileListBox.BackColor = [System.Drawing.Color]::White
})

# ========================================
# LABEL SELECTION SECTION
# ========================================
$labelGroupBox = New-Object System.Windows.Forms.GroupBox
$labelGroupBox.Text = Get-LocalizedString -Key "groupBoxes.sensitivityLabel"
$labelGroupBox.Location = New-Object System.Drawing.Point(10, 190)  # Updated from 140: 10 + 170 (FileGroupBox height) + 10 (margin)
$labelGroupBox.Size = New-Object System.Drawing.Size(700, 140)
$form.Controls.Add($labelGroupBox)

$labelLabel = New-Object System.Windows.Forms.Label
$labelLabel.Text = Get-LocalizedString -Key "labels.selectLabel"
$labelLabel.Location = New-Object System.Drawing.Point(10, 25)
$labelLabel.Size = New-Object System.Drawing.Size(680, 20)
$labelGroupBox.Controls.Add($labelLabel)

# Create styled button-like radio buttons for each label
$selectedLabelId = $null
$labelButtons = @()
$yPos = 50
$xPos = 10

# Define colors for professional look
$defaultBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)  # Light gray
$hoverBackColor = [System.Drawing.Color]::FromArgb(220, 230, 242)    # Light blue
$selectedBackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)   # Windows blue
$defaultForeColor = [System.Drawing.Color]::FromArgb(51, 51, 51)     # Dark gray
$selectedForeColor = [System.Drawing.Color]::White

if ($labels -and $labels.Count -gt 0) {
    foreach ($label in $labels) {
        # Create a panel that acts as a button container
        $btnPanel = New-Object System.Windows.Forms.Panel
        $btnPanel.Location = New-Object System.Drawing.Point($xPos, $yPos)
        $btnPanel.Size = New-Object System.Drawing.Size(155, 35)
        $btnPanel.BackColor = $defaultBackColor
        $btnPanel.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnPanel.Tag = @{
            LabelId = $label.Id
            IsSelected = $false
            DefaultBackColor = $defaultBackColor
            DefaultForeColor = $defaultForeColor
            SelectedBackColor = $selectedBackColor
            SelectedForeColor = $selectedForeColor
            HoverBackColor = $hoverBackColor
        }
        
        # Add subtle border effect
        $btnPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        
        # Create label inside panel for text
        $btnLabel = New-Object System.Windows.Forms.Label
        $btnLabel.Text = $label.DisplayName
        $btnLabel.Location = New-Object System.Drawing.Point(5, 8)
        $btnLabel.Size = New-Object System.Drawing.Size(145, 19)
        $btnLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnLabel.ForeColor = $defaultForeColor
        $btnLabel.BackColor = [System.Drawing.Color]::Transparent
        $btnLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $btnLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnPanel.Controls.Add($btnLabel)
        
        # Click event for panel
        $btnPanel.Add_Click({
            param($sender, $e)
            
            # Deselect all buttons first
            foreach ($btn in $script:labelButtons) {
                $btn.BackColor = $btn.Tag.DefaultBackColor
                $btn.Controls[0].ForeColor = $btn.Tag.DefaultForeColor
                $btn.Tag.IsSelected = $false
            }
            
            # Select this button
            $sender.BackColor = $sender.Tag.SelectedBackColor
            $sender.Controls[0].ForeColor = $sender.Tag.SelectedForeColor
            $sender.Tag.IsSelected = $true
            $script:selectedLabelId = $sender.Tag.LabelId
            
            # Save last selected label if configured
            if ($script:appConfig.preferences.rememberLastLabel) {
                $script:appConfig.preferences.lastSelectedLabelId = $sender.Tag.LabelId
            }
        })
        
        # Click event for label inside
        $btnLabel.Add_Click({
            param($sender, $e)
            # Get the parent panel
            $panel = $sender.Parent
            
            # Deselect all buttons first
            foreach ($btn in $script:labelButtons) {
                $btn.BackColor = $btn.Tag.DefaultBackColor
                $btn.Controls[0].ForeColor = $btn.Tag.DefaultForeColor
                $btn.Tag.IsSelected = $false
            }
            
            # Select this button
            $panel.BackColor = $panel.Tag.SelectedBackColor
            $panel.Controls[0].ForeColor = $panel.Tag.SelectedForeColor
            $panel.Tag.IsSelected = $true
            $script:selectedLabelId = $panel.Tag.LabelId
            
            # Save last selected label if configured
            if ($script:appConfig.preferences.rememberLastLabel) {
                $script:appConfig.preferences.lastSelectedLabelId = $panel.Tag.LabelId
            }
        })
        
        # Hover effects for panel
        $btnPanel.Add_MouseEnter({
            param($sender, $e)
            if (-not $sender.Tag.IsSelected) {
                $sender.BackColor = $sender.Tag.HoverBackColor
            }
        })
        
        $btnPanel.Add_MouseLeave({
            param($sender, $e)
            if (-not $sender.Tag.IsSelected) {
                $sender.BackColor = $sender.Tag.DefaultBackColor
            }
        })
        
        # Hover effects for label
        $btnLabel.Add_MouseEnter({
            param($sender, $e)
            if (-not $sender.Parent.Tag.IsSelected) {
                $sender.Parent.BackColor = $sender.Parent.Tag.HoverBackColor
            }
        })
        
        $btnLabel.Add_MouseLeave({
            param($sender, $e)
            if (-not $sender.Parent.Tag.IsSelected) {
                $sender.Parent.BackColor = $sender.Parent.Tag.DefaultBackColor
            }
        })
        
        $labelGroupBox.Controls.Add($btnPanel)
        $labelButtons += $btnPanel
        
        $xPos += 165
        if ($xPos -gt 490) {
            $xPos = 10
            $yPos += 45
        }
    }
    
    # Pre-select last used label if configured
    if ($appConfig.preferences.rememberLastLabel -and $appConfig.preferences.lastSelectedLabelId) {
        $lastLabelId = $appConfig.preferences.lastSelectedLabelId
        $btnToSelect = $labelButtons | Where-Object { $_.Tag.LabelId -eq $lastLabelId }
        
        if ($btnToSelect) {
            $btnToSelect.BackColor = $btnToSelect.Tag.SelectedBackColor
            $btnToSelect.Controls[0].ForeColor = $btnToSelect.Tag.SelectedForeColor
            $btnToSelect.Tag.IsSelected = $true
            $script:selectedLabelId = $lastLabelId
            Write-Log -Message "Pre-selected last used label" -Level 'INFO' -Source 'LabelButtons' -Context @{ LabelId = $lastLabelId }
        }
    }
} else {
    $manualLabel = New-Object System.Windows.Forms.Label
    $manualLabel.Text = Get-LocalizedString -Key "labels.noLabels"
    $manualLabel.Location = New-Object System.Drawing.Point(10, 20)
    $manualLabel.Size = New-Object System.Drawing.Size(640, 40)
    $manualLabel.ForeColor = [System.Drawing.Color]::Red
    $labelGroupBox.Controls.Add($manualLabel)
}

# ========================================
# PROGRESS SECTION
# ========================================
$progressGroupBox = New-Object System.Windows.Forms.GroupBox
$progressGroupBox.Text = Get-LocalizedString -Key "groupBoxes.progress"
$progressGroupBox.Location = New-Object System.Drawing.Point(10, 340)  # Updated from 290: 10 + 170 (FileGroupBox) + 10 + 140 (LabelGroupBox) + 10
$progressGroupBox.Size = New-Object System.Drawing.Size(700, 100)
$form.Controls.Add($progressGroupBox)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 25)
$progressBar.Size = New-Object System.Drawing.Size(680, 25)
$progressBar.Style = 'Continuous'
$progressGroupBox.Controls.Add($progressBar)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 55)
$statusLabel.Size = New-Object System.Drawing.Size(680, 35)
$statusLabel.Text = Get-LocalizedString -Key "status.ready"
$statusLabel.ForeColor = [System.Drawing.Color]::Gray
$progressGroupBox.Controls.Add($statusLabel)

# ========================================
# SMART PRE-APPLY SUMMARY FUNCTIONS
# ========================================

function Analyze-LabelChanges {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Files,
        [Parameter(Mandatory=$true)]
        [string]$TargetLabelId,
        [Parameter(Mandatory=$true)]
        [array]$AllLabels
    )
    
    # Get target label info
    $targetLabel = $AllLabels | Where-Object { $_.Id -eq $TargetLabelId }
    $targetRank = if($targetLabel.Rank) { $targetLabel.Rank } else { 0 }
    
    # Initialize result hashtable
    $analysis = @{
        New = @()           # Files with no current label
        Upgrade = @()       # Files moving to higher sensitivity
        Downgrade = @()     # Files moving to lower sensitivity
        Same = @()          # Files already have target label
        Unchanged = @()     # Files with same rank, different label
        TotalFiles = $Files.Count
        TargetLabel = $targetLabel.DisplayName
        TargetRank = $targetRank
        RequiresProtection = $targetLabel.RequiresProtection -eq $true
    }
    
    # Analyze each file
    foreach ($file in $Files) {
        try {
            # Try to use cache first (optimization!)
            $currentLabelInfo = $null
            $currentLabelId = $null
            $currentRank = -1
            $currentDisplayName = "Ukjent"
            
            if ($script:fileLabelCache.ContainsKey($file)) {
                $cached = $script:fileLabelCache[$file]
                if ($cached -is [hashtable]) {
                    # New cache format with full info - use null-safe property access
                    $currentLabelId = if ($cached.ContainsKey('LabelId')) { $cached.LabelId } else { $null }
                    $currentRank = if ($cached.ContainsKey('Rank')) { $cached.Rank } else { -1 }
                    $currentDisplayName = if ($cached.ContainsKey('DisplayName')) { $cached.DisplayName } else { "Ukjent" }
                    
                    # Validate display name is not null/empty
                    if ([string]::IsNullOrEmpty($currentDisplayName)) {
                        $currentDisplayName = "Ukjent"
                    }
                    
                    # CRITICAL: Treat unknown/protected labels as HIGHEST security
                    # This ensures justification is required when changing from unknown labels
                    if ($currentDisplayName -like "*Ukjent*" -or $currentDisplayName -like "*beskyttet*" -or $currentDisplayName -like "*Feil*") {
                        $currentRank = 999  # Treat as highest possible rank
                        Write-Log -Message "Treating unknown/protected label as highest security" -Level 'INFO' -Source 'Analyze-LabelChanges' -Context @{ 
                            FilePath = $file
                            DisplayName = $currentDisplayName
                            AssignedRank = 999
                        }
                    }
                } else {
                    # Old cache format (just string) - need to lookup
                    $currentDisplayName = if ($cached) { $cached.ToString() } else { "Ukjent" }
                    
                    # CRITICAL: Treat unknown/protected labels as HIGHEST security
                    if ($currentDisplayName -like "*Ukjent*" -or $currentDisplayName -like "*beskyttet*" -or $currentDisplayName -like "*Feil*") {
                        $currentRank = 999  # Treat as highest possible rank
                        Write-Log -Message "Treating unknown/protected label as highest security (old format)" -Level 'INFO' -Source 'Analyze-LabelChanges' -Context @{ 
                            FilePath = $file
                            DisplayName = $currentDisplayName
                            AssignedRank = 999
                        }
                    } else {
                        # Find label by display name to get ID and rank
                        $currentLabelInfo = $AllLabels | Where-Object { $_.DisplayName -eq $currentDisplayName }
                        if ($currentLabelInfo) {
                            $currentLabelId = $currentLabelInfo.Id
                            $currentRank = if($currentLabelInfo.Rank) { $currentLabelInfo.Rank } else { 0 }
                        }
                    }
                }
            } else {
                # Cache miss - need to fetch (shouldn't happen if Update-FileListDisplay was called)
                $currentStatus = Get-AIPFileStatus -Path $file -ErrorAction SilentlyContinue
                if ($currentStatus -and $currentStatus.MainLabelId) {
                    $currentLabelId = $currentStatus.MainLabelId
                    $currentLabelInfo = $AllLabels | Where-Object { $_.Id -eq $currentLabelId }
                    $currentRank = if($currentLabelInfo.Rank) { $currentLabelInfo.Rank } else { 0 }
                    $currentDisplayName = if($currentLabelInfo) { $currentLabelInfo.DisplayName } else { "Ukjent" }
                    
                    # Update cache
                    $script:fileLabelCache[$file] = @{
                        DisplayName = $currentDisplayName
                        LabelId = $currentLabelId
                        Rank = $currentRank
                    }
                } else {
                    # No label
                    $currentDisplayName = "Ingen etikett"
                    $currentRank = -1
                }
            }
            
            # Categorize based on cached/retrieved info
            if (-not $currentLabelId) {
                # No current label
                $analysis.New += @{
                    File = $file
                    CurrentLabel = $currentDisplayName
                    CurrentRank = $currentRank
                }
            }
            elseif ($currentLabelId -eq $TargetLabelId) {
                # Same label
                $analysis.Same += @{
                    File = $file
                    CurrentLabel = $targetLabel.DisplayName
                    CurrentRank = $targetRank
                }
            }
            else {
                # Different label - check rank
                $fileInfo = @{
                    File = $file
                    CurrentLabel = $currentDisplayName
                    CurrentRank = $currentRank
                }
                
                if ($targetRank -gt $currentRank) {
                    # Upgrade
                    $analysis.Upgrade += $fileInfo
                }
                elseif ($targetRank -lt $currentRank) {
                    # Downgrade
                    $analysis.Downgrade += $fileInfo
                }
                else {
                    # Same rank, different label
                    $analysis.Unchanged += $fileInfo
                }
            }
        }
        catch {
            Write-Log -Message "Could not analyze file for label changes" -Level 'WARNING' -Source 'Analyze-LabelChanges' -Context @{ FilePath = $file } -Exception $_
        }
    }
    
    return $analysis
}

function Get-ChangeWarnings {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Analysis
    )
    
    $warnings = @()
    
    # Warning 1: Mass downgrades
    if ($Analysis.Downgrade.Count -ge 3) {
        $warnings += @{
            Severity = "High"
            Message = "Du er i ferd med å nedgradere $($Analysis.Downgrade.Count) filer til lavere følsomhet."
            Icon = "Warning"
        }
    }
    
    # Warning 2: No changes detected
    $willChange = $Analysis.New.Count + $Analysis.Upgrade.Count + $Analysis.Downgrade.Count
    if ($willChange -eq 0 -and $Analysis.Same.Count -gt 0) {
        $warnings += @{
            Severity = "Info"
            Message = "Alle filer har allerede valgt etikett. Ingen endringer vil bli gjort."
            Icon = "Information"
        }
    }
    
    # Warning 3: Large batch
    if ($Analysis.TotalFiles -gt 20) {
        $warnings += @{
            Severity = "Info"
            Message = "Du er i ferd med å merke $($Analysis.TotalFiles) filer. Dette kan ta litt tid."
            Icon = "Information"
        }
    }
    
    # Warning 4: Protection required
    if ($Analysis.RequiresProtection) {
        $warnings += @{
            Severity = "High"
            Message = "Valgt etikett krever beskyttelse. Du vil bli bedt om å angi tillatelser."
            Icon = "Shield"
        }
    }
    
    # Warning 5: Mixed changes
    if ($Analysis.Upgrade.Count -gt 0 -and $Analysis.Downgrade.Count -gt 0) {
        $warnings += @{
            Severity = "Medium"
            Message = "Både oppgraderinger ($($Analysis.Upgrade.Count)) og nedgraderinger ($($Analysis.Downgrade.Count)) vil utføres."
            Icon = "Information"
        }
    }
    
    # Always return an array (even if empty)
    return @() + $warnings
}

function Show-PreApplySummary {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Analysis,
        [Parameter(Mandatory=$false)]
        [array]$Warnings = @()
    )
    
    try {
        # Validate Analysis hashtable
        if (-not $Analysis) {
            throw "Analysis object is null"
        }
        
        # Ensure all required keys exist with defaults
        $requiredKeys = @('New', 'Upgrade', 'Downgrade', 'Same', 'Unchanged', 'TargetLabel')
        foreach ($key in $requiredKeys) {
            if (-not $Analysis.ContainsKey($key)) {
                Write-Log -Message "Analysis missing key: $key" -Level 'WARNING' -Source 'Show-PreApplySummary'
                # Add default empty array
                if ($key -eq 'TargetLabel') {
                    $Analysis[$key] = "Ukjent etikett"
                } else {
                    $Analysis[$key] = @()
                }
            }
        }
        
        # Create summary dialog
        $summaryForm = New-Object System.Windows.Forms.Form
        $summaryForm.Text = Get-LocalizedString -Key "summaryDialog.title"
        $summaryForm.Size = New-Object System.Drawing.Size(550, 500)
        $summaryForm.StartPosition = 'CenterScreen'
        $summaryForm.FormBorderStyle = 'FixedDialog'
        $summaryForm.MaximizeBox = $false
        $summaryForm.MinimizeBox = $false
        $summaryForm.TopMost = $true
        $summaryForm.Owner = $form
        $summaryForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        
        # Header label
        $targetLabelText = if ($Analysis.TargetLabel) { $Analysis.TargetLabel } else { "Ukjent" }
        $headerLabel = New-Object System.Windows.Forms.Label
        $headerLabel.Text = Get-LocalizedString -Key "summaryDialog.header" -Parameters @($targetLabelText)
    $headerLabel.Location = New-Object System.Drawing.Point(15, 15)
    $headerLabel.Size = New-Object System.Drawing.Size(510, 25)
    $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $summaryForm.Controls.Add($headerLabel)
    
    # Summary group box
    $summaryGroup = New-Object System.Windows.Forms.GroupBox
    $summaryGroup.Text = Get-LocalizedString -Key "summaryDialog.summary"
    $summaryGroup.Location = New-Object System.Drawing.Point(15, 50)
    $summaryGroup.Size = New-Object System.Drawing.Size(510, 140)
    $summaryForm.Controls.Add($summaryGroup)
    
    # Create summary text
    $yPos = 25
    $summaryItems = @(
        @{ LabelKey = "summaryDialog.new"; Count = $Analysis.New.Count; Color = [System.Drawing.Color]::Green },
        @{ LabelKey = "summaryDialog.upgrade"; Count = $Analysis.Upgrade.Count; Color = [System.Drawing.Color]::Blue },
        @{ LabelKey = "summaryDialog.downgrade"; Count = $Analysis.Downgrade.Count; Color = [System.Drawing.Color]::OrangeRed },
        @{ LabelKey = "summaryDialog.same"; Count = $Analysis.Same.Count; Color = [System.Drawing.Color]::Gray },
        @{ LabelKey = "summaryDialog.unchanged"; Count = $Analysis.Unchanged.Count; Color = [System.Drawing.Color]::DarkGray }
    )
    
    foreach ($item in $summaryItems) {
        # Label
        $itemLabel = New-Object System.Windows.Forms.Label
        $itemLabel.Text = Get-LocalizedString -Key $item.LabelKey
        $itemLabel.Location = New-Object System.Drawing.Point(15, $yPos)
        $itemLabel.Size = New-Object System.Drawing.Size(350, 18)
        $summaryGroup.Controls.Add($itemLabel)
        
        # Count
        $countLabel = New-Object System.Windows.Forms.Label
        $countLabel.Text = $item.Count
        $countLabel.Location = New-Object System.Drawing.Point(370, $yPos)
        $countLabel.Size = New-Object System.Drawing.Size(120, 18)
        $countLabel.ForeColor = $item.Color
        $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $summaryGroup.Controls.Add($countLabel)
        
        $yPos += 22
    }
    
    # Warnings section (if any)
    if ($Warnings -and $Warnings.Count -gt 0) {
        $warningsGroup = New-Object System.Windows.Forms.GroupBox
        $warningsGroup.Text = Get-LocalizedString -Key "summaryDialog.warnings"
        $warningsGroup.Location = New-Object System.Drawing.Point(15, 200)
        $warningsGroup.Size = New-Object System.Drawing.Size(510, 150)
        $warningsGroup.ForeColor = [System.Drawing.Color]::OrangeRed
        $summaryForm.Controls.Add($warningsGroup)
        
        # Create warnings list
        $warningsText = ""
        foreach ($warning in $Warnings) {
            $icon = switch ($warning.Severity) {
                "High" { "[!]" }
                "Medium" { "[i]" }
                "Info" { "[i]" }
                default { "[-]" }
            }
            $warningsText += "$icon $($warning.Message)`r`n`r`n"
        }
        
        $warningsLabel = New-Object System.Windows.Forms.Label
        $warningsLabel.Text = $warningsText.Trim()
        $warningsLabel.Location = New-Object System.Drawing.Point(15, 25)
        $warningsLabel.Size = New-Object System.Drawing.Size(480, 115)
        $warningsLabel.ForeColor = [System.Drawing.Color]::Black
        $warningsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $warningsGroup.Controls.Add($warningsLabel)
    }
    
    # Skip unchanged checkbox
    $skipCheckbox = New-Object System.Windows.Forms.CheckBox
    $skipCheckbox.Text = Get-LocalizedString -Key "summaryDialog.skipUnchanged"
    $skipCheckbox.Location = New-Object System.Drawing.Point(15, 365)
    $skipCheckbox.Size = New-Object System.Drawing.Size(510, 25)
    $skipCheckbox.Checked = $false
    $summaryForm.Controls.Add($skipCheckbox)
    
    # Buttons
    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = Get-LocalizedString -Key "summaryDialog.applyChanges"
    $applyButton.Location = New-Object System.Drawing.Point(150, 410)
    $applyButton.Size = New-Object System.Drawing.Size(120, 35)
    $applyButton.BackColor = [System.Drawing.Color]::LightGreen
    $applyButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $applyButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $summaryForm.Controls.Add($applyButton)
    $summaryForm.AcceptButton = $applyButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = Get-LocalizedString -Key "summaryDialog.cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(280, 410)
    $cancelButton.Size = New-Object System.Drawing.Size(120, 35)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $summaryForm.Controls.Add($cancelButton)
    $summaryForm.CancelButton = $cancelButton
    
        # Show dialog and return result
        $result = $summaryForm.ShowDialog()
        
        return @{
            Action = $result
            SkipUnchanged = $skipCheckbox.Checked
        }
        
    } catch {
        # === SUMMARY DIALOG CRASH HANDLER ===
        Write-Log -Message "CRITICAL: Summary dialog crashed" -Level 'CRITICAL' -Source 'Show-PreApplySummary' -Exception $_
        
        # Save crash details
        $crashInfo = @{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
            Function = "Show-PreApplySummary"
            Exception = $_.Exception.Message
            ExceptionType = $_.Exception.GetType().FullName
            StackTrace = $_.ScriptStackTrace
            AnalysisKeys = if ($Analysis) { $Analysis.Keys -join ", " } else { "null" }
        }
        
        $crashLogPath = Join-Path $logDirectory "SUMMARY_DIALOG_CRASH_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        $crashInfo | ConvertTo-Json -Depth 10 | Set-Content $crashLogPath -Encoding UTF8
        
        # Show error
        [System.Windows.Forms.MessageBox]::Show(
            "Feil i oppsummeringsdialog:`n`n$($_.Exception.Message)`n`nDetaljer: $crashLogPath",
            "Dialog feil",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        # Return cancel so apply doesn't proceed
        return @{
            Action = [System.Windows.Forms.DialogResult]::Cancel
            SkipUnchanged = $false
        }
    }
}

function Show-StatisticsDialog {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Statistics,
        [Parameter(Mandatory=$true)]
        [TimeSpan]$ElapsedTime,
        [Parameter(Mandatory=$true)]
        [string]$LogFilePath,
        [Parameter(Mandatory=$true)]
        [string]$LabelName
    )
    
    # Create statistics dialog
    $statsForm = New-Object System.Windows.Forms.Form
    $statsForm.Text = Get-LocalizedString -Key "statisticsDialog.title"
    $statsForm.Size = New-Object System.Drawing.Size(700, 550)
    $statsForm.StartPosition = 'CenterScreen'
    $statsForm.FormBorderStyle = 'Sizable'
    $statsForm.MinimumSize = New-Object System.Drawing.Size(600, 450)
    $statsForm.MaximizeBox = $true
    $statsForm.MinimizeBox = $false
    $statsForm.TopMost = $true
    $statsForm.Owner = $form
    $statsForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Header
    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = Get-LocalizedString -Key "statisticsDialog.header" -Parameters @($LabelName)
    $headerLabel.Location = New-Object System.Drawing.Point(20, 20)
    $headerLabel.Size = New-Object System.Drawing.Size(650, 30)
    $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $statsForm.Controls.Add($headerLabel)
    
    # Summary group
    $summaryGroup = New-Object System.Windows.Forms.GroupBox
    $summaryGroup.Text = Get-LocalizedString -Key "statisticsDialog.summary"
    $summaryGroup.Location = New-Object System.Drawing.Point(20, 60)
    $summaryGroup.Size = New-Object System.Drawing.Size(650, 120)
    $statsForm.Controls.Add($summaryGroup)
    
    $yPos = 25
    
    # Total processed
    $totalLabel = New-Object System.Windows.Forms.Label
    $totalLabel.Text = Get-LocalizedString -Key "statisticsDialog.totalProcessed"
    $totalLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $totalLabel.Size = New-Object System.Drawing.Size(200, 20)
    $summaryGroup.Controls.Add($totalLabel)
    
    $totalValue = New-Object System.Windows.Forms.Label
    $totalValue.Text = $Statistics.TotalProcessed
    $totalValue.Location = New-Object System.Drawing.Point(220, $yPos)
    $totalValue.Size = New-Object System.Drawing.Size(400, 20)
    $totalValue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $summaryGroup.Controls.Add($totalValue)
    
    $yPos += 25
    
    # Success count
    $successLabel = New-Object System.Windows.Forms.Label
    $successLabel.Text = Get-LocalizedString -Key "statisticsDialog.successful"
    $successLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $successLabel.Size = New-Object System.Drawing.Size(200, 20)
    $summaryGroup.Controls.Add($successLabel)
    
    $successValue = New-Object System.Windows.Forms.Label
    $successValue.Text = $Statistics.SuccessCount
    $successValue.Location = New-Object System.Drawing.Point(220, $yPos)
    $successValue.Size = New-Object System.Drawing.Size(400, 20)
    $successValue.ForeColor = [System.Drawing.Color]::Green
    $successValue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $summaryGroup.Controls.Add($successValue)
    
    $yPos += 25
    
    # Failure count
    $failureLabel = New-Object System.Windows.Forms.Label
    $failureLabel.Text = Get-LocalizedString -Key "statisticsDialog.failed"
    $failureLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $failureLabel.Size = New-Object System.Drawing.Size(200, 20)
    $summaryGroup.Controls.Add($failureLabel)
    
    $failureValue = New-Object System.Windows.Forms.Label
    $failureValue.Text = $Statistics.FailureCount
    $failureValue.Location = New-Object System.Drawing.Point(220, $yPos)
    $failureValue.Size = New-Object System.Drawing.Size(400, 20)
    $failureValue.ForeColor = if($Statistics.FailureCount -gt 0) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Gray }
    $failureValue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $summaryGroup.Controls.Add($failureValue)
    
    $yPos += 25
    
    # Elapsed time
    $timeLabel = New-Object System.Windows.Forms.Label
    $timeLabel.Text = Get-LocalizedString -Key "statisticsDialog.timeUsed"
    $timeLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $timeLabel.Size = New-Object System.Drawing.Size(200, 20)
    $summaryGroup.Controls.Add($timeLabel)
    
    $timeValue = New-Object System.Windows.Forms.Label
    if ($ElapsedTime.TotalSeconds -lt 60) {
        $timeValue.Text = Get-LocalizedString -Key "statisticsDialog.seconds" -Parameters @([math]::Round($ElapsedTime.TotalSeconds, 1))
    } else {
        $timeValue.Text = Get-LocalizedString -Key "statisticsDialog.minutesSeconds" -Parameters @($ElapsedTime.Minutes, $ElapsedTime.Seconds)
    }
    $timeValue.Location = New-Object System.Drawing.Point(220, $yPos)
    $timeValue.Size = New-Object System.Drawing.Size(400, 20)
    $timeValue.ForeColor = [System.Drawing.Color]::Blue
    $timeValue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $summaryGroup.Controls.Add($timeValue)
    
    # Change type breakdown
    $breakdownGroup = New-Object System.Windows.Forms.GroupBox
    $breakdownGroup.Text = Get-LocalizedString -Key "statisticsDialog.changesByType"
    $breakdownGroup.Location = New-Object System.Drawing.Point(20, 190)
    $breakdownGroup.Size = New-Object System.Drawing.Size(650, 160)
    $statsForm.Controls.Add($breakdownGroup)
    
    $yPos = 25
    $changeTypes = @(
        @{ LabelKey = "statisticsDialog.new"; Key = "New"; Color = [System.Drawing.Color]::Green },
        @{ LabelKey = "statisticsDialog.upgrade"; Key = "Upgrade"; Color = [System.Drawing.Color]::Blue },
        @{ LabelKey = "statisticsDialog.downgrade"; Key = "Downgrade"; Color = [System.Drawing.Color]::OrangeRed },
        @{ LabelKey = "statisticsDialog.unchanged"; Key = "Unchanged"; Color = [System.Drawing.Color]::DarkGray },
        @{ LabelKey = "statisticsDialog.same"; Key = "Same"; Color = [System.Drawing.Color]::Gray }
    )
    
    foreach ($type in $changeTypes) {
        $count = $Statistics.ChangeTypeBreakdown[$type.Key]
        if ($count -gt 0) {
            $typeLabel = New-Object System.Windows.Forms.Label
            $typeLabel.Text = Get-LocalizedString -Key $type.LabelKey
            $typeLabel.Location = New-Object System.Drawing.Point(20, $yPos)
            $typeLabel.Size = New-Object System.Drawing.Size(250, 20)
            $breakdownGroup.Controls.Add($typeLabel)
            
            $typeValue = New-Object System.Windows.Forms.Label
            $typeValue.Text = $count
            $typeValue.Location = New-Object System.Drawing.Point(280, $yPos)
            $typeValue.Size = New-Object System.Drawing.Size(350, 20)
            $typeValue.ForeColor = $type.Color
            $typeValue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $breakdownGroup.Controls.Add($typeValue)
            
            $yPos += 25
        }
    }
    
    # Log file link
    $logLink = New-Object System.Windows.Forms.LinkLabel
    $logLink.Text = Get-LocalizedString -Key "statisticsDialog.openLog"
    $logLink.Location = New-Object System.Drawing.Point(20, 370)
    $logLink.Size = New-Object System.Drawing.Size(650, 20)
    $logLink.Add_LinkClicked({
        if (Test-Path $LogFilePath) {
            Start-Process notepad.exe -ArgumentList $LogFilePath
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedString -Key "statisticsDialog.logNotFound"),
                (Get-LocalizedString -Key "statisticsDialog.error"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    $statsForm.Controls.Add($logLink)
    
    # Buttons
    $exportBtn = New-Object System.Windows.Forms.Button
    $exportBtn.Text = Get-LocalizedString -Key "statisticsDialog.exportCSV"
    $exportBtn.Location = New-Object System.Drawing.Point(150, 420)
    $exportBtn.Size = New-Object System.Drawing.Size(150, 35)
    $exportBtn.Add_Click({
        # Will implement CSV export
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveDialog.FileName = "LabelResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                # Create CSV data
                $csvData = @()
                
                # Add summary header rows
                $csvData += [PSCustomObject]@{
                    FilePath = "=== SUMMARY ==="
                    OriginalLabel = ""
                    NewLabel = $LabelName
                    ChangeType = ""
                    Status = ""
                    Timestamp = $Statistics.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                    Message = "Label Application Session"
                }
                $csvData += [PSCustomObject]@{
                    FilePath = "Total Processed"
                    OriginalLabel = ""
                    NewLabel = ""
                    ChangeType = ""
                    Status = $Statistics.TotalProcessed
                    Timestamp = ""
                    Message = ""
                }
                $csvData += [PSCustomObject]@{
                    FilePath = "Success"
                    OriginalLabel = ""
                    NewLabel = ""
                    ChangeType = ""
                    Status = $Statistics.SuccessCount
                    Timestamp = ""
                    Message = ""
                }
                $csvData += [PSCustomObject]@{
                    FilePath = "Failure"
                    OriginalLabel = ""
                    NewLabel = ""
                    ChangeType = ""
                    Status = $Statistics.FailureCount
                    Timestamp = ""
                    Message = ""
                }
                $csvData += [PSCustomObject]@{
                    FilePath = "Time Elapsed"
                    OriginalLabel = ""
                    NewLabel = ""
                    ChangeType = ""
                    Status = ""
                    Timestamp = ""
                    Message = $ElapsedTime.ToString()
                }
                $csvData += [PSCustomObject]@{
                    FilePath = ""
                    OriginalLabel = ""
                    NewLabel = ""
                    ChangeType = ""
                    Status = ""
                    Timestamp = ""
                    Message = ""
                }
                
                # Add processed files header
                $csvData += [PSCustomObject]@{
                    FilePath = "=== PROCESSED FILES ==="
                    OriginalLabel = ""
                    NewLabel = ""
                    ChangeType = ""
                    Status = ""
                    Timestamp = ""
                    Message = ""
                }
                
                # Add processed files
                foreach ($fileInfo in $Statistics.ProcessedFiles) {
                    $csvData += [PSCustomObject]@{
                        FilePath = $fileInfo.FilePath
                        OriginalLabel = $fileInfo.OriginalLabel
                        NewLabel = $fileInfo.NewLabel
                        ChangeType = $fileInfo.ChangeType
                        Status = $fileInfo.Status
                        Timestamp = $fileInfo.Timestamp
                        Message = $fileInfo.Message
                    }
                }
                
                # Add failed files section if any
                if ($Statistics.FailedFiles.Count -gt 0) {
                    $csvData += [PSCustomObject]@{
                        FilePath = ""
                        OriginalLabel = ""
                        NewLabel = ""
                        ChangeType = ""
                        Status = ""
                        Timestamp = ""
                        Message = ""
                    }
                    $csvData += [PSCustomObject]@{
                        FilePath = "=== FAILED FILES ==="
                        OriginalLabel = ""
                        NewLabel = ""
                        ChangeType = ""
                        Status = ""
                        Timestamp = ""
                        Message = ""
                    }
                    
                    foreach ($fileInfo in $Statistics.FailedFiles) {
                        $csvData += [PSCustomObject]@{
                            FilePath = $fileInfo.FilePath
                            OriginalLabel = $fileInfo.OriginalLabel
                            NewLabel = $fileInfo.NewLabel
                            ChangeType = $fileInfo.ChangeType
                            Status = "Failed"
                            Timestamp = $fileInfo.Timestamp
                            Message = $fileInfo.Error
                        }
                    }
                }
                
                # Export with UTF-8 BOM for Norwegian characters
                $csvData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                
                [System.Windows.Forms.MessageBox]::Show(
                    (Get-LocalizedString -Key "statisticsDialog.exportSuccess" -Parameters @($saveDialog.FileName)),
                    (Get-LocalizedString -Key "statisticsDialog.exportSuccessTitle"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    (Get-LocalizedString -Key "statisticsDialog.exportError" -Parameters @($_.Exception.Message)),
                    (Get-LocalizedString -Key "statisticsDialog.exportErrorTitle"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    })
    $statsForm.Controls.Add($exportBtn)
    
    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = Get-LocalizedString -Key "statisticsDialog.close"
    $closeBtn.Location = New-Object System.Drawing.Point(420, 420)
    $closeBtn.Size = New-Object System.Drawing.Size(120, 35)
    $closeBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $statsForm.Controls.Add($closeBtn)
    $statsForm.AcceptButton = $closeBtn
    
    # Show dialog
    $statsForm.ShowDialog() | Out-Null
}

# ========================================
# ACTION BUTTONS - CENTERED LAYOUT
# ========================================
# Calculate centered positions for 3 buttons
# Total width: 200 (Apply) + 10 (spacing) + 100 (Log) + 10 (spacing) + 100 (Settings) = 420px
# Form content width: 720px (740 - 20 for borders)
# Centered start: (720 - 420) / 2 = 150px
$buttonCenterStartX = 150
$buttonInitialY = 455  # Will be adjusted by Adjust-UILayout

$applyBtn = New-Object System.Windows.Forms.Button
$applyBtn.Text = Get-LocalizedString -Key "buttons.apply"
$applyBtn.Location = New-Object System.Drawing.Point($buttonCenterStartX, $buttonInitialY)
$applyBtn.Size = New-Object System.Drawing.Size(200, 35)
$applyBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$applyBtn.BackColor = [System.Drawing.Color]::LightGreen
$applyBtn.Add_Click({
    try {
        # === MASTER TRY/CATCH WRAPPER ===
        # Catches ANY unhandled exception during label application
        # Prevents application crash and provides detailed error reporting
        
        if ($script:selectedFiles.Count -eq 0) {
            $statusLabel.Text = Get-LocalizedString -Key "errors.noFilesSelected"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            return
        }
        
        if (-not $script:selectedLabelId) {
            $statusLabel.Text = Get-LocalizedString -Key "errors.noLabelSelected"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            return
        }
    
    $labelId = $script:selectedLabelId
    
    # Get selected label object with rank
    $selectedLabelObj = $labels | Where-Object { $_.Id -eq $labelId }
    $newRank = if($selectedLabelObj.Rank) { $selectedLabelObj.Rank } else { 0 }
    $requiresProtection = $selectedLabelObj.RequiresProtection -eq $true
    
    # ========================================
    # SMART PRE-APPLY SUMMARY
    # ========================================
    
    # Show "Analyzing..." status
    $statusLabel.Text = Get-LocalizedString -Key "status.analyzingChanges"
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    try {
        # Analyze label changes
        $analysis = Analyze-LabelChanges -Files $script:selectedFiles -TargetLabelId $labelId -AllLabels $labels
        $warnings = Get-ChangeWarnings -Analysis $analysis
        
        # Show summary dialog
        try {
            $summaryResult = Show-PreApplySummary -Analysis $analysis -Warnings $warnings
        }
        catch {
            Write-Log -Message "Could not show pre-apply summary dialog" -Level 'ERROR' -Source 'ApplyButton' -Exception $_
            [System.Windows.Forms.MessageBox]::Show(
                "Kunne ikke vise oppsummeringsdialog. Vil du fortsette uten oppsummering?`n`nFeil: $($_.Exception.Message)",
                "Dialogfeil",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) -eq [System.Windows.Forms.DialogResult]::No -and { return }
            
            # If user chose Yes, create a minimal result to continue
            $summaryResult = @{
                Action = [System.Windows.Forms.DialogResult]::OK
                SkipUnchanged = $false
            }
        }
        
        # Check if user cancelled
        if ($summaryResult.Action -ne [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = Get-LocalizedString -Key "status.labellingCancelled"
            $statusLabel.ForeColor = [System.Drawing.Color]::Gray
            Write-Log -Message "Label application cancelled by user in summary dialog" -Level 'INFO' -Source 'ApplyButton'
            return
        }
        
        # Filter files based on "Skip unchanged" option
        $filesToProcess = $script:selectedFiles
        if ($summaryResult.SkipUnchanged) {
            # Skip files in "Same" category
            $sameFiles = $analysis.Same | ForEach-Object { $_.File }
            $filesToProcess = $script:selectedFiles | Where-Object { $sameFiles -notcontains $_ }
            
            if ($filesToProcess.Count -eq 0) {
                $statusLabel.Text = Get-LocalizedString -Key "status.noFilesToProcess"
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
                Write-Log -Message "No files to process - all have selected label already" -Level 'INFO' -Source 'ApplyButton'
                return
            }
        }
    }
    catch {
        $statusLabel.Text = Get-LocalizedString -Key "errors.couldNotAnalyzeFiles" -Parameters @($_.Exception.Message)
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        Write-Log -Message "File analysis failed" -Level 'ERROR' -Source 'ApplyButton' -Exception $_
        return
    }
    
    # ========================================
    # ORIGINAL APPLY LOGIC (adapted)
    # ========================================
    
    # Check if we're downgrading any files (need justification)
    $needsJustification = $analysis.Downgrade.Count -gt 0
    
    # Ask for justification only if downgrading
    $userJustification = "Endret via massemerking"
    if ($needsJustification) {
    $justificationForm = New-Object System.Windows.Forms.Form
    $justificationForm.Text = 'Begrunnelse påkrevd'
    $justificationForm.Size = New-Object System.Drawing.Size(400, 180)
    $justificationForm.StartPosition = 'CenterScreen'
    $justificationForm.FormBorderStyle = 'FixedDialog'
    $justificationForm.MaximizeBox = $false
    $justificationForm.MinimizeBox = $false
    $justificationForm.TopMost = $true
    $justificationForm.Owner = $form
    
    $justLabel = New-Object System.Windows.Forms.Label
    $justLabel.Text = 'Begrunnelse for nedgradering av følsomhetsetikett:'
    $justLabel.Location = New-Object System.Drawing.Point(10, 10)
    $justLabel.Size = New-Object System.Drawing.Size(360, 40)
    $justificationForm.Controls.Add($justLabel)
    
    $justTextBox = New-Object System.Windows.Forms.TextBox
    $justTextBox.Location = New-Object System.Drawing.Point(10, 50)
    $justTextBox.Size = New-Object System.Drawing.Size(360, 40)
    $justTextBox.Multiline = $true
    $justTextBox.Text = 'Endret via massemerking'
    $justificationForm.Controls.Add($justTextBox)
    
    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = 'OK'
    $okBtn.Location = New-Object System.Drawing.Point(200, 100)
    $okBtn.Size = New-Object System.Drawing.Size(80, 30)
    $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $justificationForm.Controls.Add($okBtn)
    $justificationForm.AcceptButton = $okBtn
    
    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = 'Avbryt'
    $cancelBtn.Location = New-Object System.Drawing.Point(290, 100)
    $cancelBtn.Size = New-Object System.Drawing.Size(80, 30)
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $justificationForm.Controls.Add($cancelBtn)
    $justificationForm.CancelButton = $cancelBtn
    
    $result = $justificationForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return
    }
    
    $userJustification = $justTextBox.Text
    }
    
    # Handle labels that require protection (optional setup)
    $protectionSettings = $null
    if ($requiresProtection) {
        $protectForm = New-Object System.Windows.Forms.Form
        $protectForm.Text = Get-LocalizedString -Key "protectionDialog.title"
        $protectForm.Size = New-Object System.Drawing.Size(500, 380)
        $protectForm.StartPosition = 'CenterScreen'
        $protectForm.FormBorderStyle = 'FixedDialog'
        $protectForm.MaximizeBox = $false
        $protectForm.MinimizeBox = $false
        $protectForm.TopMost = $true
        $protectForm.Owner = $form
        
        $protectLabel = New-Object System.Windows.Forms.Label
        $protectLabel.Text = Get-LocalizedString -Key "protectionDialog.description"
        $protectLabel.Location = New-Object System.Drawing.Point(10, 10)
        $protectLabel.Size = New-Object System.Drawing.Size(460, 40)
        $protectForm.Controls.Add($protectLabel)
        
        # Permission level dropdown
        $permLabel = New-Object System.Windows.Forms.Label
        $permLabel.Text = Get-LocalizedString -Key "protectionDialog.selectPermission"
        $permLabel.Location = New-Object System.Drawing.Point(10, 60)
        $permLabel.Size = New-Object System.Drawing.Size(460, 20)
        $protectForm.Controls.Add($permLabel)
        
        $permDropdown = New-Object System.Windows.Forms.ComboBox
        $permDropdown.Location = New-Object System.Drawing.Point(10, 85)
        $permDropdown.Size = New-Object System.Drawing.Size(460, 25)
        $permDropdown.DropDownStyle = 'DropDownList'
        $permDropdown.Items.AddRange(@(
            (Get-LocalizedString -Key "protectionDialog.viewer"),
            (Get-LocalizedString -Key "protectionDialog.reviewer"),
            (Get-LocalizedString -Key "protectionDialog.coauthor"),
            (Get-LocalizedString -Key "protectionDialog.coowner"),
            (Get-LocalizedString -Key "protectionDialog.justMe")
        ))
        $permDropdown.SelectedIndex = 4  # Default to "Bare for meg"
        $protectForm.Controls.Add($permDropdown)
        
        # Email textbox (only shown for non-"Bare for meg" options)
        $emailLabel = New-Object System.Windows.Forms.Label
        $emailLabel.Text = Get-LocalizedString -Key "protectionDialog.enterUsers"
        $emailLabel.Location = New-Object System.Drawing.Point(10, 125)
        $emailLabel.Size = New-Object System.Drawing.Size(460, 20)
        $emailLabel.Visible = $false
        $protectForm.Controls.Add($emailLabel)
        
        $emailBox = New-Object System.Windows.Forms.TextBox
        $emailBox.Location = New-Object System.Drawing.Point(10, 150)
        $emailBox.Size = New-Object System.Drawing.Size(460, 60)
        $emailBox.Multiline = $true
        $emailBox.Text = ""
        $emailBox.Visible = $false
        $protectForm.Controls.Add($emailBox)
        
        # Placeholder label for email box
        $placeholderLabel = New-Object System.Windows.Forms.Label
        $placeholderLabel.Text = Get-LocalizedString -Key "protectionDialog.example"
        $placeholderLabel.Location = New-Object System.Drawing.Point(13, 153)
        $placeholderLabel.Size = New-Object System.Drawing.Size(450, 15)
        $placeholderLabel.ForeColor = [System.Drawing.Color]::Gray
        $placeholderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
        $placeholderLabel.BackColor = [System.Drawing.Color]::White
        $placeholderLabel.Visible = $false
        $protectForm.Controls.Add($placeholderLabel)
        
        # Expiration checkbox and date picker
        $expiryCheck = New-Object System.Windows.Forms.CheckBox
        $expiryCheck.Text = Get-LocalizedString -Key "protectionDialog.documentExpires"
        $expiryCheck.Location = New-Object System.Drawing.Point(10, 225)
        $expiryCheck.Size = New-Object System.Drawing.Size(460, 20)
        $expiryCheck.Checked = $false
        $protectForm.Controls.Add($expiryCheck)
        
        $expiryDate = New-Object System.Windows.Forms.DateTimePicker
        $expiryDate.Location = New-Object System.Drawing.Point(30, 250)
        $expiryDate.Size = New-Object System.Drawing.Size(430, 25)
        $expiryDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
        $expiryDate.MinDate = [DateTime]::Now.AddDays(1)
        $expiryDate.Value = [DateTime]::Now.AddMonths(3)
        $expiryDate.Enabled = $false
        $protectForm.Controls.Add($expiryDate)
        
        # Enable/disable email box based on permission selection
        $permDropdown.Add_SelectedIndexChanged({
            if ($permDropdown.SelectedIndex -eq 4) {
                # "Bare for meg" selected
                $emailLabel.Visible = $false
                $emailBox.Visible = $false
                $placeholderLabel.Visible = $false
            } else {
                # Other permissions require email input
                $emailLabel.Visible = $true
                $emailBox.Visible = $true
                $placeholderLabel.Visible = $true
            }
        })
        
        # Toggle expiry date picker
        $expiryCheck.Add_CheckedChanged({
            $expiryDate.Enabled = $expiryCheck.Checked
        })
        
        # Hide placeholder when user types
        $emailBox.Add_TextChanged({
            if ($emailBox.Text.Length -gt 0) {
                $placeholderLabel.Visible = $false
            } else {
                $placeholderLabel.Visible = $true
            }
        })
        
        # OK button
        $okProtectBtn = New-Object System.Windows.Forms.Button
        $okProtectBtn.Text = Get-LocalizedString -Key "protectionDialog.ok"
        $okProtectBtn.Location = New-Object System.Drawing.Point(290, 300)
        $okProtectBtn.Size = New-Object System.Drawing.Size(90, 35)
        $okProtectBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $protectForm.Controls.Add($okProtectBtn)
        $protectForm.AcceptButton = $okProtectBtn
        
        # Cancel button
        $cancelProtectBtn = New-Object System.Windows.Forms.Button
        $cancelProtectBtn.Text = Get-LocalizedString -Key "protectionDialog.cancel"
        $cancelProtectBtn.Location = New-Object System.Drawing.Point(390, 300)
        $cancelProtectBtn.Size = New-Object System.Drawing.Size(90, 35)
        $cancelProtectBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $protectForm.Controls.Add($cancelProtectBtn)
        $protectForm.CancelButton = $cancelProtectBtn
        
        $protectResult = $protectForm.ShowDialog()
        if ($protectResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPermission = $permDropdown.SelectedIndex
            
            # Map selection to permission type
            $protectionSettings = @{
                PermissionType = $selectedPermission
                Emails = $emailBox.Text.Trim()
                HasExpiry = $expiryCheck.Checked
                ExpiryDate = $expiryDate.Value
            }
        } elseif ($protectResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
            # User cancelled
            [System.Windows.Forms.MessageBox]::Show(
                (Get-LocalizedString -Key "protectionDialog.cancelled"),
                (Get-LocalizedString -Key "protectionDialog.cancelledTitle"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }
    }
    
    # Initialize counters
    $totalFiles = $filesToProcess.Count
    $successCount = 0
    $failureCount = 0
    $currentFile = 0
    
    # Initialize stopwatch for timing
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Initialize statistics tracking
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
        FailedFiles = @()
        ProcessedFiles = @()
        StartTime = Get-Date
    }
    
    # Initialize log
    Write-Log -Message "==========================================" -Level 'INFO'
    Write-Log -Message "Label application session started" -Level 'INFO' -Source 'LabelApplication'
    Write-Log -Message "Target label: $($selectedLabelObj.DisplayName) (ID: $labelId)" -Level 'INFO' -Source 'LabelApplication'
    Write-Log -Message "Total files to process: $totalFiles" -Level 'INFO' -Source 'LabelApplication'
    if ($summaryResult.SkipUnchanged) {
        Write-Log -Message "Skipping $($analysis.Same.Count) files with unchanged labels" -Level 'INFO' -Source 'LabelApplication'
    }
    Write-Log -Message "==========================================" -Level 'INFO'
    
    # Disable buttons during processing
    $applyBtn.Enabled = $false
    $browseBtn.Enabled = $false
    $clearBtn.Enabled = $false
    foreach ($btn in $labelButtons) {
        $btn.Enabled = $false
    }
    
    # ========================================
    # ASYNC vs SYNC DECISION FOR LABEL APPLICATION
    # ========================================
    # Threshold: >=30 files = async, <30 files = sync
    # - Async overhead (runspace setup, marshaling): ~2-3 seconds
    # - Break-even point: ~25-30 files (tested empirically)
    # - Sync mode benefits: simpler error handling, immediate feedback
    # - Async mode benefits: 3-4x faster for large batches, responsive UI
    # - Performance: 50 files sync=40s, async=12s; 100 files sync=80s, async=22s
    $useAsync = ($totalFiles -ge 30 -and $script:runspacePool -ne $null)
    
    if ($useAsync) {
        Write-Log -Message "Using ASYNC mode for label application" -Level 'INFO' -Source 'ApplyButton' -Context @{ TotalFiles = $totalFiles }
        
        # Prepare shared statistics (thread-safe counters)
        $sharedStats = @{
            TotalProcessed = 0
            SuccessCount = 0
            FailureCount = 0
            ChangeTypeBreakdown_New = 0
            ChangeTypeBreakdown_Upgrade = 0
            ChangeTypeBreakdown_Downgrade = 0
            ChangeTypeBreakdown_Unchanged = 0
            ChangeTypeBreakdown_Same = 0
        }
        
        try {
            # Start async batch label application
            $labelJobs = Start-AsyncBatchLabelApplication -FilesToProcess $filesToProcess `
                                                           -LabelId $labelId `
                                                           -Analysis $analysis `
                                                           -RunspacePool $script:runspacePool `
                                                           -SharedStats $sharedStats `
                                                           -SelectedLabelObj $selectedLabelObj `
                                                           -Justification $userJustification `
                                                           -ProtectionSettings $protectionSettings `
                                                           -RequiresProtection $requiresProtection
            
            # Wait for completion with UI updates
            $script:asyncSharedProgress.Processed = 0
            $script:asyncSharedProgress.Total = $totalFiles
            
            $asyncResults = Wait-AsyncJobsWithUI -Jobs $labelJobs `
                                                  -SharedProgress $script:asyncSharedProgress `
                                                  -ProgressBar $progressBar `
                                                  -StatusLabel $statusLabel `
                                                  -Form $form `
                                                  -OperationType "Påfører etiketter"
            
            # Process results and build stats
            foreach ($result in $asyncResults) {
                # Validate result object
                if (-not $result) {
                    Write-Log -Message "Received null result from async job" -Level 'WARNING' -Source 'ApplyButton'
                    continue
                }
                
                # If result has no FilePath, it's a malformed result - skip logging but stats already updated
                if (-not $result.FilePath) {
                    Write-Log -Message "Result missing FilePath property - skipping file log but stats already counted" -Level 'WARNING' -Source 'ApplyButton' -Context @{
                        Success = $result.Success
                        ErrorMessage = if ($result.ErrorMessage) { $result.ErrorMessage } else { "None" }
                    }
                    # Don't continue - still try to process what we can
                    # Stats were already incremented in runspace, so we're just logging here
                    # But skip adding to ProcessedFiles/FailedFiles arrays since we don't have FilePath
                    continue
                }
                
                if ($result.Success) {
                    $stats.ProcessedFiles += @{
                        FilePath = $result.FilePath
                        OriginalLabel = if ($result.OriginalLabel) { $result.OriginalLabel } else { "Ukjent" }
                        NewLabel = if ($result.NewLabel) { $result.NewLabel } else { "Ukjent" }
                        ChangeType = if ($result.ChangeType) { $result.ChangeType } else { "Unknown" }
                        Status = "Success"
                        Timestamp = if ($result.Timestamp) { $result.Timestamp } else { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
                        Message = if ($result.Message) { $result.Message } else { "" }
                    }
                    
                    # Update cache inline (optimization)
                    if ($requiresProtection) {
                        $script:fileLabelCache[$result.FilePath] = @{
                            DisplayName = $selectedLabelObj.DisplayName
                            LabelId = $labelId
                            Rank = $newRank
                        }
                    } else {
                        $script:fileLabelCache[$result.FilePath] = $selectedLabelObj.DisplayName
                    }
                    
                    Write-Log -Message "Label applied successfully with protection" -Level 'INFO' -Source 'LabelApplication' -Context @{
                        FilePath = $result.FilePath
                        OriginalLabel = $result.OriginalLabel
                        NewLabel = $result.NewLabel
                        ChangeType = $result.ChangeType
                        Protection = $result.Message
                    }
                } else {
                    $stats.FailedFiles += @{
                        FilePath = $result.FilePath
                        OriginalLabel = if ($result.OriginalLabel) { $result.OriginalLabel } else { "Ukjent" }
                        NewLabel = if ($result.NewLabel) { $result.NewLabel } else { "Ukjent" }
                        ChangeType = if ($result.ChangeType) { $result.ChangeType } else { "Unknown" }
                        Timestamp = if ($result.Timestamp) { $result.Timestamp } else { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
                        Error = if ($result.ErrorMessage) { $result.ErrorMessage } else { "Unknown error" }
                    }
                    
                    Write-Log -Message "Label application failed" -Level 'ERROR' -Source 'LabelApplication' -Context @{
                        FilePath = if ($result.FilePath) { $result.FilePath } else { "Unknown" }
                        OriginalLabel = if ($result.OriginalLabel) { $result.OriginalLabel } else { "" }
                        ChangeType = if ($result.ChangeType) { $result.ChangeType } else { "" }
                        ErrorMessage = if ($result.ErrorMessage) { $result.ErrorMessage } else { "" }
                    }
                }
            }
            
            # Copy thread-safe counters to stats
            $stats.TotalProcessed = $sharedStats.TotalProcessed
            $stats.SuccessCount = $sharedStats.SuccessCount
            $stats.FailureCount = $sharedStats.FailureCount
            $stats.ChangeTypeBreakdown.New = $sharedStats.ChangeTypeBreakdown_New
            $stats.ChangeTypeBreakdown.Upgrade = $sharedStats.ChangeTypeBreakdown_Upgrade
            $stats.ChangeTypeBreakdown.Downgrade = $sharedStats.ChangeTypeBreakdown_Downgrade
            $stats.ChangeTypeBreakdown.Unchanged = $sharedStats.ChangeTypeBreakdown_Unchanged
            $stats.ChangeTypeBreakdown.Same = $sharedStats.ChangeTypeBreakdown_Same
            
            $successCount = $stats.SuccessCount
            $failureCount = $stats.FailureCount
            
        } catch {
            Write-Log -Message "Async label application failed, falling back to sync" -Level 'ERROR' -Source 'ApplyButton' -Exception $_
            [System.Windows.Forms.MessageBox]::Show(
                "Async label application feilet:`n`n$($_.Exception.Message)`n`nFaller tilbake til synkron modus.",
                "Async Feil",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            # Fall back to sync mode
            $useAsync = $false
        }
    }
    
    # ========================================
    # SYNCHRONOUS MODE (fallback or small batches)
    # ========================================
    if (-not $useAsync) {
        Write-Log -Message "Using SYNC mode for label application" -Level 'INFO' -Source 'ApplyButton' -Context @{ TotalFiles = $totalFiles }
        
        # Process each file
        try {
            foreach ($file in $filesToProcess) {
        $currentFile++
        $fileName = [System.IO.Path]::GetFileName($file)
        
        # Determine change type and original label for this file from analysis
        $changeType = "Unknown"
        $originalLabel = "Ukjent"
        
        $fileAnalysis = $analysis.New | Where-Object { $_.File -eq $file }
        if ($fileAnalysis) { 
            $changeType = "New"
            $originalLabel = $fileAnalysis.CurrentLabel
        }
        else {
            $fileAnalysis = $analysis.Upgrade | Where-Object { $_.File -eq $file }
            if ($fileAnalysis) { 
                $changeType = "Upgrade"
                $originalLabel = $fileAnalysis.CurrentLabel
            }
            else {
                $fileAnalysis = $analysis.Downgrade | Where-Object { $_.File -eq $file }
                if ($fileAnalysis) { 
                    $changeType = "Downgrade"
                    $originalLabel = $fileAnalysis.CurrentLabel
                }
                else {
                    $fileAnalysis = $analysis.Unchanged | Where-Object { $_.File -eq $file }
                    if ($fileAnalysis) { 
                        $changeType = "Unchanged"
                        $originalLabel = $fileAnalysis.CurrentLabel
                    }
                    else {
                        $fileAnalysis = $analysis.Same | Where-Object { $_.File -eq $file }
                        if ($fileAnalysis) { 
                            $changeType = "Same"
                            $originalLabel = $fileAnalysis.CurrentLabel
                        }
                    }
                }
            }
        }
        
        # Update progress
        $progressBar.Value = [int](($currentFile / $totalFiles) * 100)
        $statusLabel.Text = "Behandler ($currentFile/$totalFiles): $fileName"
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        try {
            # CRITICAL: Check if file is locked before attempting label application
            # This prevents crashes when files are open in Office applications
            try {
                $fileStream = [System.IO.File]::Open($file, 'Open', 'ReadWrite', 'None')
                $fileStream.Close()
            } catch [System.IO.IOException] {
                if ($_.Exception.Message -like "*being used by another process*" -or 
                    $_.Exception.Message -like "*file is in use*") {
                    # File is locked - skip it and continue
                    Write-Log -Message "File is locked by another process, skipping" -Level 'WARNING' -Source 'ApplyButton-SyncMode' -Context @{ FilePath = $file }
                    $failureCount++
                    $stats.FailureCount++
                    $stats.TotalProcessed++
                    $stats.FailedFiles += @{
                        FilePath = $file
                        OriginalLabel = $originalLabel
                        NewLabel = $selectedLabelObj.DisplayName
                        ChangeType = $changeType
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Error = "Filen er åpen i et annet program (låst)"
                    }
                    continue  # Skip to next file
                } else {
                    throw  # Other IO errors should be handled normally
                }
            }
            
            # Apply label - try without justification first
            try {
                # For protected labels with custom permissions
                if ($requiresProtection -and $protectionSettings) {
                    $permissionType = $protectionSettings.PermissionType
                    
                    # Get current user's email
                    $currentUserEmail = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    try {
                        $currentUserEmail = ([ADSI]"LDAP://<SID=$([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)>").mail
                        if (-not $currentUserEmail) {
                            $username = $env:USERNAME
                            $domain = $env:USERDNSDOMAIN
                            if ($domain) {
                                $currentUserEmail = "$username@$domain".ToLower()
                            }
                        }
                    }
                    catch {
                        $currentUserEmail = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    }
                    
                    # Determine users and permissions based on selection
                    if ($permissionType -eq 4) {
                        # "Bare for meg" - Owner only
                        $userList = @($currentUserEmail)
                        $permissionLevel = "CoOwner"
                        $permDesc = "bare for meg ($currentUserEmail)"
                    }
                    else {
                        # Other options require email input
                        if (-not $protectionSettings.Emails) {
                            Write-Log -Message "No users specified for selected permission" -Level 'ERROR' -Source 'ApplyButton-SyncMode' -Context @{ FilePath = $file }
                            $failureCount++
                            continue
                        }
                        
                        $userList = $protectionSettings.Emails -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
                        
                        switch ($permissionType) {
                            0 { $permissionLevel = "Viewer"; $permDesc = "leser" }      # Leser - Bare vise
                            1 { $permissionLevel = "Reviewer"; $permDesc = "kontrollør" } # Kontrollør - Vise, redigere
                            2 { $permissionLevel = "CoAuthor"; $permDesc = "medforfatter" } # Medforfatter
                            3 { $permissionLevel = "CoOwner"; $permDesc = "medeier" }   # Medeier - Alle tillatelser
                        }
                    }
                    
                    try {
                        Write-Log -Message "Applying label with protection" -Level 'INFO' -Source 'ApplyButton-SyncMode' -Context @{ FilePath = $file; Protection = $permDesc; Users = ($userList -join ', ') }
                        
                        # Create custom permissions
                        $customPermission = New-AIPCustomPermissions -Users $userList -Permissions $permissionLevel -ErrorAction Stop
                        
                        # Apply label with custom protection
                        Set-AIPFileLabel -Path $file -LabelId $labelId -CustomPermissions $customPermission -PreserveFileDetails -ErrorAction Stop
                        $successCount++
                        Write-Log "SUKSESS: $file (beskyttelse: $permDesc)"
                        
                        # Update statistics
                        $stats.SuccessCount++
                        $stats.TotalProcessed++
                        $stats.ChangeTypeBreakdown[$changeType]++
                        $stats.ProcessedFiles += @{
                            FilePath = $file
                            OriginalLabel = $originalLabel
                            NewLabel = $selectedLabelObj.DisplayName
                            ChangeType = $changeType
                            Status = "Success"
                            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            Message = "Beskyttelse: $permDesc"
                        }
                        
                        # Update cache with new label (optimization: no rescan needed)
                        $script:fileLabelCache[$file] = @{
                            DisplayName = $selectedLabelObj.DisplayName
                            LabelId = $labelId
                            Rank = $newRank
                        }
                    }
                    catch {
                        Write-Log -Message "Could not apply protection to file" -Level 'ERROR' -Source 'ApplyButton-SyncMode' -Context @{ FilePath = $file; Protection = $permDesc } -Exception $_
                        $failureCount++
                        
                        # Update statistics
                        $stats.FailureCount++
                        $stats.TotalProcessed++
                        $stats.FailedFiles += @{
                            FilePath = $file
                            OriginalLabel = $originalLabel
                            NewLabel = $selectedLabelObj.DisplayName
                            ChangeType = $changeType
                            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            Error = $_.Exception.Message
                        }
                        
                        # DON'T throw - continue processing other files
                    }
                }
                else {
                    # Normal label application without custom protection
                    # Check if this specific file needs justification (downgrade case)
                    $fileNeedsJustification = ($changeType -eq "Downgrade")
                    
                    if ($fileNeedsJustification -and $userJustification) {
                        # Apply with justification
                        Set-AIPFileLabel -Path $file -LabelId $labelId -JustificationMessage $userJustification -PreserveFileDetails -ErrorAction Stop
                        $successCount++
                        Write-Log "SUKSESS (med begrunnelse): $file"
                    } else {
                        # Apply without justification
                        Set-AIPFileLabel -Path $file -LabelId $labelId -PreserveFileDetails -ErrorAction Stop
                        $successCount++
                        Write-Log "SUKSESS: $file"
                    }
                    
                    # Update statistics
                    $stats.SuccessCount++
                    $stats.TotalProcessed++
                    $stats.ChangeTypeBreakdown[$changeType]++
                    $stats.ProcessedFiles += @{
                        FilePath = $file
                        OriginalLabel = $originalLabel
                        NewLabel = $selectedLabelObj.DisplayName
                        ChangeType = $changeType
                        Status = "Success"
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Message = ""
                    }
                    
                    # Update cache with new label (optimization: no rescan needed)
                    $script:fileLabelCache[$file] = $selectedLabelObj.DisplayName
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                
                # Handle justification requirement
                if ($errorMessage -like "*Justification*") {
                    try {
                        Set-AIPFileLabel -Path $file -LabelId $labelId -JustificationMessage $userJustification -PreserveFileDetails -ErrorAction Stop
                        $successCount++
                        Write-Log "SUKSESS (med begrunnelse): $file"
                        
                        # Update statistics
                        $stats.SuccessCount++
                        $stats.TotalProcessed++
                        $stats.ChangeTypeBreakdown[$changeType]++
                        $stats.ProcessedFiles += @{
                            FilePath = $file
                            OriginalLabel = $originalLabel
                            NewLabel = $selectedLabelObj.DisplayName
                            ChangeType = $changeType
                            Status = "Success"
                            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            Message = "Med begrunnelse"
                        }
                        
                        # Update cache with new label (optimization: no rescan needed)
                        $script:fileLabelCache[$file] = @{
                            DisplayName = $selectedLabelObj.DisplayName
                            LabelId = $labelId
                            Rank = $newRank
                        }
                    }
                    catch {
                        throw
                    }
                }
                # Handle ad-hoc protection requirement (shouldn't happen if we use CustomPermissions correctly)
                elseif ($errorMessage -like "*ad-hoc protection*" -or $errorMessage -like "*AdhocProtectionRequired*") {
                    Write-Log -Message "Unexpected ad-hoc protection error despite custom permissions" -Level 'ERROR' -Source 'ApplyButton-SyncMode' -Context @{ FilePath = $file; ErrorMessage = $errorMessage }
                    $failureCount++
                    
                    # Update statistics
                    $stats.FailureCount++
                    $stats.TotalProcessed++
                    $stats.FailedFiles += @{
                        FilePath = $file
                        OriginalLabel = $originalLabel
                        NewLabel = $selectedLabelObj.DisplayName
                        ChangeType = $changeType
                        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        Error = $errorMessage
                    }
                }
                else {
                    throw
                }
            }
        }
        catch {
            $failureCount++
            $errorMsg = $_.Exception.Message
            Write-Log -Message "Label application failed for file" -Level 'ERROR' -Source 'ApplyButton-SyncMode' -Context @{ FilePath = $file; ErrorMessage = $errorMsg }
            
            # Update statistics
            $stats.FailureCount++
            $stats.TotalProcessed++
            $stats.FailedFiles += @{
                FilePath = $file
                OriginalLabel = $originalLabel
                NewLabel = $selectedLabelObj.DisplayName
                ChangeType = $changeType
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Error = $errorMsg
            }
        }
        }
        } catch {
            # === SYNC FOREACH CRASH HANDLER ===
            # Catches crashes that escape individual file error handling
            Write-Log -Message "CRITICAL: Sync foreach loop crashed" -Level 'CRITICAL' -Source 'ApplyButton-SyncMode' -Context @{
                CurrentFile = if (Test-Path variable:file) { $file } else { "Unknown" }
                ProcessedSoFar = $currentFile
                TotalFiles = $totalFiles
            } -Exception $_
            
            # Save crash details
            $crashInfo = @{
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
                Phase = "Synchronous Label Application"
                CurrentFile = if (Test-Path variable:file) { $file } else { "Unknown" }
                ProcessedFiles = $currentFile
                TotalFiles = $totalFiles
                Exception = $_.Exception.Message
                ExceptionType = $_.Exception.GetType().FullName
                StackTrace = $_.ScriptStackTrace
            }
            
            $crashLogPath = Join-Path $logDirectory "SYNC_FOREACH_CRASH_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            try {
                $crashInfo | ConvertTo-Json -Depth 10 | Set-Content $crashLogPath -Encoding UTF8
            } catch {
                # Fallback
                $crashLogPath = Join-Path $logDirectory "SYNC_FOREACH_CRASH_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
                $crashInfo | Out-String | Set-Content $crashLogPath
            }
            
            # Show error and open log
            $errorMsg = "Kritisk feil under synkron filbehandling!`n`n"
            $errorMsg += "Behandlet: $currentFile av $totalFiles filer`n"
            $errorMsg += "Feil: $($_.Exception.Message)`n`n"
            $errorMsg += "Detaljer: $crashLogPath"
            
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Sync Crash", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            
            if (Test-Path $crashLogPath) {
                Start-Process notepad.exe -ArgumentList $crashLogPath
            }
            
            # Don't re-throw - let outer handler deal with cleanup
        }
    }  # End of if (-not $useAsync) - synchronous mode
    
    # Stop stopwatch
    $stopwatch.Stop()
    $elapsedTime = $stopwatch.Elapsed
    
    # Complete
    $progressBar.Value = 100
    Write-Log -Message "==========================================" -Level 'INFO'
    Write-Log -Message "Label application session completed" -Level 'INFO' -Source 'LabelApplication'
    Write-Log -Message "Results: $successCount success, $failureCount failures" -Level 'INFO' -Source 'LabelApplication'
    Write-Log -Message "Time elapsed: $($elapsedTime.ToString())" -Level 'INFO' -Source 'LabelApplication'
    Write-Log -Message "Log file: $logFilePath" -Level 'INFO' -Source 'LabelApplication'
    Write-Log -Message "==========================================" -Level 'INFO'
    
    # Refresh display to show updated labels (cache already updated inline)
    if ($successCount -gt 0) {
        Update-FileListDisplay
    }
    
    # Show statistics dashboard
    $statusLabel.Text = if ($failureCount -eq 0) { "SUKSESS: Alle $successCount filer merket!" } else { "FULLFØRT: $successCount suksess, $failureCount feil" }
    $statusLabel.ForeColor = if ($failureCount -eq 0) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Orange }
    
    Show-StatisticsDialog -Statistics $stats -ElapsedTime $elapsedTime -LogFilePath $logFilePath -LabelName $selectedLabelObj.DisplayName
    
    # Re-enable buttons
    $applyBtn.Enabled = $true
    $browseBtn.Enabled = $true
    $clearBtn.Enabled = $true
    foreach ($btn in $labelButtons) {
        $btn.Enabled = $true
    }
        
    } catch {
        # === MASTER CATCH BLOCK ===
        # Captures ANY unhandled exception during label application
        
        Write-Log -Message "CRITICAL ERROR in Apply button handler" -Level 'CRITICAL' -Source 'ApplyButton' -Exception $_
        
        # Save detailed crash info
        $crashInfo = @{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
            Operation = "Label Application"
            SelectedLabel = if ($script:selectedLabelId) { $script:selectedLabelId } else { "None" }
            FileCount = $script:selectedFiles.Count
            Exception = $_.Exception.Message
            ExceptionType = $_.Exception.GetType().FullName
            StackTrace = $_.ScriptStackTrace
            InnerException = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { "None" }
        }
        
        $crashLogPath = Join-Path $logDirectory "APPLY_CRASH_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        try {
            $crashInfo | ConvertTo-Json -Depth 10 | Set-Content $crashLogPath -Encoding UTF8
        } catch {
            $crashLogPath = Join-Path $logDirectory "APPLY_CRASH_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $crashInfo | Out-String | Set-Content $crashLogPath
        }
        
        # Reset UI
        $progressBar.Value = 0
        $progressBar.Style = 'Continuous'
        $statusLabel.Text = "FEIL: Operasjon feilet - se detaljer"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        
        # Re-enable buttons
        $applyBtn.Enabled = $true
        $browseBtn.Enabled = $true
        $clearBtn.Enabled = $true
        foreach ($btn in $labelButtons) {
            $btn.Enabled = $true
        }
        
        # Show detailed error dialog
        $errorMsg = "En kritisk feil oppstod under etikettpaaforing:`n`n"
        $errorMsg += "Feil: $($_.Exception.Message)`n`n"
        $errorMsg += "Type: $($_.Exception.GetType().Name)`n`n"
        if ($_.Exception.InnerException) {
            $errorMsg += "Indre feil: $($_.Exception.InnerException.Message)`n`n"
        }
        $errorMsg += "Detaljert informasjon lagret i:`n$crashLogPath`n`n"
        $errorMsg += "Klikk OK for aa aapne feilrapporten."
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            $errorMsg,
            "Kritisk feil",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        # Open crash log
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            if (Test-Path $crashLogPath) {
                Start-Process notepad.exe -ArgumentList $crashLogPath
            }
        }
    }
})
$form.Controls.Add($applyBtn)

$viewLogBtn = New-Object System.Windows.Forms.Button
$viewLogBtn.Text = Get-LocalizedString -Key "buttons.viewLog"
$viewLogBtn.Location = New-Object System.Drawing.Point(($buttonCenterStartX + 200 + 10), $buttonInitialY)
$viewLogBtn.Size = New-Object System.Drawing.Size(100, 35)
$viewLogBtn.Add_Click({
    if (Test-Path $logFilePath) {
        Start-Process notepad.exe -ArgumentList $logFilePath
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "Ingen loggfil eksisterer ennå. Påfør etiketter for å generere en loggfil.",
            "Ingen loggfil",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})
$form.Controls.Add($viewLogBtn)

# Settings button
$settingsBtn = New-Object System.Windows.Forms.Button
$settingsBtn.Text = Get-LocalizedString -Key "buttons.settings"
$settingsBtn.Location = New-Object System.Drawing.Point(($buttonCenterStartX + 200 + 10 + 100 + 10), $buttonInitialY)
$settingsBtn.Size = New-Object System.Drawing.Size(100, 35)
    $settingsBtn.Add_Click({
    $configChanged = Show-SettingsDialog -Config $script:appConfig
    if ($configChanged) {
        $statusLabel.Text = Get-LocalizedString -Key "status.settingsUpdated"
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    }
})
$form.Controls.Add($settingsBtn)

# ========================================
# FOOTER INFO
# ========================================
$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Location = New-Object System.Drawing.Point(10, 500)  # Updated from 450
$footerLabel.Size = New-Object System.Drawing.Size(700, 20)
$footerLabel.Text = Get-LocalizedString -Key "form.footer"
$footerLabel.ForeColor = [System.Drawing.Color]::Gray
$footerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$footerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($footerLabel)

# ========================================
# SHOW FORM
# ========================================
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })

# Save configuration when form closes
$form.Add_FormClosing({
    param($sender, $e)
    
    try {
        # Cleanup runspace pool if active
        if ($script:runspacePool) {
            Write-Log -Message "Cleaning up runspace pool on exit" -Level 'INFO' -Source 'Shutdown'
            try {
                $script:runspacePool.Close()
                $script:runspacePool.Dispose()
                Write-Log -Message "Runspace pool disposed successfully" -Level 'INFO' -Source 'Shutdown'
            } catch {
                Write-Log -Message "Error disposing runspace pool" -Level 'WARNING' -Source 'Shutdown' -Exception $_
            }
        }
        
        # Update window position if configured to remember
        if ($appConfig.ui.rememberWindowPosition) {
            $appConfig.ui.windowPositionX = $form.Location.X
            $appConfig.ui.windowPositionY = $form.Location.Y
        }
        
        # Update window size
        $appConfig.ui.windowWidth = $form.Width
        $appConfig.ui.windowHeight = $form.Height
        
        # Save configuration
        $saveResult = Save-AppConfig -Config $appConfig
        if ($saveResult) {
            Write-Log -Message "Configuration saved successfully on exit" -Level 'INFO' -Source 'Shutdown'
        }
        else {
            Write-Log -Message "Failed to save configuration on exit" -Level 'WARNING' -Source 'Shutdown'
        }
        
        Write-Log -Message "==========================================" -Level 'INFO'
        Write-Log -Message "FileLabeler closed" -Level 'INFO' -Source 'Shutdown'
        Write-Log -Message "==========================================" -Level 'INFO'
    }
    catch {
        Write-Log -Message "Exception while closing application" -Level 'ERROR' -Source 'Shutdown' -Exception $_
    }
})

[void]$form.ShowDialog()

# Stop transcript logging
try {
    Stop-Transcript | Out-Null
} catch {
    # Transcript might not be running, ignore error
}
