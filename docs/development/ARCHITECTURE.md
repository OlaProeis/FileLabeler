# FileLabeler Architecture

**Technical overview and design decisions for developers**

---

## Table of Contents

- [Technology Stack](#technology-stack)
- [Project Structure](#project-structure)
- [Core Architecture](#core-architecture)
- [Key Components](#key-components)
- [Data Flow](#data-flow)
- [Performance Optimizations](#performance-optimizations)
- [Design Patterns](#design-patterns)
- [Critical Implementation Notes](#critical-implementation-notes)

---

## Technology Stack

| Layer | Technology | Version | Purpose |
|-------|------------|---------|---------|
| **Language** | PowerShell | 5.1+ | Core logic and UI |
| **GUI Framework** | Windows Forms | .NET 4.7.2+ | User interface |
| **Labeling** | PurviewInformationProtection | Latest | Microsoft Purview integration |
| **Async** | PowerShell Runspaces | Built-in | Non-blocking operations |
| **Testing** | Pester | 3.4.0+ | Unit and integration tests |
| **Platform** | Windows | 10/11 | Target OS |

---

## Project Structure

```
FileLabeler/
├── FileLabeler.ps1           # Main application (~5500 lines)
├── labels_config.json        # Label definitions (user-configured)
├── app_config.json           # Application settings (auto-generated)
│
├── docs/                     # Documentation
│   ├── USER_GUIDE.md        # User documentation (Norwegian)
│   ├── INSTALLATION.md      # Setup guide
│   ├── CONFIGURATION.md     # Configuration reference
│   ├── TROUBLESHOOTING.md   # Problem solving
│   ├── CHANGELOG.md         # Version history
│   ├── ROADMAP.md           # Future plans
│   └── development/         # Developer docs
│       ├── ARCHITECTURE.md  # This file
│       ├── TESTING.md       # Testing guide
│       ├── CONTRIBUTING.md  # Contribution guidelines
│       └── FEATURES.md      # Feature reference
│
├── tests/                    # Test suites
│   ├── FileLabeler.Tests.ps1              # Unit tests (58 tests)
│   ├── FileLabeler.Integration.Tests.ps1  # Integration tests (48 tests)
│   └── README.md                          # Test documentation
│
├── run_tests.ps1             # Quick test runner
├── run_integration_tests.ps1 # Integration test runner
│
└── archive/                  # Historical documentation (gitignored)
    ├── features/            # Feature implementation docs
    ├── bugfixes/           # Bug fix history
    └── uat/                # UAT materials
```

---

## Core Architecture

### Architectural Pattern

FileLabeler follows a **monolithic PowerShell script architecture** with clear functional sections:

```
┌─────────────────────────────────────────────┐
│           FileLabeler.ps1                   │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  Initialization & Setup              │  │
│  │  - Module loading                    │  │
│  │  - Configuration                     │  │
│  │  - Async pool creation               │  │
│  └──────────────────────────────────────┘  │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  Helper Functions                    │  │
│  │  - File operations                   │  │
│  │  - Label management                  │  │
│  │  - UI updates                        │  │
│  └──────────────────────────────────────┘  │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  Async Operations                    │  │
│  │  - Runspace pool                     │  │
│  │  - Folder scanning                   │  │
│  │  - Label retrieval                   │  │
│  │  - Batch application                 │  │
│  └──────────────────────────────────────┘  │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  UI Construction                     │  │
│  │  - Form creation                     │  │
│  │  - Control layout                    │  │
│  │  - Event handlers                    │  │
│  └──────────────────────────────────────┘  │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  Business Logic                      │  │
│  │  - Label application                 │  │
│  │  - Change analysis                   │  │
│  │  - Protection handling               │  │
│  └──────────────────────────────────────┘  │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  Application Loop                    │  │
│  │  - Event handling                    │  │
│  │  - UI updates                        │  │
│  └──────────────────────────────────────┘  │
└─────────────────────────────────────────────┘
```

### Why Monolithic?

**Advantages:**
- ✅ Single file deployment
- ✅ Easy to convert to EXE
- ✅ No module dependencies
- ✅ Simple debugging
- ✅ Clear execution flow

**Trade-offs:**
- ⚠️ Large file size (~5500 lines)
- ⚠️ Requires careful organization
- ⚠️ Testing requires function extraction

---

## Key Components

### 1. Label Cache (`$fileLabelCache`)

**Purpose:** In-memory cache of file label status

**Structure:**
```powershell
$fileLabelCache = @{
    "C:\path\to\file.docx" = @{
        Id = "guid-of-label"
        Rank = 2
        Name = "Confidential"
        Status = "Success" # or "Error", "Unknown"
    }
}
```

**Key Functions:**
- Stores label information for selected files
- Prevents redundant API calls
- Updated inline during label application (performance optimization)
- Thread-safe access via PowerShell runspace synchronization

**Location:** Lines 95-130

---

### 2. Dynamic UI Layout (`Adjust-UILayout`)

**Purpose:** Automatically resize UI based on file count

**Logic:**
```powershell
Function Adjust-UILayout {
    # Calculate optimal rows (3-10)
    $rows = [Math]::Min([Math]::Max($selectedFiles.Count, 3), 10)
    
    # Calculate listbox height (16px per row)
    $listBoxHeight = $rows * 16
    
    # Reposition all dependent controls
    # - File count label
    # - Folder selection
    # - Label buttons
    # - Apply button
    # - Settings button
    # - Log button
}
```

**Features:**
- Minimum: 3 rows (48px)
- Maximum: 10 rows (160px)
- 16px per item (optimized for Segoe UI 9pt)
- Form height adjusts dynamically
- Smooth transitions

**Location:** Lines 197-242

---

### 3. Async Operations (Runspaces)

**Architecture:**

```
┌─────────────────────────────────────────┐
│  UI Thread (Main Application)           │
│  - User interaction                     │
│  - UI updates                           │
│  - Event handling                       │
└──────────────┬──────────────────────────┘
               │
               │ Spawn
               ↓
┌─────────────────────────────────────────┐
│  Runspace Pool (1-4 threads)            │
│                                         │
│  ┌─────────────┐  ┌─────────────┐      │
│  │ Worker 1    │  │ Worker 2    │      │
│  │ - Scan      │  │ - Labels    │      │
│  │   folder    │  │   retrieval │      │
│  └─────────────┘  └─────────────┘      │
│                                         │
│  ┌─────────────┐  ┌─────────────┐      │
│  │ Worker 3    │  │ Worker 4    │      │
│  │ - Label     │  │ - Label     │      │
│  │   apply     │  │   apply     │      │
│  └─────────────┘  └─────────────┘      │
└──────────────┬──────────────────────────┘
               │
               │ Results
               ↓
┌─────────────────────────────────────────┐
│  Synchronized Collections                │
│  - Thread-safe data exchange             │
│  - Monitor locking                       │
└──────────────┬──────────────────────────┘
               │
               │ UI Updates
               ↓
┌─────────────────────────────────────────┐
│  $form.Invoke() - Thread-safe UI        │
│  - Progress updates                     │
│  - Status messages                      │
│  - Result display                       │
└─────────────────────────────────────────┘
```

**Key Functions:**

#### New-FileLabelerRunspacePool (Lines 133-161)
Creates runspace pool with 1-4 threads (based on CPU cores)

#### Start-AsyncFolderScan (Lines 163-253)
Asynchronous folder scanning with recursive support

#### Start-AsyncLabelRetrieval (Lines 255-364)
Parallel label retrieval for >50 files (4x faster!)

#### Start-AsyncBatchLabelApplication (Lines 458-666)
Batch label application for ≥30 files

#### Update-UIThreadSafe (Lines 366-385)
Thread-safe UI updates from background threads

#### Wait-AsyncJobsWithUI (Lines 387-456)
Responsive waiting with progress updates

**Smart Thresholds:**
- Folder scan: Always async (any count)
- Label retrieval: >50 files
- Label apply: ≥30 files
- Thread limiting: 1-4 concurrent (max 8 based on CPU)

---

### 4. Change Analysis and Warnings

**Purpose:** Analyze label changes before application

**Process:**

```
Selected Files
     ↓
Analyze-LabelChanges (Lines 781-860)
     ↓
Categorize into:
  - New: File had no label
  - Upgrade: Higher rank
  - Downgrade: Lower rank
  - Same: Same label (different properties)
  - Unchanged: Identical label
     ↓
Get-ChangeWarnings (Lines 862-917)
     ↓
Detect warnings:
  - Mass downgrade (>50% downgrade)
  - Large batch (>50 files)
  - No changes (all unchanged)
  - Protection required
  - Mixed changes (upgrades + downgrades)
     ↓
Show-PreApplySummary (Lines 919-1049)
     ↓
User decision: Proceed or Cancel
```

**Location:** Lines 781-1049

---

### 5. Label Application Logic

**Flow:**

```
1. Get selected files
2. Get selected label
3. Analyze changes (pre-apply summary)
4. For each file:
   a. Check if downgrade → prompt justification
   b. Check if protection required → show protection dialog
   c. Apply label with Set-AIPFileLabel
   d. Update cache inline
   e. Track statistics
5. Show statistics dashboard
6. Offer CSV export
```

**Key Cmdlets:**
- `Get-AIPFileStatus` - Retrieve current label
- `Set-AIPFileLabel` - Apply label with `-PreserveFileDetails`
- `New-AIPCustomPermissions` - Create protection settings

**Location:** Apply button click handler (Lines 1080-1300+)

---

## Data Flow

### File Selection Flow

```
User Action
  ├─ File Browser → Select Files
  ├─ Folder Browser → Scan Folder
  └─ Drag & Drop → Process Dropped Items
          ↓
    Merge & Deduplicate
          ↓
    Retrieve Labels (async if >50)
          ↓
    Update Cache
          ↓
    Update UI Display
          ↓
    Adjust Layout
```

### Label Application Flow

```
User Clicks "Påfør Etikett"
          ↓
    Validate Selection
          ↓
    Analyze Changes
          ↓
    Show Summary Dialog
          ↓
    User Confirms
          ↓
  ┌─────────────────────┐
  │ For Each File:      │
  │   1. Get Status     │
  │   2. Check Rules    │
  │   3. Apply Label    │
  │   4. Update Cache   │
  │   5. Log Result     │
  └─────────────────────┘
          ↓
    Show Statistics
          ↓
    Offer CSV Export
```

---

## Performance Optimizations

### 1. Inline Cache Updates

**Before:**
```powershell
# Apply labels
foreach ($file in $files) {
    Set-AIPFileLabel -Path $file ...
}

# Refresh ALL labels (slow!)
Update-FileListDisplay
```

**After:**
```powershell
# Apply labels AND update cache inline
foreach ($file in $files) {
    Set-AIPFileLabel -Path $file ...
    
    # Update cache immediately (60-100x faster!)
    $fileLabelCache[$file] = @{
        Id = $newLabelId
        Rank = $newRank
        Name = $newName
    }
}
```

**Impact:** 60-100x faster label refresh

### 2. Enhanced Label Cache

**Before:**
```powershell
$fileLabelCache = @{
    "file.docx" = "Label Name"  # String only
}
```

**After:**
```powershell
$fileLabelCache = @{
    "file.docx" = @{
        Id = "guid"
        Rank = 2
        Name = "Confidential"
    }
}
```

**Impact:** 
- 50% reduction in API calls
- Instant analysis phase
- Reuse during change detection

### 3. Async Operations

**Impact:**
- UI remains responsive with 500+ files
- No "Not Responding" dialogs
- Maintains <100ms UI update cycles
- 4x faster with parallelism (label retrieval)

**Thresholds tuned based on testing:**
- Folder scan: Always async (prevents UI freeze)
- Label retrieval: >50 files (API call overhead vs parallelism benefit)
- Batch apply: ≥30 files (sufficient to show progress benefit)

---

## Design Patterns

### 1. Event-Driven Architecture

All UI interactions trigger event handlers:

```powershell
$button.Add_Click({
    # Event handler logic
    # Access form controls via script scope ($script:controlName)
})
```

### 2. Async-Await Pattern (PowerShell Variant)

```powershell
# Start async operation
$job = Start-AsyncOperation

# Wait with UI updates
Wait-AsyncJobsWithUI -Jobs @($job)

# Get results
$results = $job.Results
```

### 3. Factory Pattern

Runspace pool creation:

```powershell
Function New-FileLabelerRunspacePool {
    # Create and configure pool
    # Return ready-to-use pool
}
```

### 4. Strategy Pattern

Different file processing strategies based on count:

```powershell
if ($fileCount -ge $threshold) {
    # Async strategy
    Start-AsyncLabelRetrieval
} else {
    # Sync strategy
    foreach ($file in $files) { ... }
}
```

---

## Critical Implementation Notes

### ⚠️ Norwegian Character Encoding

**MUST USE UTF-8 with BOM** for `FileLabeler.ps1`

```powershell
# In VS Code/Cursor:
# Bottom right → "Save with Encoding" → "UTF-8 with BOM"
```

**Without BOM:** æ/ø/å display as Ã¥/Ã¸/Ã¦

### ⚠️ Windows Forms Initialization

These lines MUST be at the top:

```powershell
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
```

### ⚠️ Label Button Implementation

Buttons are **custom panels with labels inside**, NOT standard buttons:

```powershell
$btnPanel = New-Object System.Windows.Forms.Panel
$btnLabel = New-Object System.Windows.Forms.Label
# Store colors in $btnPanel.Tag for toggle logic
$btnPanel.Tag = @{
    SelectedColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    DefaultColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
}
```

**Why:** Better control over appearance and click behavior

### ⚠️ Thread-Safe UI Updates

Never update UI directly from background thread:

```powershell
# ❌ WRONG - Will crash
$label.Text = "Updated from background thread"

# ✅ CORRECT - Use Invoke
$form.Invoke([Action]{
    $label.Text = "Updated safely"
})
```

### ⚠️ Null-Safe Label Handling

Always check for null before accessing properties:

```powershell
if ($status -and $status.LabelId) {
    # Safe to use $status.LabelId
}
```

**Why:** Protected/encrypted labels may have null properties

---

## Code Organization

### File Structure (Logical Sections)

```powershell
# Lines 1-90: Module loading, validation, initialization
# Lines 95-130: Global variables and label cache
# Lines 133-670: Async operation functions
# Lines 680-780: Helper functions (logging, file ops)
# Lines 781-1050: Analysis and warning functions
# Lines 1060-2500: UI construction
# Lines 2500-4500: Event handlers and business logic
# Lines 4500-5500: Application loop and cleanup
```

### Naming Conventions

- **Functions:** PascalCase with verb-noun (e.g., `Get-SupportedFiles`, `Update-UIThreadSafe`)
- **Variables:** camelCase (e.g., `$fileListBox`, `$selectedFiles`)
- **Global/Script:** `$script:` or `$global:` prefix when needed
- **Constants:** UPPER_CASE (e.g., `$SUPPORTED_EXTENSIONS`)

### Comments

- Norwegian for user-facing text
- English for technical/code comments
- Section headers with `# ===== SECTION NAME =====`

---

## Testing Strategy

See [TESTING.md](TESTING.md) for comprehensive testing guide.

**Quick Overview:**
- **Unit Tests:** 58 tests, function-level testing
- **Integration Tests:** 48 tests, end-to-end workflows
- **Manual Tests:** Legacy test scripts for specific features
- **Coverage:** 100% pass rate

---

## Security Considerations

### Data Privacy
- No data sent to external services (except Microsoft Purview API)
- Logs stored locally (`Documents\FileLabeler_Logs\`)
- No passwords or sensitive data stored
- Runs with user's permissions (no privilege escalation)

### Label Access Control
- Application inherits user's Purview permissions
- Cannot apply labels user doesn't have access to
- Respects organization's Purview policies

### Code Signing
- Recommended for EXE distribution
- Use valid code signing certificate
- Prevents "Unknown Publisher" warnings

---

## Known Limitations

1. **Windows-only:** Requires Windows PowerShell and Purview client
2. **Supported file types:** Office formats and PDF only
3. **No label removal:** Designed for labeling, not removing labels
4. **Single label per batch:** All files get same label (by design)
5. **No offline mode:** Requires connection to Microsoft Purview

---

## Future Architecture Considerations

For v2.0 and beyond:

### Potential Improvements
- **Module-based architecture:** Split into separate .psm1 modules
- **Configuration provider pattern:** Abstract configuration access
- **Plugin system:** Allow custom label providers
- **Event sourcing:** Track all label changes for audit

### Challenges
- Maintaining single-file deployment
- Backwards compatibility
- EXE compilation complexity
- Performance trade-offs

---

## Related Documentation

- **[Testing Guide](TESTING.md)** - Testing architecture and suites
- **[Contributing](CONTRIBUTING.md)** - Development workflow
- **[Features Reference](FEATURES.md)** - Detailed feature documentation

---

**For questions about architecture, open a GitHub Discussion or contact the development team.**

**Last Updated:** October 2025  
**Version:** 1.1

