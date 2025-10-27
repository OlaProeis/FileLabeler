# FileLabeler User Guide

**Version:** 1.1  
**Last Updated:** October 2025

---

## Table of Contents

- [Getting Started](#getting-started)
- [Basic Workflow](#basic-workflow)
- [Features](#features)
- [Advanced Usage](#advanced-usage)
- [Tips and Best Practices](#tips-and-best-practices)
- [Common Scenarios](#common-scenarios)

---

## Getting Started

### What is FileLabeler?

FileLabeler is a Windows application that allows you to apply Microsoft Purview sensitivity labels to multiple files simultaneously. It preserves original file dates and provides an intuitive interface for bulk labeling operations.

### Prerequisites

Before using FileLabeler, ensure you have:
- ✅ FileLabeler installed (see [Installation Guide](INSTALLATION.md))
- ✅ Labels configured in `labels_config.json` (see [Configuration Guide](CONFIGURATION.md))
- ✅ Microsoft Purview Information Protection client installed
- ✅ Permission to apply sensitivity labels to files

---

## Basic Workflow

### Step 1: Launch the Application

**Method 1: Double-click the executable**
```
FileLabeler.exe
```

**Method 2: Run the PowerShell script**
```powershell
.\FileLabeler.ps1
```

The application will:
1. Check for required modules
2. Load your label configuration
3. Initialize the user interface

---

### Step 2: Select Files

You have three methods to select files:

#### Method A: File Browser

1. Click **"Select files..."** button
2. Navigate to your files
3. Select one or multiple files (Ctrl+Click or Shift+Click)
4. Click **Open**

**Supported file types:**
- Word documents: `.docx`, `.doc`
- Excel workbooks: `.xlsx`, `.xls`
- PowerPoint presentations: `.pptx`, `.ppt`
- PDF files: `.pdf`

#### Method B: Folder Browser

1. Click **"Select folder..."** button
2. Navigate to the desired folder
3. Check **"Include subfolders"** if you want recursive scanning
4. Click **Select Folder**

All supported files in the folder (and subfolders if checked) will be added to the list.

#### Method C: Drag and Drop

1. Open Windows Explorer
2. Select files and/or folders
3. Drag them to the FileLabeler window
4. Drop them on either the main window or the file list

**Visual feedback:** The application shows a light blue tint during the drag operation.

---

### Step 3: Review Selected Files

The file list shows:
- **File name** with path
- **Current label** in brackets (e.g., `[Confidential]`)
- **File count** below the list (e.g., "15 files selected")

**Possible label statuses:**
- `[Confidential]` - File has this label
- `[No label]` - File has no sensitivity label
- `[Unknown label (protected)]` - File has encrypted label not in your configuration
- `[Error retrieving]` - Could not read file label

---

### Step 4: Select Sensitivity Label

1. Click one of the label buttons
2. Selected label highlights in blue
3. Only one label can be selected at a time

The labels shown come from your `labels_config.json` configuration.

---

### Step 5: Apply Labels

1. Click **"Apply label (preserve dates)"** button

2. **Smart Preview Dialog appears:**
   - Shows categorized changes:
     - **New**: File had no label
     - **Upgrade**: Moving to higher sensitivity
     - **Downgrade**: Moving to lower sensitivity (requires justification)
     - **Same**: Same label, different properties
     - **Unchanged**: Identical label
   - Shows warnings if applicable:
     - 🔴 Mass downgrade (>50% of files)
     - 🟠 Large batch (>50 files)
     - 🟡 No changes detected
     - 🔵 Protection required
     - 🟣 Mixed changes
   - Option: **"Skip unchanged files"** for better performance
   - Click **"Continue"** to proceed or **"Cancel"** to stop

3. **If downgrade detected:**
   - Justification dialog appears
   - Enter reason for downgrade
   - Default: "Changed via bulk labeling"
   - Click **OK** to continue or **Cancel** to stop

4. **If label requires protection:**
   - Protection dialog appears with options:
     - **Viewer - View only**: Read-only access
     - **Reviewer - View, edit**: Can view and comment
     - **Co-Author - View, edit, copy and print**: Full editing
     - **Co-Owner - All permissions**: Full control including changing permissions
     - **Owner only** (default): Only you
   - **Specify users** (if not "Owner only"):
     - Enter email addresses, comma-separated
     - Example: `user1@domain.com, user2@domain.com`
   - **Expiration date** (optional):
     - Check "This document expires:"
     - Select date
   - Click **OK** to apply or **Cancel** to skip

5. **Progress bar shows:**
   - Current operation
   - Percentage complete
   - Time estimate
   - Async operations for ≥30 files (faster, responsive UI)

6. **Results message displays:**
   - Number of successful operations
   - Number of failed operations

---

### Step 6: Review Results

**Statistics Dashboard automatically appears:**
- Detailed breakdown by change type
- Time elapsed (precise timing)
- Success/failure counts with color coding
- Clickable log file link
- **"Export to CSV"** button

**Actions:**

1. **Export to CSV:**
   - Click **"Export to CSV"**
   - Choose save location
   - CSV includes:
     - Summary section
     - Per-file details: FilePath, OriginalLabel, NewLabel, ChangeType, Status, Timestamp, Message
   - UTF-8 encoding for international characters

2. **View Log:**
   - Click **"View log"**
   - Opens latest log file in Notepad
   - Log location: `Documents\FileLabeler_Logs\`
   - File format: `FileLabeler_Log_yyyyMMdd_HHmmss.txt`

3. **Clear Selection (optional):**
   - Click **"Clear selection"** to remove all files
   - Start fresh with new file selection

---

## Features

### Enhanced File List (v1.1)

**Current label display:**
- Shows current label next to each file
- Auto-refreshes after applying labels
- 60-100x faster than rescanning (inline cache updates)

**Label status indicators:**
- Green text: Successfully labeled
- Red text: Error occurred
- Gray text: Skipped (unchanged)

---

### Dynamic UI Sizing (v1.1)

**Automatic resizing:**
- File list adjusts between 3-10 rows based on file count
- Form height adapts automatically
- Smooth transitions when files are added/removed

**Benefits:**
- Optimal screen space usage
- No scrolling for small lists
- Compact view for large lists

---

### Folder Import (v1.1)

**Recursive scanning:**
- Check **"Include subfolders"** to scan entire folder tree
- Uncheck to scan only the selected folder
- Setting is remembered for drag-and-drop operations

**Performance:**
- Async scanning for large folders (responsive UI)
- Duplicate prevention
- Progress shown in real-time

---

### Drag and Drop (v1.1)

**Supports:**
- Individual files
- Multiple files
- Folders
- Mixed files and folders

**Visual feedback:**
- Light blue tint during drag
- Resets on drag leave
- Indicates drop is accepted

**Respects:**
- "Include subfolders" checkbox setting
- File type filtering
- Duplicate prevention

---

### Smart Pre-Apply Summary (v1.1)

**Change categorization:**
- **New** (green): Files getting their first label
- **Upgrade** (blue): Moving to higher sensitivity
- **Downgrade** (orange): Moving to lower sensitivity (requires justification)
- **Same** (yellow): Same label, different properties
- **Unchanged** (gray): Identical label (can be skipped)

**Warning system:**
- Mass downgrade detection (>50%)
- Large batch warning (>50 files)
- No changes detected
- Protection requirements
- Mixed changes (upgrades + downgrades)

**Optimization:**
- "Skip unchanged files" option improves performance
- Reduces unnecessary API calls

---

### Statistics Dashboard (v1.1)

**Comprehensive results:**
- Detailed breakdown by change type
- Precise timing with Stopwatch
- Success/failure counts
- Color-coded statistics
- Clickable log file link

**CSV export:**
- Complete operation log
- All file details
- UTF-8 encoding
- Ready for analysis in Excel

---

### Async Operations (v1.1)

**Responsive UI:**
- No freezing with large file sets (500+ files tested)
- Progress updates in real-time
- Form remains movable and interactive
- Maintains <100ms UI update cycles

**Smart thresholds:**
- Folder scan: Always async (any count)
- Label retrieval: >50 files
- Batch apply: ≥30 files
- Automatically selects sync/async based on file count

**Performance:**
- 4x faster with parallelism (label retrieval)
- Uses 1-4 background threads
- Thread-safe operations

---

## Advanced Usage

### Working with Protected Labels

**Labels with `RequiresProtection: true`:**

1. Select files and protected label
2. Click "Apply label"
3. Protection dialog appears
4. Choose permission level
5. Add users (or select "Owner only")
6. Set expiration date (optional)
7. Click OK

**Permission levels explained:**
- **Viewer**: Users can only view, cannot edit or print
- **Reviewer**: Users can view and add comments
- **Co-Author**: Users can view, edit, copy, and print
- **Co-Owner**: Users have all permissions including changing access
- **Owner only**: No sharing, only you have access

---

### Handling Downgrades

**Automatic detection:**
- FileLabeler detects when you're applying a lower-ranked label
- Rank is defined in `labels_config.json` (0 = lowest, higher numbers = higher sensitivity)

**Justification process:**
1. Select files with higher-ranked labels
2. Select lower-ranked label
3. Click "Apply label"
4. Summary shows downgrades
5. Click "Continue"
6. Justification dialog appears
7. Enter reason (required by Microsoft Purview policy)
8. Click OK

**Default justification:**
```
Changed via bulk labeling
```

You can customize this or enter a specific reason.

---

### Processing Large Batches

**For 100+ files:**

1. **Check system resources:**
   - Close unnecessary applications
   - Ensure sufficient RAM (8GB+ recommended)
   - Stable network connection if using network files

2. **Use async operations:**
   - Automatically enabled for ≥30 files
   - Configurable in Settings if needed

3. **Monitor progress:**
   - Progress bar shows percentage
   - Time estimate updates in real-time
   - Status messages show current operation

4. **Review results:**
   - Statistics dashboard shows detailed breakdown
   - Export to CSV for record-keeping
   - Check log file for any errors

**Best practices:**
- Process files in batches of 200-300 for optimal performance
- Local files are faster than network files
- Wait for OneDrive sync to complete before labeling

---

### Working with OneDrive Files

**Verified compatible:**
- Tested with 100+ OneDrive files
- All operations successful
- Async operations handle cloud latency gracefully

**Tips:**
1. **Wait for sync:**
   - Check OneDrive icon in system tray
   - Ensure files are fully synced (not cloud-only)
   - Green checkmark indicates ready

2. **Performance:**
   - OneDrive sync can add 2-3 seconds per file
   - Consider copying large batches locally first
   - Use async operations (automatic for ≥30 files)

3. **Troubleshooting:**
   - If files show as locked, wait for sync to complete
   - Check OneDrive status in settings
   - Verify files aren't open in Office Online

---

### Using Configuration Settings

**Access settings:**
1. Click **"Settings"** button
2. Navigate through tabs:
   - **Warnings**: Configure warning preferences
   - **Performance**: Async operation thresholds
   - **Logging**: Log verbosity and retention

**Common adjustments:**

**Disable specific warnings:**
- Uncheck warnings you don't need
- Speeds up workflow for experienced users

**Adjust async thresholds:**
- Lower thresholds = faster async activation
- Higher thresholds = less overhead for small batches
- Default values work well for most scenarios

**Enable detailed logging:**
- For troubleshooting
- Creates larger log files
- Disable when not needed

---

## Tips and Best Practices

### Performance Optimization

✅ **Do:**
- Process files in batches of 200-300
- Close files before labeling
- Use local copies for large network batches
- Wait for OneDrive sync to complete
- Keep async operations enabled

❌ **Don't:**
- Try to process 1000+ files in one batch
- Label files open in Office applications
- Process cloud-only OneDrive files
- Disable async operations unnecessarily

---

### File Selection

✅ **Do:**
- Use drag-and-drop for quick selection
- Use folder import for entire directories
- Check file count before applying
- Review current labels in the list

❌ **Don't:**
- Select mixed file types if unsure about labeling
- Forget to check "Include subfolders" setting
- Process files you don't have permission to modify

---

### Label Application

✅ **Do:**
- Review smart preview before applying
- Read warnings carefully
- Provide meaningful justification for downgrades
- Export results to CSV for records
- Check log files after large operations

❌ **Don't:**
- Skip preview without reading
- Ignore warnings (they're there for a reason)
- Use generic justifications for compliance-critical downgrades
- Forget to verify results after application

---

### Troubleshooting

✅ **Do:**
- Check log files first (`Documents\FileLabeler_Logs\`)
- Verify files aren't open in other applications
- Ensure you have proper permissions
- Test with a few files before large batches
- Read error messages carefully

❌ **Don't:**
- Ignore error messages
- Retry immediately without investigating
- Process files without checking logs
- Continue with large batches after failures

---

## Common Scenarios

### Scenario 1: Labeling a Project Folder

**Goal:** Apply "Confidential" to all Word documents in a project folder

**Steps:**
1. Click "Select folder..."
2. Navigate to project folder
3. Check "Include subfolders"
4. Click "Select Folder"
5. Files load with current labels shown
6. Click "Confidential" label button
7. Click "Apply label"
8. Review smart preview (shows all changes)
9. Click "Continue"
10. Review statistics dashboard
11. Export to CSV for project records

**Expected result:** All Word/Excel/PowerPoint/PDF files in folder tree labeled "Confidential"

---

### Scenario 2: Correcting Mislabeled Files

**Goal:** Change files from "Internal" to "Public"

**Steps:**
1. Select mislabeled files (any method)
2. Current label shows "[Internal]"
3. Click "Public" label button
4. Click "Apply label"
5. Smart preview shows "Downgrade" (if Public is lower rank)
6. Click "Continue"
7. Justification dialog appears
8. Enter: "Correcting classification error"
9. Click OK
10. Labels applied
11. Check log for verification

**Expected result:** Files successfully downgraded to "Public" with justification logged

---

### Scenario 3: Labeling Incoming Files

**Goal:** Quickly label files received from external source

**Steps:**
1. Open folder with new files
2. Drag files directly to FileLabeler window
3. Review current labels (likely "No label")
4. Select appropriate label (e.g., "Internal")
5. Click "Apply label"
6. Smart preview shows all as "New"
7. Click "Continue"
8. Labels applied immediately
9. Clear selection for next batch

**Expected result:** Fast workflow for daily file labeling tasks

---

### Scenario 4: Sharing Protected Documents

**Goal:** Label and protect sensitive documents for specific team members

**Steps:**
1. Select confidential files
2. Click "Highly Confidential" (protected label)
3. Click "Apply label"
4. Smart preview shows changes
5. Click "Continue"
6. Protection dialog appears
7. Select "Co-Author" permission level
8. Enter team members' emails: `user1@company.com, user2@company.com`
9. Optionally set expiration date
10. Click OK
11. Labels and protection applied
12. Export CSV for access records

**Expected result:** Files labeled and accessible only to specified users with co-author rights

---

### Scenario 5: Audit and Reporting

**Goal:** Generate compliance report of labeled files

**Steps:**
1. Select folder to audit
2. Check "Include subfolders"
3. Files load with current labels
4. Select label to verify/apply
5. Click "Apply label"
6. Smart preview shows:
   - How many already have correct label (Unchanged)
   - How many need labeling (New)
   - Any misclassified (Upgrade/Downgrade)
7. Check "Skip unchanged files"
8. Click "Continue"
9. Review statistics
10. Export to CSV
11. Open CSV in Excel for analysis

**Expected result:** Detailed compliance report showing all label statuses and changes

---

## Keyboard Shortcuts

Currently, FileLabeler doesn't have keyboard shortcuts, but you can use standard Windows shortcuts:

- **Ctrl+C**: Copy selected file paths from list (if supported)
- **Alt+F4**: Close application
- **Windows+D**: Show desktop (minimize all)

---

## Logging and Auditing

### Log Files

**Location:**
```
C:\Users\<username>\Documents\FileLabeler_Logs\
FileLabeler_Log_yyyyMMdd_HHmmss.txt
```

**Contents:**
- Session start/end timestamps
- Files processed
- Labels applied
- Success/failure status
- Error messages with details
- Timing information

**Log levels:**
- `[INFO]`: Normal operations
- `[WARNING]`: Non-critical issues
- `[ERROR]`: Failed operations
- `[CRITICAL]`: Application errors

**Retention:**
- Default: 30 days (configurable in Settings)
- Automatic cleanup of old logs
- Manual cleanup: Delete files in log folder

---

### CSV Export

**Format:**
```csv
Summary
Total Files,15
Successful,14
Failed,1
Time Elapsed,00:00:23

File Details
FilePath,OriginalLabel,NewLabel,ChangeType,Status,Timestamp,Message
C:\Documents\file1.docx,Internal,Confidential,Upgrade,Success,2025-10-27 14:30:15,
C:\Documents\file2.xlsx,No label,Internal,New,Success,2025-10-27 14:30:16,
C:\Documents\file3.pdf,Confidential,Confidential,Unchanged,Skipped,2025-10-27 14:30:16,Skipped unchanged file
```

**Uses:**
- Compliance auditing
- Troubleshooting
- Management reporting
- Change tracking
- Verification

---

## Getting Help

### In-App Help

- Status messages provide guidance
- Error dialogs include suggestions
- Warnings explain potential issues
- Log files contain detailed information

### Documentation

- **Installation**: [INSTALLATION.md](INSTALLATION.md)
- **Configuration**: [CONFIGURATION.md](CONFIGURATION.md)
- **Troubleshooting**: [TROUBLESHOOTING.md](TROUBLESHOOTING.md)
- **Developer Docs**: [development/](development/)

### Support

1. Check [Troubleshooting Guide](TROUBLESHOOTING.md)
2. Review log files
3. Search [GitHub Issues](https://github.com/yourusername/FileLabeler/issues)
4. Create new issue with details

---

**Happy Labeling!** 🏷️

For technical details and architecture, see [ARCHITECTURE.md](development/ARCHITECTURE.md).
