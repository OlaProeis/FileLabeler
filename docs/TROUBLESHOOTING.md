# Troubleshooting Guide

This guide covers common issues and their solutions for FileLabeler.

---

## Quick Diagnostic Steps

Before diving into specific issues:

1. **Check the log file**: `Documents\FileLabeler_Logs\FileLabeler_Log_[timestamp].txt`
2. **Verify Purview module**: `Get-Module -ListAvailable -Name PurviewInformationProtection`
3. **Test with a single file** before processing large batches
4. **Check file permissions** and ensure files aren't open in other applications

---

## Common Issues

### 🔴 Application Won't Start

**Symptoms:**
- Application window doesn't appear
- Immediate crash on startup
- Error about missing modules

**Solutions:**

1. **Verify PurviewInformationProtection module is installed:**
   ```powershell
   Get-Module -ListAvailable -Name PurviewInformationProtection
   ```
   If not found, install [Microsoft Purview Information Protection Client](https://www.microsoft.com/en-us/download/details.aspx?id=53018)

2. **Check PowerShell version:**
   ```powershell
   $PSVersionTable.PSVersion  # Should be 5.1 or higher
   ```

3. **Review startup log:**
   - Open most recent log file
   - Look for `[CRITICAL]` entries
   - Common causes: Missing module, .NET Framework issues

4. **Try running as Administrator:**
   ```powershell
   # Right-click PowerShell → Run as Administrator
   .\FileLabeler.ps1
   ```

---

### 🔴 Files Show "Ukjent" (Unknown) Label

**Symptoms:**
- Files display "Ukjent" instead of label name
- Or "Ukjent etikett (beskyttet)"
- Or "Feil ved henting"

**Causes and Solutions:**

#### Cause 1: File Never Had a Label
**Expected behavior** - Files without labels show "Ingen etikett"

**Solution:** No action needed, this is correct

#### Cause 2: Label Not in Configuration
**File has encrypted/protected label not in `labels_config.json`**

**Log entry:**
```
[WARNING] [Get-FileLabelDisplayName] Label ID not found in configuration
```

**Solution:**
1. Check log file for the label ID
2. Add missing label to `labels_config.json`:
   ```json
   {
     "DisplayName": "Label Name",
     "Id": "guid-from-log",
     "Rank": 3
   }
   ```

#### Cause 3: Permission Issues
**User doesn't have permission to read file metadata**

**Solution:**
1. Check file permissions (right-click → Properties → Security)
2. Ensure user has Read permissions
3. Try accessing file in Word/Excel to verify

#### Cause 4: Network Latency
**OneDrive or network share causing delays**

**Solution:**
- Wait for sync to complete (check OneDrive icon in system tray)
- Ensure stable network connection
- Try again after sync completes

---

### 🔴 "Ingen tilgang til filen" (No File Access)

**Error Category:** `FileAccess`  
**Log Level:** `WARNING` or `ERROR`

**Solutions:**

1. **Check file permissions:**
   - Right-click file → Properties → Security
   - Ensure your user has Read/Write permissions

2. **Check if file is read-only:**
   - Right-click file → Properties
   - Uncheck "Read-only" attribute

3. **Check folder permissions:**
   - If entire folder fails, check folder permissions
   - Verify you have Modify access

4. **OneDrive/SharePoint issues:**
   - Check sync status
   - Verify file is fully downloaded (not cloud-only)
   - Wait for sync to complete

---

### 🔴 "Filen er i bruk" (File In Use)

**Error Category:** `FileLocked`  
**Log Level:** `ERROR`

**Solutions:**

1. **Close the file:**
   - Close file in Word, Excel, PowerPoint
   - Close file in PDF readers
   - Close any applications with the file open

2. **Check for background processes:**
   ```powershell
   # Find processes using the file
   Get-Process | Where-Object {$_.MainWindowTitle -like "*filename*"}
   ```

3. **Use Task Manager:**
   - Open Task Manager (Ctrl+Shift+Esc)
   - Look for Office applications
   - End stubborn processes

4. **Wait and retry:**
   - Sometimes Windows holds file locks briefly
   - Wait 10-30 seconds and retry

5. **Reboot if persistent:**
   - Last resort for stubborn file locks

---

### 🔴 Network Folder Errors

**Error Category:** `Network`  
**Log Level:** `ERROR`

**Solutions:**

1. **Test network connectivity:**
   ```powershell
   Test-Connection -ComputerName servername -Count 2
   ```

2. **Verify UNC path access:**
   ```powershell
   Test-Path "\\server\share\folder"
   ```

3. **Check VPN connection:**
   - Ensure VPN is connected if required
   - Verify network drive mappings

4. **Verify share permissions:**
   - Contact IT if access denied
   - Check AD group memberships

5. **Map network drive:**
   ```powershell
   New-PSDrive -Name "Z" -PSProvider FileSystem -Root "\\server\share" -Persist
   ```

---

### 🔴 Configuration Errors

**Error Category:** `Config`  
**Log Level:** `ERROR` or `WARNING`

**Solutions:**

1. **Verify `labels_config.json` exists:**
   ```powershell
   Test-Path .\labels_config.json
   ```

2. **Validate JSON syntax:**
   - Use [jsonlint.com](https://jsonlint.com)
   - Check for missing commas, quotes, brackets

3. **Check for backup files:**
   - `app_config.json.invalid_*` (validation failed)
   - `app_config.json.backup_*` (corrupted)

4. **Reset to defaults:**
   ```powershell
   # Backup current config
   Copy-Item app_config.json app_config.json.backup
   
   # Delete to regenerate
   Remove-Item app_config.json
   
   # Restart FileLabeler
   ```

5. **Review log for details:**
   ```
   [ERROR] [Load-AppConfig] Config validation failed
   ```

---

### 🔴 Async Operations Disabled

**Symptoms:**
- Log shows: `[WARNING] [AsyncInitialization] Failed to create runspace pool`
- Slow performance with large file sets
- UI feels sluggish

**Impact:**
- Application still works but slower
- Operations are synchronous instead of async
- >30 files may feel unresponsive

**Solutions:**

1. **Check PowerShell runspace support:**
   ```powershell
   [RunspaceFactory]::CreateRunspacePool(1, 4) | Out-Null
   Write-Output "Runspace pool support: OK"
   ```

2. **Verify .NET Framework version:**
   - Ensure .NET Framework 4.7.2 or later
   - Update Windows if needed

3. **Check for conflicting modules:**
   ```powershell
   Get-Module | Format-Table Name, Version
   ```

4. **Restart application:**
   - Close and reopen FileLabeler
   - Check if async initializes successfully

5. **Fallback to sync mode:**
   - Application will work in sync mode (slower but functional)
   - Consider processing files in smaller batches (20-30 files)

---

### 🔴 Norwegian Characters Display Incorrectly

**Symptoms:**
- æ, ø, å show as Ã¥, Ã¸, Ã¦ or other strange characters
- UI text appears garbled

**Solution:**

**For Script Users:**
1. Open `FileLabeler.ps1` in VS Code or Cursor
2. Bottom right → Click encoding indicator
3. Select "Save with Encoding"
4. Choose **"UTF-8 with BOM"**
5. Save file
6. Restart FileLabeler

**For EXE Users:**
- This should not happen with compiled EXE
- If it does, recompile from UTF-8 BOM source

---

## Performance Issues

### 🐌 Slow File Processing

**Expected Performance:**
- 1-5 files: < 5 seconds
- 10-30 files: 10-30 seconds (sync mode)
- 30-100 files: 15-45 seconds (async mode)
- 100+ files: ~0.5 sec/file (async with parallelism)

**If Slower:**

1. **Check file count vs threshold:**
   - Async kicks in at 30+ files for batch apply
   - Check log: `[INFO] [Startup] Async operations: Enabled/Disabled`

2. **Network folder latency:**
   - Network files take longer
   - OneDrive sync can add 2-3 seconds per file
   - Consider copying files locally first

3. **Large file sizes:**
   - Files >50MB take longer
   - This is normal behavior

4. **AIP sync delay:**
   - Microsoft Purview sync can add time
   - Check Purview client status

5. **Enable async operations:**
   - Check `app_config.json` → `async.enableAsyncOperations: true`
   - Lower thresholds for faster async activation

---

### 🧠 Memory Issues

**Symptoms:**
- Application becomes slow over time
- Windows shows "Low memory" warning
- Error: `OutOfMemoryException`

**Solutions:**

1. **Close other applications:**
   - Free up RAM
   - Close browser tabs
   - Close unnecessary Office applications

2. **Process in smaller batches:**
   - Limit to 50-100 files per batch
   - Clear selection and repeat

3. **Restart FileLabeler:**
   - Closes all runspaces
   - Frees up memory

4. **Check available RAM:**
   ```powershell
   Get-CimInstance Win32_OperatingSystem | 
       Select-Object TotalVisibleMemorySize, FreePhysicalMemory
   ```
   - Recommended: 8GB+ for large batches (300+ files)

---

## Log File Analysis

### Log Location

```
C:\Users\<username>\Documents\FileLabeler_Logs\
FileLabeler_Log_yyyyMMdd_HHmmss.txt
```

### Log Levels

| Level | Meaning | Action Required |
|-------|---------|-----------------|
| INFO | Normal operation | None - informational |
| WARNING | Non-critical issue | Monitor, may need attention |
| ERROR | Operation failed | Investigate and resolve |
| CRITICAL | Application-stopping | Immediate action required |

### Finding Errors

**All errors:**
```powershell
Select-String -Path "FileLabeler_Log_*.txt" -Pattern "\[ERROR\]" -Context 0,3
```

**Critical issues:**
```powershell
Select-String -Path "FileLabeler_Log_*.txt" -Pattern "\[CRITICAL\]" -Context 0,5
```

**Specific file errors:**
```powershell
Select-String -Path "FileLabeler_Log_*.txt" -Pattern "filename.docx" | 
    Select-String -Pattern "\[ERROR\]"
```

---

## Error Categories

| Category | User Message (Norwegian) | Suggestion |
|----------|-------------------------|------------|
| **FileAccess** | Ingen tilgang til filen | Check file/folder permissions |
| **FileLocked** | Filen er i bruk | Close file in Office applications |
| **FileNotFound** | Filen ble ikke funnet | Verify file still exists |
| **Network** | Nettverksfeil | Check network connection/VPN |
| **AIPJustification** | Begrunnelse kreves | Provide justification for downgrade |
| **AIPProtection** | Beskyttelsesinnstillinger kreves | Configure protection settings |
| **Memory** | Ikke nok minne | Close apps, process fewer files |
| **Timeout** | Operasjonen tok for lang tid | Retry with fewer files |

---

## Advanced Diagnostics

### Enable Detailed Logging

**Method 1: Via UI**
1. Click "Innstillinger" (Settings)
2. Go to "Logging" tab
3. Check "Aktiver detaljert logging"
4. Click "Lagre" (Save)

**Method 2: Via Config File**
Edit `app_config.json`:
```json
{
  "logging": {
    "enableDetailedLogging": true
  }
}
```

### Check Windows Event Log

Critical errors are also written to Windows Event Log:

1. Open Event Viewer (`eventvwr.msc`)
2. Navigate to **Windows Logs** → **Application**
3. Filter by Source: **"FileLabeler"**
4. Look for Error entries

### PowerShell Diagnostics

**Check AIP Module:**
```powershell
Get-Module -ListAvailable -Name PurviewInformationProtection
Get-Command -Module PurviewInformationProtection
```

**Test AIP Cmdlets:**
```powershell
Get-AIPFileStatus -Path "C:\path\to\test-file.docx"
```

**Check Runspace Support:**
```powershell
[RunspaceFactory]::CreateRunspacePool(1, 4) | Out-Null
Write-Output "Runspace pool support: OK"
```

---

## Collecting Support Information

When reporting issues, provide:

1. **Log File:** Most recent `FileLabeler_Log_*.txt`
2. **Error Screenshot:** If error dialog shown
3. **Environment:**
   ```powershell
   # Collect diagnostic info
   @"
   Windows: $(Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption)
   PowerShell: $($PSVersionTable.PSVersion)
   AIP Module: $(Get-Module -ListAvailable -Name PurviewInformationProtection | Select-Object -ExpandProperty Version)
   FileLabeler Version: 1.1
   "@ | Out-File "$env:USERPROFILE\Desktop\FileLabeler_Diagnostics.txt"
   ```

4. **Reproduction Steps:** What were you doing when error occurred?
5. **File Count:** How many files were being processed?
6. **Location:** Local files or network/OneDrive?

---

## Frequently Asked Questions

### Q: Files processed but dates changed when I opened them?

**A:** This is expected behavior when labeling updates metadata. The `-PreserveFileDetails` flag works correctly, but Office applications detect the metadata change and auto-save, which updates the modified date.

See [DOCUMENT_AUTOSAVE_INFO.md](DOCUMENT_AUTOSAVE_INFO.md) in archive for detailed explanation.

### Q: Can I remove labels from files?

**A:** Currently not supported. FileLabeler is designed for applying labels, not removing them. Use Microsoft Purview client for label removal.

### Q: Why do some protected files show "Ukjent etikett (beskyttet)"?

**A:** The file has an encrypted label not in your `labels_config.json`. This is normal. Add the label ID to config if you need to work with these files.

### Q: Can I use FileLabeler on macOS or Linux?

**A:** No. FileLabeler requires Windows PowerShell and Microsoft Purview client, which are Windows-only.

### Q: How do I update to a new version?

**A:** Download new version, backup your `labels_config.json`, replace files, and keep your config file. See [INSTALLATION.md](INSTALLATION.md#updating-to-new-version).

---

## Still Need Help?

If your issue isn't covered here:

1. **Search existing GitHub issues:** [Issues](https://github.com/yourusername/FileLabeler/issues)
2. **Check discussions:** [Discussions](https://github.com/yourusername/FileLabeler/discussions)
3. **Create a new issue:** Include diagnostic information from above

---

**Most common solution:** Restart the application and try with a smaller batch of files first! 🔄

