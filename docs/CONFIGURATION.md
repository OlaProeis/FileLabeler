# Configuration Guide

This guide covers all configuration options for FileLabeler.

---

## Configuration Files

FileLabeler uses two main configuration files:

| File | Purpose | Required |
|------|---------|----------|
| `labels_config.json` | Sensitivity label definitions | ✅ Yes |
| `app_config.json` | Application settings and preferences | ⚠️ Auto-generated |

---

## Label Configuration (`labels_config.json`)

### File Location

The `labels_config.json` file must be in the same directory as `FileLabeler.ps1` or `FileLabeler.exe`.

### Basic Structure

```json
[
  {
    "DisplayName": "Public",
    "Id": "8bc7810f-1601-4c5d-8eaa-56870a5bf913",
    "Rank": 0
  },
  {
    "DisplayName": "Internal",
    "Id": "e6f3cebf-f3ea-4f6d-9afa-b2867c184242",
    "Rank": 1
  },
  {
    "DisplayName": "Confidential",
    "Id": "221e033f-836b-4372-a276-90a25fdd73b5",
    "Rank": 2
  },
  {
    "DisplayName": "Highly Confidential",
    "Id": "db735bfd-96f4-488e-b27d-b95706dd8a4e",
    "Rank": 3,
    "RequiresProtection": true
  }
]
```

### Field Descriptions

#### DisplayName (Required)
- **Type:** String
- **Purpose:** Label name shown in the UI
- **Example:** `"Fortrolig"`, `"Confidential"`
- **Notes:** Can use Norwegian or English based on your organization

#### Id (Required)
- **Type:** GUID string
- **Purpose:** Unique identifier for the sensitivity label in Microsoft Purview
- **Example:** `"221e033f-836b-4372-a276-90a25fdd73b5"`
- **How to get:** See [Getting Label IDs](#getting-label-ids) below

#### Rank (Required)
- **Type:** Integer (0-10)
- **Purpose:** Sensitivity level for downgrade detection
- **Scale:**
  - `0` = Lowest sensitivity (e.g., "Public")
  - Higher numbers = Higher sensitivity
  - Example: `0=Public, 1=Internal, 2=Confidential, 3=Highly Confidential`
- **Usage:** Application detects downgrades and prompts for justification when applying a lower-ranked label

#### RequiresProtection (Optional)
- **Type:** Boolean
- **Purpose:** Indicates if label requires access control settings
- **Default:** `false`
- **When true:** Application shows protection dialog for permission settings
- **Example:**
  ```json
  {
    "DisplayName": "Highly Confidential",
    "Id": "db735bfd-96f4-488e-b27d-b95706dd8a4e",
    "Rank": 3,
    "RequiresProtection": true
  }
  ```

---

## Getting Label IDs

There are four methods to retrieve your organization's sensitivity label IDs:

### Method 1: PowerShell - Security & Compliance Center (Recommended)

**Requirements:** Security & Compliance Center PowerShell module

```powershell
# Install module (if not already installed)
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# Connect to Security & Compliance Center
Connect-IPPSSession

# Get all sensitivity labels
Get-Label | Select-Object DisplayName, Guid, Id | Format-Table -AutoSize

# Save to file for reference
Get-Label | Select-Object DisplayName, Guid, Id | Export-Csv "labels.csv" -NoTypeInformation

# Disconnect when done
Disconnect-ExchangeOnline
```

### Method 2: Microsoft Purview Compliance Portal (Web UI)

1. Navigate to [https://compliance.microsoft.com](https://compliance.microsoft.com)
2. Go to **Information Protection** → **Labels**
3. Click on each label to view properties
4. Copy the **Label ID** (GUID) from the properties pane
5. Note the display name and create your `labels_config.json`

### Method 3: Using a Test File (Easiest for Quick Setup)

1. Manually apply a sensitivity label to a test document in Word, Excel, or PowerPoint
2. Save and close the document
3. Run this PowerShell command:

```powershell
Import-Module PurviewInformationProtection

# Check a labeled file
Get-AIPFileStatus -Path "C:\path\to\labeled-file.docx" | 
    Select-Object FileName, LabelId, LabelName | 
    Format-List

# Check multiple files with different labels
Get-ChildItem "C:\path\to\testfiles" -Include *.docx,*.xlsx -Recurse |
    ForEach-Object {
        $status = Get-AIPFileStatus -Path $_.FullName
        [PSCustomObject]@{
            File = $_.Name
            Label = $status.LabelName
            LabelId = $status.LabelId
        }
    } | Format-Table -AutoSize
```

### Method 4: Microsoft Graph PowerShell

```powershell
# Install Microsoft Graph PowerShell (if needed)
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "InformationProtectionPolicy.Read"

# Get sensitivity labels
Get-MgInformationProtectionLabel | 
    Select-Object DisplayName, Id | 
    Format-Table -AutoSize

# Export to JSON format for easy copying
Get-MgInformationProtectionLabel | 
    Select-Object @{N='DisplayName';E={$_.DisplayName}},
                  @{N='Id';E={$_.Id}},
                  @{N='Rank';E={0}} | 
    ConvertTo-Json | 
    Out-File "labels_template.json"

# Disconnect when done
Disconnect-MgGraph
```

---

## Application Configuration (`app_config.json`)

This file is **auto-generated** on first run. You can modify settings through the UI (Settings button) or edit the file directly.

### File Location

`app_config.json` is created in the same directory as the application.

### Default Configuration

```json
{
  "version": "1.0",
  "ui": {
    "skipUnchangedFiles": false,
    "showDetailedProgress": true
  },
  "warnings": {
    "enableMassDowngradeWarning": true,
    "enableLargeBatchWarning": true,
    "largeBatchThreshold": 50,
    "enableNoChangesWarning": true,
    "enableProtectionWarning": true,
    "enableMixedChangesWarning": false
  },
  "async": {
    "enableAsyncOperations": true,
    "folderScanThreshold": "auto",
    "labelRetrievalThreshold": 50,
    "batchApplyThreshold": 30,
    "maxConcurrentThreads": 4
  },
  "logging": {
    "enableDetailedLogging": false,
    "logRetentionDays": 30
  }
}
```

### Configuration Sections

#### UI Settings
- **skipUnchangedFiles**: Skip files with unchanged labels during application
- **showDetailedProgress**: Show detailed progress information

#### Warning Settings
- **enableMassDowngradeWarning**: Warn when downgrading many files
- **enableLargeBatchWarning**: Warn for large file batches
- **largeBatchThreshold**: File count that triggers large batch warning
- **enableNoChangesWarning**: Warn when no files will be changed
- **enableProtectionWarning**: Warn about protection requirements
- **enableMixedChangesWarning**: Warn about mixed upgrade/downgrade operations

#### Async Operation Settings
- **enableAsyncOperations**: Enable async operations for better performance
- **folderScanThreshold**: Threshold for async folder scanning (`"auto"` recommended)
- **labelRetrievalThreshold**: File count to trigger async label retrieval (default: 50)
- **batchApplyThreshold**: File count to trigger async batch application (default: 30)
- **maxConcurrentThreads**: Max threads for parallel operations (1-8)

#### Logging Settings
- **enableDetailedLogging**: Enable verbose logging for troubleshooting
- **logRetentionDays**: How long to keep log files (default: 30 days)

---

## Configuration Through UI

Most settings can be configured through the application's Settings dialog:

1. Open FileLabeler
2. Click **"Innstillinger"** (Settings) button
3. Navigate through tabs:
   - **Warnings** - Configure warning preferences
   - **Performance** - Async operation thresholds
   - **Logging** - Log verbosity and retention

4. Click **"Lagre"** (Save) to apply changes

---

## Advanced Configuration

### Custom Label Ranking

You can define custom ranking systems:

```json
[
  {
    "DisplayName": "Public",
    "Id": "...",
    "Rank": 0
  },
  {
    "DisplayName": "Internal - Low",
    "Id": "...",
    "Rank": 1
  },
  {
    "DisplayName": "Internal - Medium",
    "Id": "...",
    "Rank": 2
  },
  {
    "DisplayName": "Internal - High",
    "Id": "...",
    "Rank": 3
  },
  {
    "DisplayName": "Confidential",
    "Id": "...",
    "Rank": 5
  },
  {
    "DisplayName": "Highly Confidential",
    "Id": "...",
    "Rank": 10,
    "RequiresProtection": true
  }
]
```

**Note:** Gaps in rank numbers are allowed and can help represent sensitivity levels.

### Performance Tuning

For organizations with specific needs:

```json
{
  "async": {
    "enableAsyncOperations": true,
    "folderScanThreshold": "auto",
    "labelRetrievalThreshold": 25,     // Lower for faster UI on slower networks
    "batchApplyThreshold": 20,         // Lower for more responsive UI
    "maxConcurrentThreads": 2          // Lower for systems with limited resources
  }
}
```

**Guidelines:**
- **Fast networks/systems**: Higher thresholds (50-100)
- **Slow networks/systems**: Lower thresholds (20-30)
- **Limited CPU cores**: Lower `maxConcurrentThreads` (1-2)
- **Powerful systems**: Higher `maxConcurrentThreads` (4-8)

---

## Configuration Validation

FileLabeler validates configuration on startup:

### Labels Configuration Validation
- ✅ Valid JSON format
- ✅ All required fields present (DisplayName, Id, Rank)
- ✅ No duplicate label IDs
- ✅ Rank values are integers
- ✅ Label IDs are valid GUIDs

### App Configuration Validation
- ✅ Valid JSON format
- ✅ Threshold values within acceptable ranges
- ✅ Boolean settings are true/false
- ✅ Version compatibility

**If validation fails:**
- Application shows error message
- Creates backup: `app_config.json.invalid_[timestamp]`
- Loads default configuration

---

## Troubleshooting Configuration

### "No Labels Configured" Error

**Problem:** `labels_config.json` missing or empty

**Solution:**
1. Create `labels_config.json` in application directory
2. Use templates from this guide
3. Get label IDs using methods above

### "Invalid Label Configuration" Error

**Problem:** JSON syntax error or missing required fields

**Solution:**
1. Validate JSON syntax: [jsonlint.com](https://jsonlint.com)
2. Check all required fields present
3. Ensure proper quote marks and commas
4. Verify GUID format for label IDs

### Labels Not Appearing

**Problem:** Labels configured but not showing in UI

**Solution:**
1. Check `labels_config.json` is in same directory as application
2. Verify JSON is valid
3. Check application log for errors
4. Restart application

### Configuration Reset

To reset configuration to defaults:

```powershell
# Backup current config
Copy-Item app_config.json app_config.json.backup

# Delete current config (will regenerate on next run)
Remove-Item app_config.json

# Or rename to force regeneration
Rename-Item app_config.json app_config.json.old
```

---

## Example Configurations

### Norwegian Organization (Full Example)

```json
[
  {
    "DisplayName": "Åpen",
    "Id": "8bc7810f-1601-4c5d-8eaa-56870a5bf913",
    "Rank": 0
  },
  {
    "DisplayName": "Intern",
    "Id": "e6f3cebf-f3ea-4f6d-9afa-b2867c184242",
    "Rank": 1
  },
  {
    "DisplayName": "Personlig",
    "Id": "3f9e8c7a-2b1d-4e5f-8a9b-1c2d3e4f5a6b",
    "Rank": 2
  },
  {
    "DisplayName": "Privat",
    "Id": "4a5b6c7d-8e9f-1a2b-3c4d-5e6f7a8b9c0d",
    "Rank": 3
  },
  {
    "DisplayName": "Fortrolig",
    "Id": "221e033f-836b-4372-a276-90a25fdd73b5",
    "Rank": 4
  },
  {
    "DisplayName": "Strengt Fortrolig",
    "Id": "db735bfd-96f4-488e-b27d-b95706dd8a4e",
    "Rank": 7,
    "RequiresProtection": true
  }
]
```

### English Organization (Full Example)

```json
[
  {
    "DisplayName": "Public",
    "Id": "8bc7810f-1601-4c5d-8eaa-56870a5bf913",
    "Rank": 0
  },
  {
    "DisplayName": "General",
    "Id": "e6f3cebf-f3ea-4f6d-9afa-b2867c184242",
    "Rank": 1
  },
  {
    "DisplayName": "Confidential",
    "Id": "221e033f-836b-4372-a276-90a25fdd73b5",
    "Rank": 2
  },
  {
    "DisplayName": "Highly Confidential",
    "Id": "db735bfd-96f4-488e-b27d-b95706dd8a4e",
    "Rank": 3,
    "RequiresProtection": true
  }
]
```

---

## Next Steps

After configuration:

1. **Test with sample files** before production use
2. **Review settings** in the Settings dialog
3. **Check logs** to verify label retrieval works
4. **Read the User Guide** for operational instructions → [User Guide](USER_GUIDE.md)

---

**Configuration complete!** If you encounter issues, check [Troubleshooting](TROUBLESHOOTING.md).

