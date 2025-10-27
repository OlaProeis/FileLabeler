# Installation Guide

This guide covers all installation methods for FileLabeler.

---

## System Requirements

### Minimum Requirements

| Component | Requirement |
|-----------|-------------|
| **Operating System** | Windows 10/11 (64-bit) |
| **PowerShell** | Version 5.1 or later (included with Windows) |
| **Microsoft Purview Client** | Required - see below |
| **.NET Framework** | 4.7.2 or later (usually pre-installed) |
| **Disk Space** | ~50 MB (application + logs) |
| **RAM** | 4 GB minimum, 8 GB recommended for large batches |

### Required Software

**Microsoft Purview Information Protection Client**  
Download: [Microsoft Download Center](https://www.microsoft.com/en-us/download/details.aspx?id=53018)

This installs the `PurviewInformationProtection` PowerShell module required for labeling operations.

---

## Installation Methods

### Method 1: Run as PowerShell Script (Recommended for Development)

**Advantages:**
- Easy to modify and update
- No compilation needed
- Full control over execution

**Steps:**

1. **Download the project:**
   ```powershell
   # Using Git
   git clone https://github.com/yourusername/FileLabeler.git
   cd FileLabeler
   
   # Or download and extract ZIP from GitHub
   ```

2. **Verify files:**
   ```powershell
   # Check required files exist
   Test-Path FileLabeler.ps1       # Should return True
   Test-Path labels_config.json    # Should return True
   ```

3. **Configure labels:**
   - Edit `labels_config.json` with your organization's labels
   - See [Configuration Guide](CONFIGURATION.md) for details

4. **Run the application:**
   ```powershell
   .\FileLabeler.ps1
   ```

5. **Optional: Create a desktop shortcut:**
   ```powershell
   $WScriptShell = New-Object -ComObject WScript.Shell
   $Shortcut = $WScriptShell.CreateShortcut("$env:USERPROFILE\Desktop\FileLabeler.lnk")
   $Shortcut.TargetPath = "powershell.exe"
   $Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$PWD\FileLabeler.ps1`""
   $Shortcut.IconLocation = "imageres.dll,1"
   $Shortcut.Save()
   ```

---

### Method 2: Convert to Standalone EXE (Recommended for Distribution)

**Advantages:**
- Single executable file
- No visible PowerShell window
- Professional deployment
- Easier for non-technical users

**Steps:**

1. **Install PS2EXE module:**
   ```powershell
   Install-Module ps2exe -Scope CurrentUser -Force
   ```

2. **Convert to EXE:**
   ```powershell
   Invoke-ps2exe -inputFile .\FileLabeler.ps1 `
                 -outputFile .\FileLabeler.exe `
                 -noConsole `
                 -requireAdmin `
                 -title "FileLabeler" `
                 -description "Bulk Sensitivity Label Application Tool" `
                 -company "Your Organization" `
                 -copyright "Copyright 2025" `
                 -version "1.1.0.0"
   ```

3. **Optional: Add icon:**
   ```powershell
   Invoke-ps2exe -inputFile .\FileLabeler.ps1 `
                 -outputFile .\FileLabeler.exe `
                 -iconFile .\icon.ico `
                 -noConsole
   ```

4. **Deploy:**
   - Copy `FileLabeler.exe` and `labels_config.json` to target machines
   - Both files must be in the same directory

---

### Method 3: Intune Deployment (Enterprise)

For deploying to multiple machines via Microsoft Intune:

1. **Package the application:**
   ```powershell
   # Create deployment folder
   New-Item -Path "C:\Temp\FileLabelerDeploy" -ItemType Directory -Force
   
   # Copy files
   Copy-Item FileLabeler.exe "C:\Temp\FileLabelerDeploy\"
   Copy-Item labels_config.json "C:\Temp\FileLabelerDeploy\"
   ```

2. **Create installation script (`Install-FileLabeler.ps1`):**
   ```powershell
   # Install location
   $InstallPath = "$env:ProgramFiles\FileLabeler"
   
   # Create folder
   New-Item -Path $InstallPath -ItemType Directory -Force
   
   # Copy files
   Copy-Item "FileLabeler.exe" $InstallPath -Force
   Copy-Item "labels_config.json" $InstallPath -Force
   
   # Create Start Menu shortcut
   $WScriptShell = New-Object -ComObject WScript.Shell
   $Shortcut = $WScriptShell.CreateShortcut("$env:ProgramData\Microsoft\Windows\Start Menu\Programs\FileLabeler.lnk")
   $Shortcut.TargetPath = "$InstallPath\FileLabeler.exe"
   $Shortcut.Save()
   
   Write-Output "FileLabeler installed successfully"
   ```

3. **Create uninstall script (`Uninstall-FileLabeler.ps1`):**
   ```powershell
   # Remove installation
   Remove-Item "$env:ProgramFiles\FileLabeler" -Recurse -Force -ErrorAction SilentlyContinue
   Remove-Item "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\FileLabeler.lnk" -Force -ErrorAction SilentlyContinue
   
   Write-Output "FileLabeler uninstalled successfully"
   ```

4. **Package and upload to Intune:**
   - Create `.intunewin` package using Microsoft Win32 Content Prep Tool
   - Upload to Intune portal
   - Configure installation and uninstallation commands
   - Assign to target groups

---

## Post-Installation Setup

### 1. Verify Module Installation

```powershell
# Check PowerShell version
$PSVersionTable.PSVersion

# Check for PurviewInformationProtection module
Get-Module -ListAvailable -Name PurviewInformationProtection

# If module not found, install Purview client from Microsoft
```

### 2. Configure Labels

Edit `labels_config.json` with your organization's sensitivity labels:

```json
[
  {
    "DisplayName": "Public",
    "Id": "your-label-guid-here",
    "Rank": 0
  },
  {
    "DisplayName": "Internal",
    "Id": "your-label-guid-here",
    "Rank": 1
  },
  {
    "DisplayName": "Confidential",
    "Id": "your-label-guid-here",
    "Rank": 2,
    "RequiresProtection": true
  }
]
```

**Getting Label IDs:**  
See [Configuration Guide - Label IDs](CONFIGURATION.md#getting-label-ids) for detailed instructions.

### 3. Test Installation

```powershell
# Run a quick test
.\FileLabeler.ps1  # or .\FileLabeler.exe

# Check logs
Get-ChildItem "$env:USERPROFILE\Documents\FileLabeler_Logs\"
```

---

## Troubleshooting Installation

### "Missing Required Module" Error

**Problem:** `PurviewInformationProtection` module not found

**Solution:**
1. Download and install [Microsoft Purview Information Protection Client](https://www.microsoft.com/en-us/download/details.aspx?id=53018)
2. Restart PowerShell
3. Verify installation:
   ```powershell
   Get-Module -ListAvailable -Name PurviewInformationProtection
   ```

### "Execution Policy" Error

**Problem:** Script execution blocked

**Solution:**
```powershell
# Temporary bypass (current session)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Or use the shortcut method from Method 1
```

### "Access Denied" Error

**Problem:** Insufficient permissions

**Solution:**
- Run PowerShell as Administrator
- Or use the `-requireAdmin` flag when creating EXE

### Norwegian Characters Display Incorrectly

**Problem:** æ, ø, å show as strange characters

**Solution:**
- Ensure `FileLabeler.ps1` is saved as **UTF-8 with BOM**
- In VS Code/Cursor: Bottom right → "Save with Encoding" → "UTF-8 with BOM"

---

## Network Deployment Considerations

### Shared Network Installation

If deploying to a network share:

```powershell
# Install to network share
$NetworkPath = "\\server\share\Applications\FileLabeler"
Copy-Item FileLabeler.exe $NetworkPath -Force
Copy-Item labels_config.json $NetworkPath -Force

# Create shortcut on user desktops
$Shortcut = $WScriptShell.CreateShortcut("$env:USERPROFILE\Desktop\FileLabeler.lnk")
$Shortcut.TargetPath = "$NetworkPath\FileLabeler.exe"
$Shortcut.Save()
```

**Note:** Users must have access to the network share.

### Group Policy Deployment

For large-scale deployment via GPO:

1. Copy files to NETLOGON share
2. Create GPO startup/login script
3. Distribute to target OUs

---

## Updating to New Version

### For Script Installation:
```powershell
# Backup current version
Copy-Item FileLabeler.ps1 FileLabeler.ps1.backup

# Download new version
git pull  # or download new files

# Preserve your labels_config.json
# (Don't overwrite unless label structure changed)
```

### For EXE Installation:
1. Download new version
2. Recompile with PS2EXE
3. Replace old EXE
4. Keep existing `labels_config.json` unless changed

---

## Uninstallation

### Remove Application:
```powershell
# Remove application files
Remove-Item "C:\Path\To\FileLabeler" -Recurse -Force

# Remove shortcuts
Remove-Item "$env:USERPROFILE\Desktop\FileLabeler.lnk" -Force

# Optional: Remove log files
Remove-Item "$env:USERPROFILE\Documents\FileLabeler_Logs" -Recurse -Force
```

### Remove Purview Client:
- Go to Settings → Apps → Microsoft Purview Information Protection
- Click Uninstall

---

## Next Steps

After installation:

1. **Configure labels** → [Configuration Guide](CONFIGURATION.md)
2. **Learn to use the app** → [User Guide](USER_GUIDE.md)
3. **Test with a few files** before bulk operations

---

**Installation complete!** If you encounter issues, check [Troubleshooting](TROUBLESHOOTING.md).

