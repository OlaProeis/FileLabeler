# FileLabeler

> **Bulk Sensitivity Label Application for Microsoft Purview**

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-lightgrey.svg)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## Overview

**FileLabeler** is a PowerShell-based GUI application that enables users to quickly apply Microsoft Purview sensitivity labels to multiple files simultaneously. The application preserves original file dates and provides a modern, user-friendly interface for bulk labeling tasks.

---

## Key Features

- ✅ **Bulk Labeling** - Label multiple files at once
- ✅ **Date Preservation** - Preserves original file timestamps
- ✅ **Folder Import** - Import entire folders with subdirectories
- ✅ **Drag-and-Drop** - Drag files directly from Windows Explorer
- ✅ **Smart Preview** - Intelligent pre-apply analysis with warnings
- ✅ **Detailed Statistics** - Comprehensive results and reporting
- ✅ **Access Control** - Full support for protection settings
- ✅ **Downgrade Handling** - Automatic justification prompts when required
- ✅ **Async Operations** - Responsive UI for large datasets (500+ files)
- ✅ **Comprehensive Testing** - 106 tests, 100% pass rate

---

## Quick Start

### Requirements

- **OS**: Windows 10/11 (64-bit)
- **PowerShell**: 5.1 or later (included with Windows)
- **Microsoft Purview Information Protection Client**: [Download](https://www.microsoft.com/en-us/download/details.aspx?id=53018)

### Installation

#### Method 1: Run as Script
```powershell
# Download the project
git clone https://github.com/yourusername/FileLabeler.git
cd FileLabeler

# Run the application
.\FileLabeler.ps1
```

#### Method 2: Convert to EXE
```powershell
# Install PS2EXE
Install-Module ps2exe -Scope CurrentUser

# Convert to executable
Invoke-ps2exe -inputFile .\FileLabeler.ps1 `
              -outputFile .\FileLabeler.exe `
              -noConsole `
              -title "FileLabeler" `
              -version "1.1.0.0"
```

### Configuration

Create `labels_config.json` with your organization's labels:

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
    "Rank": 2
  }
]
```

See [Configuration Guide](docs/CONFIGURATION.md) for details on obtaining label IDs.

---

## Documentation

### For Users
- **[Installation Guide](docs/INSTALLATION.md)** - Detailed installation instructions
- **[Configuration Guide](docs/CONFIGURATION.md)** - Label setup and settings
- **[User Guide](docs/USER_GUIDE.md)** - Complete usage guide (Norwegian)
- **[Troubleshooting](docs/TROUBLESHOOTING.md)** - Common issues and solutions

### For Developers
- **[Architecture](docs/development/ARCHITECTURE.md)** - Technical overview
- **[Testing Guide](docs/development/TESTING.md)** - Test suite and quality assurance
- **[Contributing](docs/development/CONTRIBUTING.md)** - Contribution guidelines

### Other Resources
- **[Changelog](docs/CHANGELOG.md)** - Detailed version history
- **[Roadmap](docs/ROADMAP.md)** - Planned features

---

## Usage Example

```powershell
# 1. Start the application
.\FileLabeler.ps1

# 2. Select files
#    - Click "Velg filer..." (Select files)
#    - Or drag files from Explorer
#    - Or select folder with "Velg mappe..." (Select folder)

# 3. Select label
#    - Click desired label button

# 4. Review and apply
#    - Click "Påfør etikett (bevar datoer)" (Apply label - preserve dates)
#    - Review summary
#    - Confirm application

# 5. Export results
#    - Click "Eksporter til CSV" (Export to CSV)
```

---

## Testing

FileLabeler has comprehensive test coverage:

```powershell
# Run all tests
Invoke-Pester -Path .\tests\

# Run unit tests
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1

# Run integration tests
.\run_integration_tests.ps1
```

**Test Results:**
- 58 unit tests
- 48 integration tests
- 100% pass rate
- Full feature coverage

---

## Comparison with Purview Client

| Feature | FileLabeler | Purview Client |
|---------|-------------|----------------|
| Install Size | ~5 MB | ~150 MB |
| Date Preservation | ✅ Always | ⚠️ Changes date |
| Standalone App | ✅ Yes | ❌ Context menu only |
| Bulk Labeling | ✅ Optimized | ⚠️ Limited |
| Visual Feedback | ✅ Progress bar | ❌ None |
| Detailed Logging | ✅ Yes | ⚠️ Limited |
| CSV Export | ✅ Yes | ❌ No |

---

## Tech Stack

- **Language**: PowerShell 5.1+
- **GUI**: Windows Forms (`System.Windows.Forms`)
- **Labeling**: Microsoft PurviewInformationProtection module
- **Testing**: Pester 3.4.0+
- **Platform**: Windows 10/11

---

## Version

**Current Version**: v1.1  
**Status**: Production Ready

See [CHANGELOG.md](docs/CHANGELOG.md) for complete version history.

---

## Contributing

We welcome contributions!

1. Fork the project
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

See [CONTRIBUTING.md](docs/development/CONTRIBUTING.md) for detailed guidelines.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- Microsoft Purview Information Protection team
- PowerShell community
- All contributors and testers

---

## Support

- **Documentation**: [docs/](docs/)
- **Issues**: [GitHub Issues](https://github.com/yourusername/FileLabeler/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/FileLabeler/discussions)

---

**Developed with ❤️ for efficient sensitivity label management**
