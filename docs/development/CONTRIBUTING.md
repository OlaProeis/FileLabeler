# Contributing to FileLabeler

Thank you for your interest in contributing to FileLabeler! This document provides guidelines and instructions for contributing.

---

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Workflow](#development-workflow)
- [Coding Standards](#coding-standards)
- [Testing Requirements](#testing-requirements)
- [Pull Request Process](#pull-request-process)
- [Reporting Bugs](#reporting-bugs)
- [Suggesting Features](#suggesting-features)

---

## Code of Conduct

### Our Pledge

We are committed to making participation in this project a harassment-free experience for everyone, regardless of level of experience, gender, gender identity and expression, sexual orientation, disability, personal appearance, body size, race, ethnicity, age, religion, or nationality.

### Our Standards

**Positive behavior includes:**
- Using welcoming and inclusive language
- Being respectful of differing viewpoints
- Gracefully accepting constructive criticism
- Focusing on what is best for the community
- Showing empathy towards other community members

**Unacceptable behavior includes:**
- Harassment, trolling, or derogatory comments
- Public or private harassment
- Publishing others' private information without permission
- Other conduct reasonably considered inappropriate

---

## Getting Started

### Prerequisites

Before contributing, ensure you have:

1. **Windows 10/11** development machine
2. **PowerShell 5.1+** installed
3. **Git** for version control
4. **VS Code or Cursor** (recommended IDE)
5. **Pester 3.4.0+** for testing
6. **Microsoft Purview Information Protection Client** (for testing with real labels)

### Setting Up Development Environment

```powershell
# Clone the repository
git clone https://github.com/yourusername/FileLabeler.git
cd FileLabeler

# Create a development branch
git checkout -b feature/your-feature-name

# Open in your IDE
code .  # or cursor .

# Run tests to ensure everything works
Invoke-Pester -Path .\tests\
```

### Understanding the Codebase

Before making changes:

1. Read [ARCHITECTURE.md](ARCHITECTURE.md) - Understand the technical design
2. Review [TESTING.md](TESTING.md) - Understand testing strategy
3. Check [ROADMAP.md](../ROADMAP.md) - See planned features
4. Browse existing code - Understand conventions and patterns

---

## Development Workflow

### 1. Find or Create an Issue

Before starting work:

- **Check existing issues** to avoid duplicate work
- **Open a new issue** if your feature/bug isn't tracked
- **Discuss major changes** in issue before implementation
- **Get feedback** from maintainers before significant work

### 2. Create a Feature Branch

```powershell
# Always branch from main
git checkout main
git pull origin main

# Create descriptive branch name
git checkout -b feature/add-file-search
# or
git checkout -b bugfix/fix-norwegian-encoding
# or
git checkout -b docs/update-installation-guide
```

**Branch naming conventions:**
- `feature/` - New features
- `bugfix/` - Bug fixes
- `docs/` - Documentation updates
- `refactor/` - Code improvements without functional changes
- `test/` - Test additions or improvements

### 3. Make Your Changes

Follow these practices:

#### Write Clean Code
- Follow [Coding Standards](#coding-standards) below
- Keep functions focused and small
- Use descriptive names
- Comment complex logic
- Avoid code duplication

#### Maintain Backward Compatibility
- Don't break existing functionality
- Preserve configuration file formats
- Maintain API stability

#### Test Thoroughly
- Add unit tests for new functions
- Add integration tests for workflows
- Run all tests before committing
- Test with real files when possible

### 4. Commit Your Changes

```powershell
# Stage changes
git add .

# Commit with descriptive message
git commit -m "feat: Add file search functionality

- Added search box to filter file list
- Implemented real-time filtering
- Added file count indicator
- Includes unit tests

Fixes #123"
```

**Commit message format:**
```
<type>: <subject>

<body>

<footer>
```

**Types:**
- `feat:` New feature
- `fix:` Bug fix
- `docs:` Documentation changes
- `style:` Formatting changes (no code change)
- `refactor:` Code refactoring
- `test:` Adding or updating tests
- `chore:` Maintenance tasks

**Examples:**
```
feat: Add drag-and-drop support for folders

fix: Resolve Norwegian character encoding issue

docs: Update installation guide for Windows 11

test: Add integration tests for label application

refactor: Consolidate file scanning logic
```

### 5. Push and Create Pull Request

```powershell
# Push your branch
git push origin feature/your-feature-name

# Create PR via GitHub UI
# Fill in PR template
# Link related issues
# Request reviews
```

---

## Coding Standards

### PowerShell Style Guide

#### Naming Conventions

```powershell
# Functions: PascalCase with Verb-Noun
Function Get-FileLabel { }
Function Update-LabelCache { }

# Variables: camelCase
$fileCount = 0
$selectedFiles = @()

# Constants: UPPER_CASE
$SUPPORTED_EXTENSIONS = @('.docx', '.xlsx', '.pptx')

# Private/Helper: Prefix with underscore
Function _InternalHelper { }
```

#### Code Formatting

```powershell
# Braces on same line for scriptblocks
$button.Add_Click({
    # Code here
})

# Braces on new line for functions
Function Get-Example
{
    # Code here
}

# Use 4-space indentation (not tabs)
if ($condition) {
    # Indented code
}

# Align parameters in multi-line calls
Set-AIPFileLabel -Path $file `
                 -LabelId $labelId `
                 -PreserveFileDetails
```

#### Comments

```powershell
# Norwegian for user-facing messages
$statusLabel.Text = "Behandler filer..."  # Processing files

# English for code documentation
# Calculate optimal listbox height based on file count
$listBoxHeight = $rows * 16

# Section headers
# ===== LABEL CACHE MANAGEMENT =====

# Function documentation
<#
.SYNOPSIS
    Retrieves label information for a file
.PARAMETER FilePath
    Path to the file
.RETURNS
    Hashtable with label information
#>
Function Get-FileLabel {
    param([string]$FilePath)
    # Implementation
}
```

### UI Conventions

#### Control Naming

```powershell
# Descriptive names
$fileListBox       # ListBox for files
$applyButton       # Apply label button
$progressBar       # Progress indicator
$statusLabel       # Status message label

# Avoid generic names
$list1            # ❌ Not descriptive
$button2          # ❌ Not descriptive
```

#### Layout Standards

```powershell
# Standard margins
$margin = 10

# Consistent spacing
$controlSpacing = 10

# Standard fonts
$font = New-Object System.Drawing.Font("Segoe UI", 9)
$headerFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

# Standard colors
$selectedColor = [System.Drawing.Color]::FromArgb(0, 120, 212)  # Blue
$defaultColor = [System.Drawing.Color]::FromArgb(240, 240, 240) # Light gray
```

### Error Handling

```powershell
# Always use try-catch for external operations
try {
    $status = Get-AIPFileStatus -Path $file -ErrorAction Stop
    # Process status
}
catch {
    Write-Log -Message "Failed to get status: $_" -Level ERROR
    # Handle error gracefully
}

# Validate inputs
if (-not $file) {
    throw "File path cannot be null or empty"
}

# Use ErrorAction appropriately
Get-AIPFileStatus -Path $file -ErrorAction SilentlyContinue
```

### Logging Standards

```powershell
# Use structured logging
Write-Log -Message "Operation started" `
          -Level INFO `
          -Context @{FileCount=$count; Operation="LabelApplication"}

# Log levels
Write-Log -Level INFO      # Normal operations
Write-Log -Level WARNING   # Non-critical issues
Write-Log -Level ERROR     # Operation failures
Write-Log -Level CRITICAL  # Application-stopping errors
```

---

## Testing Requirements

### Test Coverage

All contributions must include appropriate tests:

#### For New Features
- ✅ Unit tests for all new functions
- ✅ Integration tests for workflows
- ✅ Manual test script if needed
- ✅ Update existing tests if behavior changes

#### For Bug Fixes
- ✅ Regression test to prevent recurrence
- ✅ Update related tests
- ✅ Verify fix with manual testing

#### For Refactoring
- ✅ All existing tests must still pass
- ✅ Add tests if coverage gaps found
- ✅ Performance tests if optimization claimed

### Running Tests

```powershell
# Run all tests
Invoke-Pester -Path .\tests\

# Run specific test suite
Invoke-Pester -Path .\tests\FileLabeler.Tests.ps1

# Run specific test
Invoke-Pester -Path .\tests\ -TestName "*Label Cache*"

# Check coverage
Invoke-Pester -Path .\tests\ -CodeCoverage .\FileLabeler.ps1
```

### Writing Tests

```powershell
# Follow Pester 3.x syntax
Describe "Function Name" {
    Context "When specific scenario" {
        It "Should expected behavior" {
            # Arrange
            $input = "test"
            
            # Act
            $result = Get-Function -Input $input
            
            # Assert
            $result | Should Be "expected"
        }
    }
}
```

See [TESTING.md](TESTING.md) for comprehensive testing guide.

---

## Pull Request Process

### Before Submitting

**Checklist:**
- [ ] Code follows style guidelines
- [ ] All tests pass
- [ ] New tests added for new features
- [ ] Documentation updated
- [ ] Commits follow commit message format
- [ ] No merge conflicts with main
- [ ] No Norwegian encoding issues (UTF-8 BOM)

### PR Description Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix (non-breaking change fixing an issue)
- [ ] New feature (non-breaking change adding functionality)
- [ ] Breaking change (fix or feature causing existing functionality to change)
- [ ] Documentation update

## Related Issues
Fixes #(issue number)
Relates to #(issue number)

## Testing
- [ ] Unit tests added/updated
- [ ] Integration tests added/updated
- [ ] Manual testing completed

## Screenshots (if applicable)
[Add screenshots for UI changes]

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] Tests pass locally
```

### Review Process

1. **Automated Checks:** GitHub Actions run tests
2. **Code Review:** Maintainers review code
3. **Feedback:** Address review comments
4. **Approval:** Get approval from maintainer(s)
5. **Merge:** Maintainer merges PR

**Review timeline:**
- Simple changes: 1-3 days
- Complex features: 1-2 weeks
- Breaking changes: Discussion required

---

## Reporting Bugs

### Before Reporting

1. **Check existing issues** - Your bug may already be reported
2. **Verify it's reproducible** - Can you reproduce it consistently?
3. **Test with latest version** - Bug may already be fixed
4. **Gather information** - Collect logs and diagnostic info

### Bug Report Template

```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce:
1. Go to '...'
2. Click on '...'
3. See error

**Expected behavior**
What you expected to happen.

**Screenshots**
If applicable, add screenshots.

**Environment:**
- Windows Version: [e.g., Windows 11 22H2]
- PowerShell Version: [e.g., 5.1.22621.963]
- FileLabeler Version: [e.g., 1.1]
- AIP Client Version: [e.g., 2.14.109.0]

**Log File**
```
[Paste relevant log entries]
```

**Additional context**
Any other information about the problem.
```

### Where to Report

- **GitHub Issues:** [https://github.com/yourusername/FileLabeler/issues](https://github.com/yourusername/FileLabeler/issues)
- **Label as:** `bug`
- **Priority labels:** `critical`, `high`, `medium`, `low`

---

## Suggesting Features

### Feature Request Template

```markdown
**Is your feature request related to a problem?**
A clear description of the problem. Ex. I'm always frustrated when [...]

**Describe the solution you'd like**
Clear description of what you want to happen.

**Describe alternatives you've considered**
Other solutions or features you've considered.

**Does this fit FileLabeler's purpose?**
FileLabeler is a mass labeling tool. Does your feature align with this purpose?

**How often would you use this feature?**
- [ ] Daily
- [ ] Weekly
- [ ] Monthly
- [ ] Rarely

**Additional context**
Screenshots, mockups, or examples.

**Would you be willing to contribute this feature?**
Yes / No / Maybe with guidance
```

### Feature Evaluation

Features are evaluated based on:

1. **Alignment** - Does it fit FileLabeler's "mass labeling" purpose?
2. **User value** - How many users would benefit?
3. **Complexity** - Implementation and maintenance cost
4. **UI impact** - Does it complicate the interface?

See [ROADMAP.md](../ROADMAP.md) for current priorities.

---

## Documentation

### When to Update Docs

Update documentation for:
- ✅ New features
- ✅ Changed behavior
- ✅ New configuration options
- ✅ Deprecated functionality
- ✅ Installation changes

### Documentation Files

```
docs/
├── USER_GUIDE.md         # End-user instructions
├── INSTALLATION.md       # Setup guide
├── CONFIGURATION.md      # Config reference
├── TROUBLESHOOTING.md    # Problem solving
├── CHANGELOG.md          # Version history
└── development/
    ├── ARCHITECTURE.md   # Technical design
    ├── TESTING.md        # Testing guide
    ├── CONTRIBUTING.md   # This file
    └── FEATURES.md       # Feature reference
```

### Documentation Standards

- Use clear, concise language
- Include code examples
- Add screenshots for UI features
- Keep examples up to date
- Cross-reference related docs

---

## Release Process

### Versioning

FileLabeler follows [Semantic Versioning](https://semver.org/):

- `MAJOR.MINOR.PATCH` (e.g., 1.2.3)
- **MAJOR:** Breaking changes
- **MINOR:** New features (backward compatible)
- **PATCH:** Bug fixes (backward compatible)

### Release Checklist

For maintainers releasing new versions:

1. **Update version numbers:**
   - README.md
   - CHANGELOG.md
   - FileLabeler.ps1 header
   - docs/development/ARCHITECTURE.md

2. **Update documentation:**
   - CHANGELOG.md with all changes
   - README.md if needed
   - Migration guide if breaking changes

3. **Run all tests:**
   ```powershell
   Invoke-Pester -Path .\tests\
   ```

4. **Create release:**
   - Tag in Git: `git tag -a v1.2.0 -m "Version 1.2.0"`
   - Push tag: `git push origin v1.2.0`
   - Create GitHub release with notes

5. **Compile EXE:**
   ```powershell
   Invoke-ps2exe -inputFile .\FileLabeler.ps1 `
                 -outputFile .\FileLabeler.exe `
                 -noConsole `
                 -version "1.2.0.0"
   ```

6. **Publish release:**
   - Upload EXE to GitHub release
   - Include sample `labels_config.json`
   - Add release notes

---

## Getting Help

### For Contributors

- **GitHub Discussions:** Ask questions, discuss ideas
- **GitHub Issues:** Report problems with contribution process
- **Documentation:** Check existing docs first

### Contact

- **Project Maintainers:** [List maintainers]
- **GitHub:** [@yourusername](https://github.com/yourusername)

---

## Recognition

Contributors are recognized in:
- GitHub contributors list
- CHANGELOG.md (for significant contributions)
- README.md acknowledgments (for major features)

---

## License

By contributing, you agree that your contributions will be licensed under the same MIT License that covers the project.

---

**Thank you for contributing to FileLabeler!**

Your efforts help make file labeling easier for everyone. 🚀

