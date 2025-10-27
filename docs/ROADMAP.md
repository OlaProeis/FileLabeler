# Product Roadmap

This document outlines planned features and enhancements for FileLabeler.

---

## Current Version

**Version:** 1.1  
**Status:** Production Ready  
**Release Date:** October 2025

See [CHANGELOG.md](CHANGELOG.md) for completed features.

---

## Upcoming Releases

### v1.2 - Code Quality & Quick Wins (Q4 2025)

**Focus:** Polish, user-requested features, and foundation improvements

#### Planned Features

##### 🔧 Code Cleanup and Optimization (High Priority)
**Status:** Planned  
**Effort:** Medium (4-6 hours)  
**Description:** Comprehensive code review and optimization
- Remove redundant code blocks
- Refactor duplicate logic into reusable functions
- Optimize performance
- Improve code organization
- Enforce consistent naming conventions

**Why:** Clean foundation enables easier feature development and better maintainability

##### ⭐ Remove Individual Files from List (User Requested)
**Status:** Planned  
**Effort:** Low (2-3 hours)  
**Description:** Allow users to remove individual files from selection without clearing entire list
- Right-click context menu on file list
- "Fjern fil" / "Remove file" menu item
- Support multi-select removal
- Update file count and layout

**Use Case:** User selected 50 files but realizes 3 shouldn't be included

##### 💡 Batch Label Removal
**Status:** Planned  
**Effort:** Low (1-2 hours)  
**Description:** Remove labels from files (set to "No label")
- Add "Ingen etikett" / "No Label" option
- Or dedicated "Fjern etikett" button
- Uses `Set-AIPFileLabel -RemoveLabel`

**Use Case:** Clear labels from files during testing or reorganization

##### 💡 File Type Filtering
**Status:** Planned  
**Effort:** Low (2-3 hours)  
**Description:** Filter by file type when importing folders
- Checkboxes to select which types to import:
  - ☑ Word documents (.docx, .doc)
  - ☑ Excel workbooks (.xlsx, .xls)
  - ☑ PowerPoint presentations (.pptx, .ppt)
  - ☑ PDF files (.pdf)

**Use Case:** User wants to label only Word documents in a mixed folder

##### 💡 Search/Filter File List
**Status:** Planned  
**Effort:** Low (2-3 hours)  
**Description:** Search box to filter file list
- Real-time filter as user types
- Search filename only (not path)
- Case-insensitive
- Show "X of Y files" counter

**Use Case:** User has 100 files, wants to find all "contract" files to verify labeling

---

### v1.3 - Advanced Features (Q1 2026)

**Focus:** Power user features and workflow enhancements

#### Planned Features

##### 🔍 Per-File Label Selection (Exploratory)
**Status:** Investigation Phase  
**Effort:** High (8-12 hours)  
**Description:** Apply different labels to different files in one operation

**Current Limitation:** All files get same label (mass labeling paradigm)

**Proposed Use Case:**
```
User has 10 files:
- 3 files → "Confidential"
- 5 files → "Internal"
- 2 files → "Personal"

Currently requires 3 separate operations
Feature would allow one operation
```

**Design Challenges:**
1. Conflicts with "Mass Labeling" purpose
2. UI complexity (dropdown per file vs table view)
3. Smart Summary impact (redesign needed)

**Approaches Under Consideration:**
- **Table/Grid View** - DataGridView with dropdown per file
- **Two-Step Wizard** - Group files by label, then review
- **Smart Grouping** - Auto-group by current label, allow changes

**Phase 1 (v1.3):** Investigation only
- Research best practices
- Create UI mockups
- User feedback survey
- Document pros/cons

**Phase 2 (TBD):** Implementation (only if approved and design is intuitive)

##### 💡 Undo Last Operation
**Status:** Planned  
**Effort:** Medium (3-4 hours)  
**Description:** Revert last label application
- Store previous state after application
- "Angre siste merking" button
- Reapply original labels
- Limit to last operation only (not full history)

**Use Case:** User accidentally applied wrong label to 100 files

##### 💡 Label Statistics / Report
**Status:** Planned  
**Effort:** Medium (3-5 hours)  
**Description:** Show statistics about current file selection
- Current label distribution:
  ```
  Confidential: 15 files (45%)
  Internal: 10 files (30%)
  No label: 8 files (25%)
  
  Total: 33 files
  ```
- Optional pie chart or bar chart visualization
- Export to CSV

**Use Case:** Auditing and compliance reporting

##### 💡 Label Templates / Presets
**Status:** Planned  
**Effort:** Medium (3-4 hours)  
**Description:** Save common label configurations as templates
- "Contracts Template" → Auto-select "Confidential"
- "Memos Template" → Auto-select "Internal"
- Save/load from config file
- Quick access via Templates dropdown

**Use Case:** Frequently label specific types of documents with specific labels

---

### v2.0 - Major Enhancements (Q2-Q3 2026)

**Focus:** Modernization and broader appeal

#### Planned Features

##### 💡 Dark Mode
**Status:** Concept  
**Effort:** Medium (4-6 hours)  
**Priority:** Low

**Description:** Dark color theme option
- Theme toggle in Settings
- Dark color palette definition
- Apply to all controls
- Save preference in `app_config.json`

**Rationale:** Modern UX trend, easier on eyes in low-light environments

##### 💡 Multi-Language Support
**Status:** Concept  
**Effort:** High (8-10 hours)  
**Priority:** Low

**Description:** Support multiple UI languages
- Load UI strings from language files
- Settings → Select language
- Restart to apply
- Supported languages: Norwegian (default), English, potentially others

**Current:** All UI text hardcoded in Norwegian

**Challenge:** Significant refactoring required

---

## Priority Matrix

| Feature | Priority | Effort | User Value | Complexity | Version |
|---------|----------|--------|------------|------------|---------|
| Code cleanup | High | Medium | High (Quality) | Low | v1.2 |
| Remove individual files | Medium | Low | High | Low | v1.2 |
| Batch label removal | Medium | Low | Medium | Low | v1.2 |
| File type filtering | Low | Low | Medium | Low | v1.2 |
| Search/filter list | Low | Low | Medium | Low | v1.2 |
| Undo operation | Medium | Medium | High | Medium | v1.3 |
| Label statistics | Low | Medium | Low | Low | v1.3 |
| Label templates | Low | Medium | Low | Medium | v1.3 |
| Per-file labels | Medium | High | High | High | v1.3/v2.0 |
| Dark mode | Very Low | Medium | Low | Medium | v2.0 |
| Multi-language | Very Low | High | Low | High | v2.0 |

---

## Feature Evaluation Criteria

New features are evaluated based on:

1. **User Value**: How much does it improve user experience?
2. **Complexity**: How difficult is implementation?
3. **Alignment**: Does it fit the "mass labeling" purpose?
4. **Maintenance**: What's the long-term maintenance cost?
5. **UI Impact**: Does it complicate the interface?

### Design Philosophy

FileLabeler is a **mass labeling tool** - designed for bulk operations, not file-by-file management.

**New features should:**
- ✅ Maintain simplicity and ease of use
- ✅ Not conflict with "bulk" paradigm
- ✅ Add clear value for typical use cases
- ✅ Not overcomplicate the UI

**Avoid features that:**
- ❌ Turn it into a general file manager
- ❌ Duplicate existing tools (e.g., Windows Explorer)
- ❌ Add complexity without proportional value
- ❌ Require significant UI redesign

---

## Community Feedback

**We want to hear from you!**

Help prioritize features by:
1. Voting on feature requests in [GitHub Discussions](https://github.com/yourusername/FileLabeler/discussions)
2. Reporting pain points in [Issues](https://github.com/yourusername/FileLabeler/issues)
3. Sharing use cases and workflows
4. Participating in beta testing for new versions

**Your feedback shapes the roadmap!**

---

## Tentative Release Schedule

| Version | Focus | Target |
|---------|-------|--------|
| v1.2 | Code Quality & Quick Wins | Q4 2025 |
| v1.3 | Advanced Features | Q1 2026 |
| v2.0 | Modernization | Q2-Q3 2026 |

**Note:** Schedule is tentative and may adjust based on:
- User feedback and priorities
- Development capacity
- Technical challenges
- Security updates or critical fixes

---

## Declined Features

Features we've considered but decided against:

### Email Integration
**Reason:** Out of scope. FileLabeler focuses on labeling, not distribution.

### Office Add-in
**Reason:** Microsoft Purview client already provides this. FileLabeler's value is in bulk operations.

### Label Policy Management
**Reason:** This is an admin function. FileLabeler is for end-users applying labels to files.

### File Content Search
**Reason:** This turns FileLabeler into a file manager. Use Windows Search or other tools.

---

## Contributing Ideas

Have a feature idea? We'd love to hear it!

1. **Check existing roadmap** to see if it's already planned
2. **Open a Discussion** on GitHub Discussions to propose your idea
3. **Provide context:**
   - What problem does it solve?
   - How often would you use it?
   - Does it fit the "mass labeling" purpose?
   - Any UI mockups or examples?

See [development/CONTRIBUTING.md](development/CONTRIBUTING.md) for contribution guidelines.

---

## Changelog Reference

For completed features, see [CHANGELOG.md](CHANGELOG.md).

---

**Last Updated:** October 2025  
**Next Review:** After v1.1 release and user feedback

**This roadmap is a living document and subject to change based on user needs, technical feasibility, and development priorities.**

