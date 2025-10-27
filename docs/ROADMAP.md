# Product Roadmap

This document outlines planned features and enhancements for FileLabeler.

---

## Current Version

**Version:** 1.1  
**Status:** Production Ready  
**Release Date:** October 2025

See [CHANGELOG.md](CHANGELOG.md) for all completed features.

---

## Future Considerations

The following features have been suggested but are not currently planned for implementation. They are listed here for reference and community discussion.

### Potential Features (Not Scheduled)

#### Label Templates / Presets
**Description:** Save common label configurations as templates
- "Contracts Template" → Auto-select "Confidential"
- "Memos Template" → Auto-select "Internal"
- Save/load from config file

**Status:** Under consideration  
**Community feedback welcome**

#### Undo Last Operation
**Description:** Revert last label application
- Store previous state after application
- Reapply original labels
- Limit to last operation only

**Status:** Under consideration  
**May be added based on user demand**

#### Label Statistics / Report
**Description:** Show statistics about current file selection
- Current label distribution with percentages
- Export to report format
- Useful for auditing

**Status:** Under consideration  
**Low priority**

---

## Design Philosophy

FileLabeler is a **mass labeling tool** designed for bulk operations, not file-by-file management.

**Core principles:**
- ✅ Maintain simplicity and ease of use
- ✅ Focus on bulk operations
- ✅ Keep UI clean and uncluttered
- ✅ Preserve high performance

**We avoid features that:**
- ❌ Turn it into a general file manager
- ❌ Duplicate existing tools
- ❌ Add complexity without clear value
- ❌ Conflict with the "bulk labeling" purpose

---

## Community Feedback

**We want to hear from you!**

If you have feature suggestions:
1. Check if it aligns with FileLabeler's bulk labeling purpose
2. Open a [GitHub Discussion](https://github.com/yourusername/FileLabeler/discussions)
3. Share your use case and how the feature would help
4. Vote on existing feature requests

**Your feedback helps shape the future of FileLabeler!**

---

## Contributing

Have a feature idea you'd like to implement yourself?

See [development/CONTRIBUTING.md](development/CONTRIBUTING.md) for guidelines on contributing to FileLabeler.

---

## Changelog Reference

For completed features and version history, see [CHANGELOG.md](CHANGELOG.md).

---

**Last Updated:** October 2025  
**Status:** Feature-complete for current scope. Future development driven by community needs.

**This roadmap represents potential directions, not commitments. All features are subject to evaluation based on alignment with project goals and user value.**
