# Changelog

All notable changes to the Outlook Move-to-Folder Add-in will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-10-20

### Added
- Initial release of Outlook Move-to-Folder Add-in
- Toolbar button integration for Outlook Desktop and Web
- Instant folder search with Fuse.js fuzzy matching
- Keyboard navigation (Arrow keys, Enter, Escape)
- Recent folders memory (last 8 used folders)
- Folder caching with 6-hour TTL
- Optimistic UI feedback with status messages
- Office notification banner integration
- Auto-close taskpane after successful move
- Microsoft Graph API integration for folder fetching and email moving
- Office.js SSO authentication with token prefetching
- Responsive UI with Tailwind CSS
- Manual folder refresh button
- Error handling and retry functionality
- Comprehensive documentation (README, INSTALLATION, QUICKSTART)
- Icon generator HTML tool
- Local development setup with http-server

### Features
- **Search**: Real-time fuzzy search through all mail folders
- **Navigation**: Full keyboard support for hands-free operation
- **Performance**: Prefetched tokens and cached folders for instant UX
- **Feedback**: Multi-layer feedback (status message → notification → auto-close)
- **Memory**: Remembers recently used folders for quick access
- **Cross-platform**: Works in Outlook Desktop (O365) and Outlook Web

### Technical Details
- No build tools required (plain HTML/JS)
- CDN-based dependencies (Tailwind CSS, Fuse.js, Office.js)
- localStorage for caching and recent folders
- Recursive folder fetching with Microsoft Graph
- Debounced search (50ms) for performance
- Configurable cache TTL and recent folder limit

## [Unreleased]

### Planned Features
- Multi-select email support
- Batch move operations using Graph `$batch` endpoint
- Mobile app support (iOS/Android)
- Undo functionality
- Folder favorites/pinning
- Search result highlighting
- Dark mode support
- Custom folder icons
- Move history tracking
- Keyboard shortcuts customization
- Drag-and-drop support (if possible in add-in context)

### Known Issues
- Single email moves only (no multi-select yet)
- Requires internet connection (no offline mode)
- Cannot show optimistic email removal from list (add-in sandbox limitation)
- O365 accounts only (no local Exchange/POP/IMAP support)

---

## Version History

- **1.0.0** (2025-10-20): Initial release with core functionality

---

## Upgrade Guide

### From Future Versions

Upgrade instructions will be added here as new versions are released.

### Breaking Changes

None yet - this is the initial release.

---

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines on proposing changes.

---

## Support

For issues and feature requests, please open an issue on GitHub or contact the maintainers.


