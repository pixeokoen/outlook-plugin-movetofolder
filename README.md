# 📁 Outlook Move-to-Folder Add-in

> **🚀 New here?** Start with **[START_HERE.md](./START_HERE.md)** for a 5-minute setup guide!

A modern, keyboard-driven Outlook add-in that brings back the **fast, reliable "Move to Folder"** functionality to Outlook Desktop (O365) and Outlook Web.

Built with Microsoft Graph API, this add-in provides an instant search interface with optimistic UI feedback to make email management feel instantaneous, even though operations happen in the cloud.

---

## ✨ Features

### Core Functionality
- **🎯 Toolbar Integration** - One-click access from Outlook ribbon/toolbar
- **⚡ Instant Search** - Fuzzy search through all mail folders with real-time filtering
- **⌨️ Keyboard-Driven** - Navigate and move emails without touching the mouse
- **📌 Recent Folders** - Quick access to your most-used folders
- **💾 Smart Caching** - Folders cached locally for instant loading (6-hour TTL)
- **🔄 Auto-Refresh** - Manual refresh button to update folder list on demand

### UX Optimizations

This add-in implements several techniques to mask Microsoft Graph API latency and create a perceived-instant experience:

1. **Prefetched Authentication** - Auth token fetched on taskpane open
2. **Warm Graph Connection** - Initial Graph call made during initialization
3. **Optimistic UI Feedback** - Immediate checkmark and status message (50-150ms)
4. **Auto-Close Taskpane** - Closes automatically after move confirmation (300ms)
5. **Office Notification Banner** - Native Outlook notification for confirmation
6. **Folder Caching** - Instant folder list from localStorage
7. **Recent Folders Memory** - Last 8 used folders shown at top

**Result:** Users perceive instant feedback even though Graph API calls take 500-1500ms.

---

## 🎨 User Interface

```
┌─────────────────────────────────────┐
│  Type to search folders...      ↻   │  ← Auto-focused search
│  ↑↓ Navigate  Enter Move  Esc Close │  ← Keyboard hints
├─────────────────────────────────────┤
│  RECENT FOLDERS                     │
│  📁 Clients                         │  ← Last used folders
│  📁 Archive                         │
├─────────────────────────────────────┤
│  ALL FOLDERS                        │
│  📁 Archive / 2024                  │  ← Full folder paths
│  📁 Clients / Project Alpha         │  ← Fuzzy searchable
│  📁 Inbox / Newsletters             │
│  ...                                │
└─────────────────────────────────────┘
```

---

## 🚀 Quick Start

### Installation

See **[INSTALLATION.md](./INSTALLATION.md)** for complete setup instructions.

**TL;DR:**
```powershell
# 1. Start web server
npm install -g http-server
http-server -p 3000 --cors -c-1

# 2. Add icons to src/assets/ (see src/assets/ICONS_README.md)

# 3. Sideload in Outlook
# - Desktop: File → Get Add-ins → Add from file → select manifest.xml
# - Web: Settings → Get Add-ins → Add from URL → http://localhost:3000/manifest.xml
```

### Usage

1. **Open any email** in Outlook
2. **Click "Move to Folder"** button in toolbar
3. **Type** to search for destination folder
4. **Press ↓/↑** to navigate results (or click)
5. **Press Enter** to move email
6. **Done!** Taskpane closes, notification confirms move

---

## 🏗️ Architecture

### Technology Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| **Manifest** | Office Add-in XML | Defines add-in, permissions, toolbar integration |
| **Frontend** | HTML + Tailwind CSS | Taskpane UI with utility-first styling |
| **Logic** | Vanilla JavaScript | Core functionality (no framework dependencies) |
| **Search** | Fuse.js | Fuzzy string matching for folder search |
| **API** | Microsoft Graph | Email and folder operations |
| **Auth** | Office.js SSO | Single Sign-On via `Office.auth.getAccessToken()` |
| **Storage** | localStorage | Folder cache and recent folders |

### File Structure

```
outlook-plugin-movetofolder/
├── manifest.xml              # Office Add-in manifest
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Taskpane UI
│   │   ├── taskpane.js      # Core logic (auth, Graph, search, UI)
│   │   └── taskpane.css     # [Optional] Custom styles
│   └── assets/
│       ├── icon-16.png      # Toolbar icon (16x16)
│       ├── icon-32.png      # Toolbar icon (32x32)
│       ├── icon-64.png      # Store icon (64x64)
│       ├── icon-80.png      # High-res icon (80x80)
│       └── ICONS_README.md  # Icon instructions
├── INSTALLATION.md           # Sideloading guide
├── README.md                # This file
└── kickoff_briefing.md      # Original project spec
```

### Data Flow

```
User clicks toolbar button
         ↓
Taskpane opens (taskpane.html)
         ↓
Office.js initializes
         ↓
Prefetch auth token (warm connection)
         ↓
Load folders (cache or Graph API)
         ↓
Initialize Fuse.js fuzzy search
         ↓
Display UI (auto-focus search)
         ↓
User types → filter folders
         ↓
User presses Enter
         ↓
Show "Moving..." (50ms)
         ↓
Call Graph API: POST /me/messages/{id}/move
         ↓
Show "Moved" (150ms) + Office notification
         ↓
Auto-close taskpane (300ms)
         ↓
Update recent folders in localStorage
```

---

## 🔑 Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **↓** / **↑** | Navigate folder results |
| **Enter** | Move email to selected folder |
| **Esc** | Close taskpane |
| **Type** | Filter folders (fuzzy search) |

---

## ⚙️ Configuration

Edit these constants in `src/taskpane/taskpane.js`:

```javascript
const CONFIG = {
    CACHE_KEY: 'mailFoldersCache',      // localStorage key for folder cache
    RECENT_KEY: 'recentFolders',        // localStorage key for recent folders
    CACHE_TTL: 1000 * 60 * 60 * 6,     // Cache TTL: 6 hours
    RECENT_LIMIT: 8,                    // Number of recent folders to remember
    DEBOUNCE_DELAY: 50,                 // Search debounce (ms)
    FUSE_THRESHOLD: 0.3                 // Fuzzy match threshold (0=exact, 1=match anything)
};
```

---

## 📊 Microsoft Graph API Usage

### Endpoints Used

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/me/mailFolders` | GET | Fetch all mail folders |
| `/me/mailFolders/{id}/childFolders` | GET | Fetch child folders recursively |
| `/me/messages/{id}/move` | POST | Move email to destination folder |
| `/me` | GET | Warm up Graph connection (prefetch) |

### Permissions Required

- `Mail.ReadWrite` - Read and write access to user's mail
- `Mail.ReadWrite.Shared` - Access to shared mailboxes (optional)

Configured in `manifest.xml`:
```xml
<Permissions>ReadWriteMailbox</Permissions>
<Scopes>
  <Scope>Mail.ReadWrite</Scope>
  <Scope>Mail.ReadWrite.Shared</Scope>
</Scopes>
```

---

## 🔒 Security & Privacy

- **Authentication:** Uses Office.js Single Sign-On (SSO) - no separate login required
- **Permissions:** Only requests mail read/write permissions (no access to files, calendar, etc.)
- **Data Storage:** Only folder metadata and recent folder IDs stored locally
- **No External Services:** All data flows between Outlook ↔ Microsoft Graph ↔ Add-in
- **No Telemetry:** No analytics or tracking implemented

---

## 🚧 Known Limitations

### Add-in Sandbox Constraints

Due to Office Add-in architecture, the following are **not possible**:

1. **Cannot manipulate Outlook's email list directly** - Add-in runs in isolated iframe
2. **Cannot show "optimistic removal"** of email from list - Email disappears when Outlook refreshes naturally
3. **No local/instant moves** - All operations go through Microsoft Graph (cloud-based)
4. **Limited notification options** - Can only use Office notification banners (yellow bar)

### Current Implementation

- **Single email moves only** - Multi-select not yet implemented (planned)
- **Desktop & Web only** - Mobile support requires additional manifest configuration
- **Requires internet** - No offline mode (Graph API dependency)
- **O365 accounts only** - Does not work with local Exchange or POP/IMAP accounts

---

## 🛠️ Development

### Local Development Setup

1. **Clone repository:**
   ```bash
   git clone https://github.com/yourusername/outlook-move-to-folder.git
   cd outlook-move-to-folder
   ```

2. **Add icon assets:**
   - See `src/assets/ICONS_README.md`
   - Place icon-16.png, icon-32.png, icon-64.png, icon-80.png in `src/assets/`

3. **Start web server:**
   ```bash
   npm install -g http-server
   http-server -p 3000 --cors -c-1
   ```

4. **Sideload in Outlook:**
   - See [INSTALLATION.md](./INSTALLATION.md) for detailed steps

5. **Debug:**
   - **Outlook Desktop:** Press `F12` to open DevTools
   - **Outlook Web:** Press `F12` in browser, navigate to taskpane iframe
   - Check Console for errors, Network tab for Graph API calls

### Making Changes

- **UI changes:** Edit `src/taskpane/taskpane.html`
- **Styling:** Modify Tailwind classes or add custom CSS
- **Logic changes:** Edit `src/taskpane/taskpane.js`
- **Manifest changes:** Edit `manifest.xml` (requires add-in reinstall)

**Hot reload:** Most changes reload automatically if using Live Server. For manifest changes, you must remove and re-add the add-in.

### Testing Checklist

- [ ] Taskpane opens from toolbar button
- [ ] Search input auto-focuses on open
- [ ] Folders load (check for ~100+ folders)
- [ ] Cache works (refresh page, should load instantly)
- [ ] Fuzzy search returns relevant results
- [ ] Recent folders appear after first move
- [ ] Keyboard navigation (↑↓ Enter Esc)
- [ ] Email moves successfully
- [ ] Office notification banner appears
- [ ] Taskpane auto-closes after move
- [ ] Error states display properly
- [ ] Works in Outlook Desktop
- [ ] Works in Outlook Web

---

## 🐛 Troubleshooting

See **[INSTALLATION.md](./INSTALLATION.md)** for detailed troubleshooting steps.

**Common issues:**
- **Add-in button missing:** Check manifest is installed, restart Outlook
- **Authentication fails:** Ensure using O365 account, check Graph permissions
- **Folders won't load:** Check console for errors, verify internet connection
- **Icons don't show:** Verify files exist, check web server is serving assets

**Debug tips:**
```javascript
// Open browser console (F12) in taskpane and run:
localStorage.clear();           // Clear cache
location.reload();              // Reload taskpane
console.log(state.folders);     // Inspect folder data
```

---

## 🚀 Production Deployment

### Hosting Requirements

For production use, you need:
- **HTTPS web server** (required by Office Add-ins in production)
- **Public URL** for manifest and files
- **SSL certificate** (Let's Encrypt is fine)

### Deployment Steps

1. **Host files on HTTPS server:**
   - Upload `src/` directory and `manifest.xml`
   - Ensure all files accessible via HTTPS

2. **Update manifest.xml:**
   - Replace all `localhost:3000` URLs with production URLs
   - Update `<Id>` with unique GUID
   - Update `<ProviderName>` and `<SupportUrl>`

3. **Deploy via Microsoft 365 Admin Center:**
   - Go to https://admin.microsoft.com
   - Navigate to **Settings** → **Integrated apps** → **Add-ins**
   - Click **Upload custom apps** → Upload manifest.xml
   - Configure deployment (who can access)
   - Deploy organization-wide or to specific users

4. **OR distribute manifest manually:**
   - Share manifest.xml file with users
   - Users sideload via **File** → **Get Add-ins** → **Add from file**

---

## 📝 Contributing

Contributions welcome! Areas for improvement:

- [ ] **Multi-select support** - Move multiple emails at once
- [ ] **Batch operations** - Use Graph `$batch` endpoint
- [ ] **Mobile support** - Optimize for Outlook mobile apps
- [ ] **Drag-and-drop** - Drag emails onto folder list (if possible in add-in context)
- [ ] **Folder favorites** - Pin specific folders to top
- [ ] **Custom folder icons** - Visual indicators for folder types
- [ ] **Search highlighting** - Highlight matching text in results
- [ ] **Undo functionality** - Move email back to original folder
- [ ] **Analytics** - Track most-used folders for optimization
- [ ] **Dark mode** - Respect Outlook's theme

---

## 📄 License

MIT License - see LICENSE file for details.

---

## 🙏 Acknowledgments

- Built for users frustrated by Microsoft's unreliable built-in "Move to Folder" feature
- Inspired by the need for keyboard-driven email management
- Powered by Microsoft Graph API and Office.js

---

## 📚 Additional Resources

- **Office Add-ins Documentation:** https://learn.microsoft.com/en-us/office/dev/add-ins/
- **Microsoft Graph Mail API:** https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview
- **Office.js API Reference:** https://learn.microsoft.com/en-us/javascript/api/overview/outlook
- **Manifest Schema:** https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/
- **Fuse.js Documentation:** https://fusejs.io/
- **Tailwind CSS:** https://tailwindcss.com/

---

## 💬 Support

Having issues? Check the [INSTALLATION.md](./INSTALLATION.md) troubleshooting section or open an issue on GitHub.

---

**Made with ❤️ for efficient email management**

