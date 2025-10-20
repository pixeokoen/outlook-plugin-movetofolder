# 📄 Project Briefing: Outlook Move-to-Folder Add-in (O365 Desktop/Web)

## 🎯 Goal
Develop a **custom Outlook Add-in** that replaces the broken built-in "Move to Folder" functionality. The add-in must integrate natively into the **Outlook for Windows (O365 desktop)** toolbar and also work in **Outlook Web**. It will allow users to quickly move selected emails to any folder via an instant search interface.

The add-in should provide a fast, responsive user experience, even though it relies on **Microsoft Graph** (which means actual moves aren’t local or instant). The UX must feel instant through clever UI behavior (optimistic updates, caching, feedback, etc.).

---

## 🧠 Context
- The default Outlook “Move to Folder” feature is unreliable.
- We want a **cross-platform**, **cloud-based** solution — not a local COM/VSTO add-in.
- The add-in will be built as a **modern Office Web Add-in**, using **Microsoft Graph**.
- It must be deployable via **Office 365 Admin Center** or sideloadable in Outlook Desktop (new version) and Web.

---

## 🧩 Core Features
### 1. Toolbar Integration
- Adds a button in the Outlook **toolbar/ribbon** (under Message Read/Compose surfaces).
- Clicking the button opens a **dropdown (taskpane or dialog)** with a focused search bar.

### 2. Search UI
- Search input is auto-focused on open.
- Typing instantly filters through all available folders (client-side fuzzy match).
- Keyboard navigation:
  - **Arrow Up/Down:** Navigate folder results.
  - **Enter:** Confirm move.
  - **Escape:** Cancel.

### 3. Folder Management
- Fetch all folders using Microsoft Graph:
  - `GET /me/mailFolders?$top=1000&$expand=childFolders($levels=max)`
- Cache folders locally using `localStorage` or `IndexedDB`.
- Cache validity: ~6–12 hours.
- Add optional manual refresh (↻ button).

### 4. Message Movement
- Use Graph API:
  - `POST /me/messages/{id}/move` with `{ "destinationId": folder.id }`
- Support moving multiple selected messages when possible.
- Optimistic UI update:
  - Fade out selected email immediately.
  - Show toast: “✅ Moved to [FolderName]”.
  - If Graph fails → revert or show warning toast.

### 5. Performance + UX Tricks
- **Instant UX illusion:** Optimistic removal and toast.
- **Caching:** Folder tree preloaded and refreshed silently.
- **Batch Graph requests:** Use `$batch` endpoint for multiple moves.
- **No blocking spinners.** Use subtle toasts.
- **Recent folders:** Store last N used folders locally and show them on top of results.

---

## 🧰 Tech Stack
- **Manifest:** Office Add-in manifest (XML).
- **Frontend:** HTML + JS (or React + Tailwind, optional Fuse.js for fuzzy search).
- **APIs:** Microsoft Graph (`Mail.ReadWrite`, `Mail.ReadWrite.Shared`).
- **Storage:** `localStorage` for caching folders and recent folders.
- **Authentication:** `OfficeRuntime.auth.getAccessToken()`.

---

## 🔐 Permissions
Manifest XML:
```xml
<Permissions>ReadWriteMailbox</Permissions>
<Authorization>
  <Scopes>
    <Scope>Mail.ReadWrite</Scope>
    <Scope>Mail.ReadWrite.Shared</Scope>
  </Scopes>
</Authorization>
```

---

## 🧱 Architecture Overview
### Files
- `manifest.xml` → defines add-in, toolbar button, taskpane source.
- `index.html` → taskpane UI with search + list.
- `main.js` → handles Graph calls, folder caching, keyboard nav, move action.
- (Optional) `style.css` → simple Tailwind or minimal CSS.

### Workflow
1. User selects email → clicks **Move** button in toolbar.
2. Taskpane opens → input auto-focused.
3. Cached folders load instantly.
4. User types → results filter instantly.
5. Press **Enter** → email moved via Graph.
6. Toast confirms success or failure.

---

## ⚙️ Folder Cache Logic
```js
const CACHE_KEY = 'mailFoldersCache';
const CACHE_TTL = 1000 * 60 * 60 * 6; // 6 hours

async function loadFolders(force = false) {
  const cached = JSON.parse(localStorage.getItem(CACHE_KEY) || '{}');
  const isValid = cached.timestamp && Date.now() - cached.timestamp < CACHE_TTL;

  if (!force && isValid) return cached.folders;

  const folders = await fetchFoldersFromGraph();
  localStorage.setItem(CACHE_KEY, JSON.stringify({ folders, timestamp: Date.now() }));
  return folders;
}
```

---

## ✅ Deliverables
1. `manifest.xml` — properly configured for Outlook toolbar button.
2. `index.html` — minimal search + results + recent UI.
3. `main.js` — all logic (Graph calls, caching, UI control).
4. (Optional) Tailwind setup for styling.
5. Working sideloadable version tested in Outlook Web + Outlook Desktop (O365).

---

## 🧭 Next Steps
1. Implement minimal working prototype (manifest + HTML + JS).
2. Integrate folder caching and optimistic UI feedback.
3. Add keyboard navigation and recent-folder memory.
4. Polish with proper error handling + UX feedback.

---

## 🔎 Reference APIs
- Microsoft Graph Mail API Docs: https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview
- Office.js Reference: https://learn.microsoft.com/en-us/javascript/api/overview/outlook
- Add-in manifest schema: https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest

---

## 🧭 Goal Summary (for Cursor)
Build a **modern, cross-platform Outlook Add-in** with a toolbar-integrated, keyboard-driven folder search UI that moves selected emails through Microsoft Graph with cached folder data and optimistic UI feedback.

The UX must feel instantaneous, even though actual moves are asynchronous.

