# Contributing to Outlook Move-to-Folder Add-in

Thank you for considering contributing to this project! This document provides guidelines and instructions for contributing.

## Code of Conduct

Be respectful, constructive, and professional in all interactions.

## How Can I Contribute?

### Reporting Bugs

Before submitting a bug report:
- Check if the issue has already been reported
- Verify you're using the latest version
- Test in both Outlook Desktop and Web (if possible)

**Bug Report Template:**
```
**Describe the bug:**
A clear description of what the bug is.

**To Reproduce:**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '...'
3. See error

**Expected behavior:**
What you expected to happen.

**Screenshots/Console Errors:**
If applicable, add screenshots or console errors (F12).

**Environment:**
- Outlook version: [Desktop/Web]
- Browser (if Web): [e.g. Chrome 120]
- OS: [e.g. Windows 11]

**Additional context:**
Any other relevant information.
```

### Suggesting Features

We welcome feature suggestions! Before submitting:
- Check if it's already planned (see [CHANGELOG.md](./CHANGELOG.md) Unreleased section)
- Consider if it fits the project's scope (keyboard-driven, instant UX)

**Feature Request Template:**
```
**Feature Description:**
Clear description of the feature.

**Use Case:**
Why is this feature useful? What problem does it solve?

**Proposed Implementation:**
(Optional) How you think it could be implemented.

**Alternatives Considered:**
Other ways you've considered solving this problem.
```

## Development Setup

### Prerequisites

- Node.js 14+ installed
- Git installed
- Outlook Desktop (O365) or Outlook Web access
- Code editor (VS Code recommended)

### Getting Started

1. **Fork the repository**
   ```bash
   # Click "Fork" on GitHub
   ```

2. **Clone your fork**
   ```bash
   git clone https://github.com/your-username/outlook-move-to-folder.git
   cd outlook-move-to-folder
   ```

3. **Install dependencies**
   ```bash
   npm install
   ```

4. **Generate icons** (if not present)
   - Open `src/assets/icon-generator.html` in browser
   - Download all icons to `src/assets/`

5. **Start development server**
   ```bash
   npm start
   ```

6. **Sideload in Outlook**
   - See [INSTALLATION.md](./INSTALLATION.md) for detailed steps

### Making Changes

1. **Create a branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-bug-fix
   ```

2. **Make your changes**
   - Edit files in `src/taskpane/` or `manifest.xml`
   - Test thoroughly in Outlook

3. **Test your changes**
   - Verify in Outlook Desktop (if possible)
   - Verify in Outlook Web
   - Check browser console (F12) for errors
   - Test keyboard navigation
   - Test with slow network (throttle in DevTools)

4. **Commit your changes**
   ```bash
   git add .
   git commit -m "feat: add awesome feature"
   ```

   **Commit message format:**
   - `feat:` new feature
   - `fix:` bug fix
   - `docs:` documentation changes
   - `style:` formatting, no code change
   - `refactor:` code refactoring
   - `test:` adding tests
   - `chore:` maintenance tasks

5. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```

6. **Create Pull Request**
   - Go to GitHub and create a Pull Request
   - Fill out the PR template
   - Wait for review

## Code Style Guidelines

### JavaScript

- Use ES6+ syntax (const/let, arrow functions, async/await)
- Use descriptive variable names
- Add comments for complex logic
- Keep functions focused and single-purpose
- Handle errors gracefully

**Example:**
```javascript
// Good
async function fetchFoldersFromGraph() {
    try {
        const token = await getAccessToken();
        const response = await fetch(/* ... */);
        return await response.json();
    } catch (error) {
        console.error('Failed to fetch folders:', error);
        throw error;
    }
}

// Avoid
async function getData() {
    var t = await getAccessToken(); // var, unclear name
    return fetch(/* ... */).then(r => r.json()); // inconsistent style
}
```

### HTML

- Use semantic HTML elements
- Include ARIA attributes for accessibility
- Keep structure clean and indented
- Use Tailwind utility classes for styling

### CSS

- Prefer Tailwind utility classes
- Use custom CSS only when necessary
- Keep custom styles in `<style>` tag or separate file
- Use CSS variables for theme colors

### Manifest XML

- Maintain proper XML formatting
- Validate before committing: `npm run validate`
- Update version numbers appropriately

## Testing Checklist

Before submitting a PR, verify:

- [ ] Code runs without console errors
- [ ] Taskpane opens and displays correctly
- [ ] Folder search works with various queries
- [ ] Keyboard navigation works (↑↓ Enter Esc)
- [ ] Email moves successfully
- [ ] Recent folders update correctly
- [ ] Cache refresh works
- [ ] Error states display properly
- [ ] Works in Outlook Desktop (if accessible)
- [ ] Works in Outlook Web
- [ ] No breaking changes (or documented)

## Project Structure

```
outlook-plugin-movetofolder/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # UI structure
│   │   └── taskpane.js      # Core logic
│   └── assets/
│       ├── icon-*.png       # Icon files
│       ├── icon-generator.html
│       └── ICONS_README.md
├── manifest.xml             # Add-in manifest
├── package.json
├── README.md
├── INSTALLATION.md
├── QUICKSTART.md
├── CHANGELOG.md
├── CONTRIBUTING.md          # This file
└── LICENSE
```

## Key Files to Know

- **`src/taskpane/taskpane.js`** - Main logic (auth, Graph API, search, UI)
- **`src/taskpane/taskpane.html`** - UI layout and structure
- **`manifest.xml`** - Add-in configuration and permissions
- **`package.json`** - Dependencies and scripts

## Architecture Overview

```
User clicks button
    ↓
Office.js initializes taskpane
    ↓
Prefetch auth token (Office.auth.getAccessToken)
    ↓
Load folders (cache or Graph API)
    ↓
Initialize Fuse.js for search
    ↓
User searches → filter folders → render results
    ↓
User selects → move email (Graph API)
    ↓
Show feedback → close taskpane
```

## Common Development Tasks

### Adding a new feature

1. Update `taskpane.js` with logic
2. Update `taskpane.html` if UI changes needed
3. Test thoroughly
4. Update documentation (README, CHANGELOG)
5. Submit PR

### Changing UI styling

1. Modify Tailwind classes in `taskpane.html`
2. Or add custom CSS in `<style>` tag
3. Test responsiveness
4. Verify in both Desktop and Web

### Modifying Graph API calls

1. Update functions in `taskpane.js`
2. Check Microsoft Graph documentation
3. Verify permissions in `manifest.xml`
4. Test with real data
5. Handle errors appropriately

### Updating manifest

1. Edit `manifest.xml`
2. Validate: `npm run validate`
3. Remove and re-add add-in in Outlook
4. Test new behavior

## Resources

- [Office Add-ins Docs](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/)
- [Office.js Reference](https://learn.microsoft.com/en-us/javascript/api/overview/outlook)
- [Tailwind CSS Docs](https://tailwindcss.com/docs)
- [Fuse.js Docs](https://fusejs.io/)

## Questions?

If you have questions:
1. Check existing documentation (README, INSTALLATION, etc.)
2. Search existing issues on GitHub
3. Open a new issue with the "question" label

## Recognition

Contributors will be recognized in:
- GitHub contributors page
- Release notes (for significant contributions)
- README acknowledgments section

Thank you for contributing! 🎉


