# Quick Start Guide

Get your Outlook Move-to-Folder add-in running in under 5 minutes.

## Prerequisites Check

- [ ] Outlook Desktop (new version) OR Outlook Web access
- [ ] Node.js installed (check with `node --version`)
- [ ] Internet connection

## 5-Minute Setup

### Step 1: Generate Icons (1 minute)

1. Open `src/assets/icon-generator.html` in your browser
2. Click **"Download All"**
3. Save all 4 PNG files to `src/assets/` folder
4. Verify you have:
   - `src/assets/icon-16.png`
   - `src/assets/icon-32.png`
   - `src/assets/icon-64.png`
   - `src/assets/icon-80.png`

### Step 2: Install Dependencies (30 seconds)

```powershell
npm install
```

### Step 3: Start Server (10 seconds)

```powershell
npm start
```

Server runs at `http://localhost:3000` - **Keep this terminal open!**

### Step 4: Sideload Add-in (2 minutes)

#### For Outlook Desktop (Windows):

1. Open **Outlook** (new version)
2. Click **File** → **Get Add-ins**
3. Click **My add-ins** (left sidebar)
4. Click **+ Add a custom add-in** → **Add from file...**
5. Browse to your project folder
6. Select `manifest.xml`
7. Click **OK** on security dialog
8. Click **Install**

#### For Outlook Web:

1. Go to https://outlook.office.com
2. Click **Settings** (⚙️) → **View all Outlook settings**
3. Go to **Mail** → **Customize actions** → **Get Add-ins**
4. Click **My add-ins** → **+ Add a custom add-in**
5. Select **Add from URL**
6. Enter: `http://localhost:3000/manifest.xml`
7. Click **OK** → **Install**

### Step 5: Test (30 seconds)

1. **Open any email** in Outlook
2. Look for **"Move to Folder"** button in toolbar
3. **Click it** - taskpane opens!
4. **Type** to search folders
5. **Press Enter** to move email

**Done!** 🎉

---

## Troubleshooting

### Can't see the add-in button?

- Restart Outlook completely
- Verify server is running (`http://localhost:3000` should work in browser)
- Check if manifest was installed: File → Get Add-ins → My add-ins

### "Error Loading Folders"?

- Check browser console (F12) for errors
- Ensure you're signed into Office 365
- Verify internet connection

### Icons not showing?

- Make sure you downloaded all 4 icon files
- Check they're in `src/assets/` folder
- Verify web server is running

---

## What's Next?

- Read [README.md](./README.md) for full documentation
- See [INSTALLATION.md](./INSTALLATION.md) for detailed setup
- Customize icons in `icon-generator.html`
- Edit `manifest.xml` to change add-in name/description

---

## Daily Usage

**To use the add-in:**

1. Start server: `npm start` (or keep it running)
2. Open Outlook
3. Select email → Click "Move to Folder" → Search → Enter

**To stop server:**

- Press `Ctrl+C` in terminal

**To restart after changes:**

```powershell
# Stop server (Ctrl+C), then:
npm start
```

Manifest changes require reinstalling the add-in. JavaScript/HTML changes reload automatically.

---

## Commands Reference

```powershell
# Start web server
npm start

# Start and open browser
npm run dev

# Validate manifest
npm run validate

# Install dependencies
npm install
```

---

**Need help?** See [INSTALLATION.md](./INSTALLATION.md) for detailed troubleshooting.


