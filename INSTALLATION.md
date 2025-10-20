# Installation Guide - Outlook Move-to-Folder Add-in

This guide provides efficient, step-by-step instructions for sideloading the add-in in Outlook Desktop (new O365 version) and Outlook Web.

---

## Prerequisites

- **Outlook Desktop:** New Outlook for Windows (Microsoft 365)  
  *Not the classic Outlook - must be the modern web-based version*
- **OR Outlook Web:** Access to Outlook.com or Office 365 web mail
- **Node.js & http-server** (for local hosting) OR **any local web server**
- **Icons:** Icon files in `src/assets/` (see `src/assets/ICONS_README.md`)

---

## Step 1: Prepare Icon Assets

Before proceeding, ensure you have the required icon files:

```bash
src/assets/icon-16.png
src/assets/icon-32.png
src/assets/icon-64.png
src/assets/icon-80.png
```

See `src/assets/ICONS_README.md` for instructions on creating or obtaining these icons.

---

## Step 2: Start Local Web Server

The add-in files must be served over HTTPS (or HTTP for local testing).

### Option A: Using http-server (Recommended)

```powershell
# Install http-server globally (one-time)
npm install -g http-server

# Navigate to project directory
cd C:\_Dev\outlook-plugin-movetofolder

# Start server with CORS enabled
http-server -p 3000 --cors -c-1

# Server will run at: http://localhost:3000
```

### Option B: Using Python

```powershell
# Python 3
python -m http.server 3000

# Server will run at: http://localhost:3000
```

### Option C: Using VS Code Live Server

1. Install "Live Server" extension in VS Code
2. Right-click `manifest.xml` → "Open with Live Server"
3. Note the port (usually 5500)
4. Update manifest.xml URLs to match the port

**Important:** Keep the server running while using the add-in.

---

## Step 3: Update Manifest URLs

Edit `manifest.xml` and replace all instances of `localhost:3000` with your actual server address and port.

**Find and replace:**
- FROM: `https://localhost:3000`
- TO: `http://localhost:3000` (or your server URL)

**For production deployment:** Replace with your actual HTTPS hosting URL.

---

## Step 4A: Sideload in Outlook Desktop (Windows)

### Method 1: Via File Share (Recommended)

1. **Create a shared folder:**
   ```powershell
   # Create a network share or use a local folder
   mkdir C:\OutlookAddIns
   copy manifest.xml C:\OutlookAddIns\
   ```

2. **Add the manifest location in Outlook:**
   - Open **Outlook Desktop** (new version)
   - Go to **File** → **Get Add-ins** (or **Manage Add-ins**)
   - Click **My add-ins** (left sidebar)
   - Under **Custom add-ins**, click **+ Add a custom add-in** → **Add from file...**
   - Browse to `C:\OutlookAddIns\manifest.xml`
   - Click **OK** to confirm the security dialog
   - Click **Install**

### Method 2: Via Registry (Advanced)

**Warning:** Editing registry requires administrator privileges.

1. Open Registry Editor (`Win + R` → type `regedit`)

2. Navigate to:
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer
   ```

3. If the `Developer` key doesn't exist, create it:
   - Right-click `WebExt` → **New** → **Key** → Name it `Developer`

4. Create a new String Value:
   - Right-click `Developer` → **New** → **String Value**
   - Name: `ManifestPath1`
   - Value: `C:\_Dev\outlook-plugin-movetofolder\manifest.xml`

5. Restart Outlook

6. Verify installation:
   - Open an email
   - Look for **"Move to Folder"** button in the ribbon

### Method 3: Microsoft 365 Admin Center Deployment

For organization-wide deployment (requires admin access):

1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Settings** → **Integrated apps** → **Add-ins**
3. Click **Upload custom apps**
4. Upload the `manifest.xml` file
5. Configure deployment settings (who can access it)
6. Deploy

---

## Step 4B: Sideload in Outlook Web

1. **Open Outlook Web:**
   - Go to https://outlook.office.com or https://outlook.live.com
   - Sign in to your account

2. **Access Add-ins:**
   - Click the **Settings gear icon** (⚙️) in top-right
   - Click **View all Outlook settings** at the bottom
   - Go to **Mail** → **Customize actions** → **Get Add-ins** (or just search for "add-ins")

3. **Sideload the manifest:**
   - Click **My add-ins** (left sidebar)
   - Under **Custom add-ins**, click **+ Add a custom add-in** → **Add from URL...**
   - Enter: `http://localhost:3000/manifest.xml`
   - Click **OK** to confirm
   - Click **Install**

   **Alternative (if URL doesn't work):**
   - Select **Add from file...**
   - Upload your `manifest.xml` file
   - Click **Install**

4. **Verify installation:**
   - Open any email
   - Look for **"Move to Folder"** button in the toolbar/ribbon
   - Click it to open the taskpane

---

## Step 5: Test the Add-in

1. **Open an email** in Outlook
2. **Click the "Move to Folder" button** in the toolbar
3. **Taskpane should open** with:
   - Auto-focused search input
   - List of all mail folders
   - Recent folders (if any)
4. **Type to search** for a folder
5. **Press ↓ or ↑** to navigate
6. **Press Enter** to move the email
7. **Verify:** Email should move and taskpane closes

---

## Troubleshooting

### Add-in button doesn't appear

**Solutions:**
- Ensure the manifest is correctly installed (check Add-ins list)
- Restart Outlook completely
- Verify web server is running (`http://localhost:3000` should be accessible)
- Check manifest.xml for syntax errors
- Clear Office cache:
  ```powershell
  # Close Outlook first
  Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\"
  ```

### Taskpane shows "Error Loading Folders"

**Solutions:**
- Open browser DevTools (F12 in new Outlook Desktop)
- Check Console for errors
- Verify you're signed into Office 365
- Check Microsoft Graph API permissions
- Try re-authenticating: Sign out and back into Outlook

### Authentication fails

**Solutions:**
- Ensure you're using a Microsoft 365 account (not a local Outlook account)
- Verify the `WebApplicationInfo` section in manifest.xml
- Check if your organization blocks custom add-ins
- Try using Outlook Web instead of Desktop

### Icons don't load

**Solutions:**
- Verify icon files exist in `src/assets/`
- Check file names match exactly (case-sensitive)
- Ensure web server is serving the assets folder
- Test icon URLs directly in browser: `http://localhost:3000/src/assets/icon-16.png`

### CORS errors in console

**Solutions:**
- Restart web server with CORS enabled:
  ```powershell
  http-server -p 3000 --cors -c-1
  ```
- Or add CORS headers to your web server configuration

### Folders not loading

**Solutions:**
- Check browser console for Graph API errors
- Verify internet connection
- Clear localStorage:
  ```javascript
  // Open browser console in taskpane (F12)
  localStorage.clear();
  location.reload();
  ```
- Click the refresh button (↻) in the taskpane

---

## Uninstalling the Add-in

### Outlook Desktop

**Method 1: Via UI**
1. Go to **File** → **Get Add-ins** → **My add-ins**
2. Find "Move to Folder" add-in
3. Click the **three dots (...)** → **Remove**

**Method 2: Registry (if used)**
1. Open Registry Editor
2. Navigate to: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`
3. Delete the `ManifestPath1` string value
4. Restart Outlook

### Outlook Web

1. Settings (⚙️) → **View all Outlook settings**
2. **Mail** → **Customize actions** → **Get Add-ins**
3. **My add-ins** → Find "Move to Folder"
4. Click **three dots (...)** → **Remove**

---

## Next Steps

Once installed and tested:

1. **Customize:** Edit `manifest.xml` to change add-in name, description, icons
2. **Deploy:** Host files on a public HTTPS server for permanent deployment
3. **Share:** Distribute manifest.xml to team members or deploy via admin center
4. **Enhance:** Modify `taskpane.js` to add custom features or adjust behavior

---

## Production Deployment Checklist

Before deploying to production:

- [ ] Replace all `localhost:3000` URLs with production HTTPS URLs
- [ ] Host all files on a reliable HTTPS server (required for production)
- [ ] Generate unique GUID for manifest `<Id>` element
- [ ] Update `WebApplicationInfo` with proper Azure AD app registration (if needed)
- [ ] Test in both Outlook Desktop and Web
- [ ] Test with multiple email accounts
- [ ] Verify Graph API permissions are properly configured
- [ ] Test move operation with various folder types
- [ ] Update `ProviderName` and `SupportUrl` in manifest
- [ ] Create proper icons (not placeholders)
- [ ] Test error scenarios (network failure, auth failure, etc.)

---

## Support & Resources

- **Office Add-ins Documentation:** https://learn.microsoft.com/en-us/office/dev/add-ins/
- **Microsoft Graph API:** https://learn.microsoft.com/en-us/graph/
- **Sideloading Guide:** https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins
- **Manifest Reference:** https://learn.microsoft.com/en-us/office/dev/add-ins/reference/manifest/

---

**Installation complete!** Your Outlook Move-to-Folder add-in should now be ready to use.


