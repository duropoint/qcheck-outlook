# Q Check — Outlook Add-in

Outlook task pane that runs PSC Q Check on any IMO number directly inside Outlook (Mac, Windows, Web).

---

## How it works

1. Read or compose any email in Outlook
2. Click the **Q Check** button in the ribbon
3. Task pane opens on the right side
4. Highlight an IMO in the email → click **"Use selected text"** → it fills the IMO field automatically
5. (Or type the IMO manually)
6. Toggle Company/Vessel, click **Run Q Check**, get the assessment + shareable link

Same logic as the Chrome extension, just docked inside Outlook.

---

## Installation — 3 stages

### STAGE 1 — Host the files on GitHub Pages (one-time, ~10 min)

The add-in files must be served over HTTPS. GitHub Pages is free and easy.

1. Go to https://github.com — sign in with your `duropoint` account.
2. Click **New** (green button, top left) to create a new repository.
3. Name it exactly: `qcheck-outlook`
4. Set it to **Public** (required for free GitHub Pages).
5. Check **"Add a README file"**, then click **Create repository**.
6. In the new repo, click **Add file → Upload files**.
7. Drag in **every file from this folder** (manifest.xml, taskpane.html, taskpane.css, taskpane.js, functionfile.html, and the entire `assets` folder with all 5 PNGs).
   - Tip: select all files in the unzipped folder and drag them in at once. GitHub will preserve the `assets/` subfolder structure.
8. Scroll to the bottom → click **Commit changes**.
9. In the repo, click **Settings** (top tabs) → **Pages** (left sidebar).
10. Under "Build and deployment", set:
    - Source: **Deploy from a branch**
    - Branch: **main** / **/ (root)**
    - Click **Save**.
11. Wait 1–2 minutes. Refresh the Pages settings tab. You'll see:
    > **Your site is live at https://duropoint.github.io/qcheck-outlook/**
12. Open that URL in your browser. You should see a directory listing (or a blank page — either is fine). To verify it's actually working, visit:
    `https://duropoint.github.io/qcheck-outlook/taskpane.html`
    You should see the Q Check form UI.

✅ **Hosting done.** The manifest already points to these URLs.

### STAGE 2 — Sideload the add-in into Outlook for Mac (one-time, ~3 min)

1. Open **Outlook for Mac**.
2. Top menu: **Tools** → **Get Add-ins**.
3. The Add-ins dialog opens. Bottom-left corner: click **My add-ins**.
4. Scroll down to **Custom add-ins** → click **+ Add a custom add-in** → **Add from file…**
5. Select the **`manifest.xml`** file from this folder.
6. A security warning appears — click **Install**.
7. Close the dialog. Restart Outlook.

✅ **Add-in installed.** You'll now see a **Q Check** button in the ribbon when reading or composing any email.

### STAGE 3 — Configure your API key (one-time per device)

1. Open any email in Outlook.
2. Click the **Q Check** ribbon button → task pane opens on the right.
3. Click the **⚙ gear icon** (top-right of the pane).
4. Paste your API key (the `Q_CHECK_API_KEY` from Render).
5. Click **Test** to verify, then **Save**.
6. Click **Back** to return to the form.

✅ **Ready to use.**

---

## Daily usage

- Read an email with an IMO → highlight the number → click **Q Check** in the ribbon → click **Use selected text** → toggle mode → **Run Q Check**.
- Or just open the pane and type/paste the IMO directly.
- Results show the same color-coded badges and shareable link as the Chrome extension and the PSC Platform itself.

---

## Important: enable CORS on your API

The Outlook task pane runs from origin `https://duropoint.github.io`. Your Render API must allow this origin, or every call will fail with "Failed to fetch" (same as we saw with the Chrome popup window).

Add this header to your API responses (in your Render Flask/FastAPI server):

**Python Flask** — install `flask-cors`, then:
```python
from flask_cors import CORS
CORS(app, origins=[
    "https://duropoint.github.io",
    "chrome-extension://*"
], allow_headers=["Content-Type", "X-API-Key"])
```

**Python FastAPI** —
```python
from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://duropoint.github.io"],
    allow_methods=["POST"],
    allow_headers=["Content-Type", "X-API-Key"],
)
```

Redeploy on Render. After that, the add-in will work.

> Reminder: the **Chrome extension** doesn't need CORS — it uses the service worker. Only the Outlook add-in does, because Office task panes can't bypass CORS.

---

## Updating the add-in later

When you change the code:

1. Re-upload the changed files to GitHub (drag-and-drop overwrite).
2. **Bump the version number** in `manifest.xml` (e.g. `1.0.0.0` → `1.0.0.1`).
3. Re-upload `manifest.xml` to GitHub.
4. In Outlook: Tools → Get Add-ins → My add-ins → find Q Check → remove it → re-add from the updated manifest URL or file.

For most changes (HTML/CSS/JS), you don't need to re-sideload — just clear the Outlook cache and reload. But bumping the manifest version is safest.

---

## Centralized deployment for EUROMAR (later)

Once tested, hand the **`manifest.xml`** file (with hosted URLs) to whoever administers EUROMAR's Microsoft 365 tenant. They'll go to:

> **Microsoft 365 Admin Center** → Settings → Integrated apps → Upload custom apps → Office Add-in → upload `manifest.xml` → select users/groups → deploy

The add-in will then appear in every selected user's Outlook (Mac, Windows, web, mobile) automatically within 6–24 hours, with zero action from them.

---

## Troubleshooting

- **"Q Check" button doesn't appear**: Restart Outlook completely (Cmd+Q, reopen). Confirm in Tools → Get Add-ins → My add-ins that Q Check is listed and enabled.
- **Task pane shows blank / error loading**: Open `https://duropoint.github.io/qcheck-outlook/taskpane.html` in your browser. If it doesn't load there, GitHub Pages isn't set up yet (wait 2 more minutes or recheck Stage 1).
- **"Failed to fetch" when running a check**: CORS issue. See the CORS section above.
- **"API key not set"**: Click the ⚙ icon in the pane and configure it.
- **"Use selected text" does nothing**: Make sure you actually highlighted text in the email body before clicking. Subject line text doesn't count in Mac Outlook.

---

## Files

```
qcheck-outlook/
├── manifest.xml           Add-in definition (this is what you sideload)
├── taskpane.html          Task pane UI
├── taskpane.css           Styling
├── taskpane.js            Logic + Office.js integration
├── functionfile.html      Required helper file
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
└── INSTALL.md             (this file)
```
