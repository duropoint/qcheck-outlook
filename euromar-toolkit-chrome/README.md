# EUROMAR Toolkit — Chrome Extension Shell

A thin Chrome extension shell that embeds the EUROMAR Toolkit hosted UI inside a native browser popup window. The extension contains no tool logic — everything is loaded from the hosted URL via an iframe. When the hosted toolkit is updated (new tools, bug fixes, UI changes), all installed copies of the extension reflect the changes automatically on the next popup open, with no reinstallation required.

## Install (Developer Mode)

1. Open `chrome://extensions`
2. Enable **Developer mode** (toggle, top-right)
3. Click **Load unpacked**
4. Select this `euromar-toolkit-chrome/` folder
5. The **EUROMAR Toolkit** icon appears in the Chrome toolbar

## How Updates Work

Push changes to the hosted toolkit repository → they appear automatically the next time the popup is opened. The extension files in this folder rarely need to change.

## When You DO Need to Reinstall

Only when modifying the extension shell itself:

- **`manifest.json`** — new permissions, updated context menu entries, new host permissions
- **`background.js`** — new API bridge behaviour, new window-management logic

UI changes, new tools, bug fixes, and API endpoint changes all live in the hosted toolkit and require no extension update.

## Configure API Access

1. Click the EUROMAR Toolkit icon in the Chrome toolbar
2. A popup window opens with the hosted toolkit UI
3. Use the toolkit's built-in settings panel to enter your API key and (optionally) a custom API base URL
4. Settings are stored in `chrome.storage.local` — persisted per device, never synced to the cloud

## Adding New Tools

Tools are added to the hosted toolkit repository, not to this extension. The extension is a stable container. The only time this extension folder needs a change is if a new tool requires a new API domain — in that case, add the domain to `host_permissions` in `manifest.json` and reinstall.

## Changing the Hosted URL

The hosted toolkit URL is a single constant defined at the top of both `background.js` and `popup.js`:

```js
const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";
```

Update both occurrences and reload the unpacked extension in `chrome://extensions`.
