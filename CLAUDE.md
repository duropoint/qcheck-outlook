# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

The **EUROMAR Toolkit** is a collection of internal browser-based tools shared across the team. A single hosted UI (served from GitHub Pages) is embedded by two shells:

- **Outlook Add-in** — `manifest.xml` sideloaded into Outlook; loads the task pane via Office.js
- **Chrome Extension** — `euromar-toolkit-chrome/`; opens a side panel embedding the same hosted URL

Pushing to `main` updates both shells automatically — no reinstall needed for UI or logic changes. The shells only need updating when manifest-level changes are required (new permissions, new API domains, new content scripts).

## Repository layout

```
taskpane.html          Hosted UI — single HTML shell, all views in DOM simultaneously
taskpane.css           All styling; no preprocessor
taskpane.js            All logic; flat script loaded after </body>
env.js                 Environment abstraction (Outlook / Chrome extension / browser)
manifest.xml           Outlook Add-in manifest
functionfile.html      Required by manifest.xml; calls Office.onReady() only
INSTALL.md             End-user install instructions
scripts/               Hosted tool scripts fetched at run-time by the extension shell
  sfbr-seafarers.js
  sfbr-zoho.js
  sfbr-styles.css
euromar-toolkit-chrome/   Chrome Extension thin shell (v2.0.0+)
  manifest.json
  background.js
  popup.js / popup.html / popup.css
  relay.js             Generic ISOLATED-world relay (never changes)
  icons/
```

## Deployment

No build step. Push to `main` → GitHub Pages serves within 60–90 seconds.

When changing `manifest.xml`, bump `<Version>` (e.g. `1.0.0.0` → `1.0.0.1`) so Outlook picks up the change after re-sideloading.

When changing `euromar-toolkit-chrome/manifest.json`, bump `"version"` and reload the unpacked extension at `chrome://extensions`.

## Environment abstraction — `env.js`

`env.js` must be loaded before `taskpane.js`. It exposes a single `Env` object and abstracts three runtime contexts:

| `Env.env` | Detection | Settings storage | Ready trigger |
|---|---|---|---|
| `"extension"` | `chrome.storage.sync` available | `chrome.storage.sync` (cached) | `chrome.storage.sync.get` completes |
| `"outlook"` | `window === window.top` AND `Office.context.requirements != null` | `Office.context.roamingSettings` | `Office.onReady()` |
| `"browser"` | fallback | `localStorage` | DOMContentLoaded |

**Critical rules:**
- `window === window.top` guard in `_detectEnv()` prevents office.js from being mistaken for an Outlook host when the page is loaded inside an iframe (e.g. the Chrome extension side panel). Without this, office.js waits for an Office handshake that never comes and the panel hangs blank.
- `taskpane.html` conditionally loads office.js only when `window.self === window.top`. This prevents office.js from navigating the iframe to `about:blank` after its host-detection timeout expires.
- All five settings must be present for `initApp()` to show the main view; missing any shows the settings view instead (`firstRun = true`).

## Hosted UI architecture (`taskpane.html` / `taskpane.js`)

### Single-page view switching

`showView(el)` hides every view in the list then removes `hidden` from the target. No routing — always call `showView()` to switch screens.

Views: `mainView`, `maritimeView`, `formView`, `loadingView`, `errorView`, `confirmView`, `companyResult`, `vesselResult`, `settingsView`, `zammadSearchView`, `zammadReportsView`.

### Settings

Five settings stored per-environment (see `env.js` table above):

| Constant | Key | Purpose |
|---|---|---|
| `SETTING_API_BASE` | `qcheck_api_base` | Backend base URL |
| `SETTING_API_KEY` | `qcheck_api_key` | Auth key for Q Check API |
| `SETTING_COMPANIES_KEY` | `qcheck_companies_key` | Auth key for Companies Search API |
| `SETTING_ZAMMAD_TOKEN` | `qcheck_zammad_token` | Zammad personal token |
| `SETTING_DASHBOARD_KEY` | `qcheck_dashboard_key` | Dashboard API key |

### Tools

**Q Check (Maritime → Q Check)**
- Two modes: Company and Vessel, toggled by `setMode()`
- Before a Company Q Check, `runQCheck()` calls `searchCompanies({ imo })` and checks `isSameImo()` — if already in DB, shows `confirmView`
- API: `POST /api/v1/qcheck/company` and `POST /api/v1/qcheck/vessel` with `X-API-Key` header, 120 s timeout

**Vessel Search (Maritime → Vessel Search)**
- Debounced search via `GET /api/ships/search?name=…`
- Results list → detail card
- In Chrome extension context (`Env.env === "extension"` or `?context=extension` URL param): shows **"… into Zammad Case"** button
- Button routes through `window.parent.postMessage({ type: "zvl_fill", vessel })` → `popup.js` → `background.js` → `chrome.scripting.executeScript` injecting `injectFillVessel` directly into the active Zammad tab (no content script needed)

**Zammad Reports (Maritime → Zammad Reports)**
- Generates a PDF report via `POST https://zammad-dashboard.onrender.com/api/v1/report`
- Filters: date range, vessel IMO, open/closed states

### Autocomplete

`setupAutocomplete({ inputEl, pairedEl, searchParam, dropdownEl })` — debounced 300 ms, bidirectional IMO ↔ Name sync. The `.ts-dropdown` is `position: fixed` with JS-calculated coordinates to avoid clipping from `overflow-y: auto` on `.view-body`. Do not change to `position: absolute`.

### Color mapping

`colorClass(value)` → `"green"` / `"amber"` / `"red"`. Check `"not acceptable"` and `"very low"` before `"low"`, and `"high"` before the default fallback.

## Chrome Extension shell (`euromar-toolkit-chrome/`) — v2.0.0 generic shell

### Architecture

From v2.0.0 the extension is a **fixed thin shell** — it exposes 8 generic operations and never needs updating for new tools. All tool logic (Zoho bridge, future tools) is hosted under `/scripts/` on GitHub Pages and fetched at run-time.

```
euromar-toolkit-chrome/
  manifest.json          permissions: scripting, tabs, notifications, downloads, contextMenus, storage, sidePanel
  background.js          relay receiver + same 8 ops for relay-originated requests
  popup.js               iframe ↔ chrome.* bridge with 8 ops
  popup.html / popup.css unchanged
  relay.js               GENERIC ISOLATED-world relay (never changes)
  icons/

scripts/                 (HOSTED, auto-deploys via GitHub Pages)
  sfbr-seafarers.js      Zoho bridge — Seafarers Panel side
  sfbr-zoho.js           Zoho bridge — Zoho form side
  sfbr-styles.css        Styles for both
  ...future tool scripts
```

### The 8 generic shell ops

The hosted toolkit (taskpane.js) calls `callExtOp(type, payload)` which posts a message to popup.js and returns a Promise. Op handlers live in popup.js (for iframe-originated requests) and in background.js (for relay-originated requests — duplicated handlers because the two contexts cannot share modules).

| Op | Payload | Purpose |
|---|---|---|
| `exec-on-tab` | `{ tabId, scriptUrl, world }` | Fetch JS from `/scripts/<scriptUrl>` and run in tab. Auto-injects relay.js. Default world: MAIN. |
| `css-on-tab` | `{ tabId, cssUrl }` | Fetch CSS and inject |
| `open-tab` | `{ url, ops?, delay? }` | Open URL, wait for complete, run ops sequentially against the new tab |
| `get-tab-info` | `{ tabId }` (or "active") | Return id/url/title |
| `close-tab` | `{ tabId }` | Close a tab |
| `notify` | `{ title, message, iconUrl? }` | Chrome OS notification |
| `badge` | `{ text, color? }` | Set extension icon badge |
| `download` | `{ url, filename?, saveAs? }` | Download file via chrome.downloads |

### Hosted scripts ↔ shell signalling

Hosted scripts run in MAIN world and have no `chrome.*` access. To call back to the shell (e.g. open another tab, badge, download), they emit a relay-signal element:

```js
const el = document.createElement('span');
el.className = '__euromar_relay__';
el.setAttribute('data-euromar-action', 'open-tab');
el.setAttribute('data-euromar-payload', JSON.stringify({ url: '...', ops: [...] }));
document.documentElement.appendChild(el);
```

`relay.js` (always auto-injected in ISOLATED world by `exec-on-tab`) watches for these elements via `MutationObserver`, parses the payload, and forwards to `background.js` via `chrome.runtime.sendMessage({ source: 'relay', action, payload })`. background.js dispatches through its own copy of the op handlers.

### MAIN-world CSP caveat

`exec-on-tab` injects fetched code via `new Function(code).call(window)` running in MAIN world. This is subject to the page's `script-src 'unsafe-eval'` CSP. Normal web apps (Seafarers Panel, Zoho forms, Zammad) allow it. Hardened pages (e.g. banking sites) may not — those pages would need a per-page workaround.

### Backwards-compat handlers

`popup.js` and `background.js` keep two legacy paths so existing features don't break:
- `toolkit-api` — generic API POST to the QCheck backend
- `zvl_fill` — Vessel Search → Zammad ticket fill (still uses inline `injectFillVessel` function)

### `?context=extension`

`popup.js` always appends `?context=extension` to the iframe URL. `taskpane.js` reads this to detect the Chrome extension context, since `Env.env` is `"browser"` inside the hosted iframe.

### When to reload / reinstall the extension

After the v2.0.0 refactor: **the extension folder should not need changes for new tools**. New tool = push a JS file to `scripts/` on `main`.

**Reload only** (click ↺ on `chrome://extensions`, settings preserved) — required only if a file inside `euromar-toolkit-chrome/` itself changes (shell logic, relay, popup HTML/CSS).

**Reinstall** (zip / "Load unpacked" again) — only when `manifest.json` permissions change. This should be a rare event.

> **⚠️ Always ask the user before implementing anything that requires modifying the extension folder or `manifest.json`.**
> The whole point of the v2.0.0 shell is to make this unnecessary for new tools. If you find yourself needing to add a file to `euromar-toolkit-chrome/`, stop — there is almost certainly a way to express the feature using the 8 existing ops with a hosted script.

## CORS

`https://duropoint.github.io` must be in the backend's CORS allowed origins with `Content-Type` and `X-API-Key` headers permitted.

## manifest.xml notes

- `<Id>` GUID uniquely identifies the add-in to Office. Change only to force a full reinstall.
- `functionfile.html` is required by the manifest; it only calls `Office.onReady()`.
- The manifest registers the button for both `MessageReadCommandSurface` and `MessageComposeCommandSurface`.
