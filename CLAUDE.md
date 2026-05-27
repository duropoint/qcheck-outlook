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
euromar-toolkit-chrome/   Chrome Extension shell
  manifest.json
  background.js
  popup.js
  popup.html
  popup.css
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
- Button routes through `window.parent.postMessage({ type: "zvl_fill", vessel })` → `popup.js` → `background.js` → `content.js` on the active Zammad tab

**Zammad Reports (Maritime → Zammad Reports)**
- Generates a PDF report via `POST https://zammad-dashboard.onrender.com/api/v1/report`
- Filters: date range, vessel IMO, open/closed states

### Autocomplete

`setupAutocomplete({ inputEl, pairedEl, searchParam, dropdownEl })` — debounced 300 ms, bidirectional IMO ↔ Name sync. The `.ts-dropdown` is `position: fixed` with JS-calculated coordinates to avoid clipping from `overflow-y: auto` on `.view-body`. Do not change to `position: absolute`.

### Color mapping

`colorClass(value)` → `"green"` / `"amber"` / `"red"`. Check `"not acceptable"` and `"very low"` before `"low"`, and `"high"` before the default fallback.

## Chrome Extension shell (`euromar-toolkit-chrome/`)

### Architecture

The extension is a **thin shell** — no tool logic lives here. All tools load from the hosted URL via an iframe in the side panel.

```
TOOLKIT_URL constant (top of background.js and popup.js)
     ↓ iframe src
taskpane.html (GitHub Pages)
     ↓ postMessage
popup.js (extension page)
     ↓ chrome.runtime.sendMessage
background.js (service worker)
```

### Message bridge

Two bridge types, both initiated by the iframe via `window.parent.postMessage`:

| `type` | Direction | Purpose |
|---|---|---|
| `toolkit-api` | iframe → popup → background → fetch → back | Generic API POST to the backend |
| `zvl_fill` | iframe → popup → background → `executeScript` | Fill vessel fields in active Zammad tab |

Responses are posted back to the iframe at `TOOLKIT_ORIGIN` with a matching `requestId`.

### `?context=extension`

`popup.js` always appends `?context=extension` to the iframe URL. `taskpane.js` reads this parameter to show the **"… into Zammad Case"** button in the Vessel Search detail view, since `Env.env` is `"browser"` inside the iframe (the `chrome` object is not available to web-origin iframes).

### When to reinstall the extension

Only when changing `euromar-toolkit-chrome/manifest.json`:
- New permissions
- New context menu entries

`host_permissions` is set to `<all_urls>` so new tools and new domains **never** require a reinstall. DOM injection on any site is handled via `chrome.scripting.executeScript` directly from `background.js` — no `content_scripts` registration needed.

UI changes, new tools, API logic changes, new domain interactions → update hosted files on `main` only.

## CORS

`https://duropoint.github.io` must be in the backend's CORS allowed origins with `Content-Type` and `X-API-Key` headers permitted.

## manifest.xml notes

- `<Id>` GUID uniquely identifies the add-in to Office. Change only to force a full reinstall.
- `functionfile.html` is required by the manifest; it only calls `Office.onReady()`.
- The manifest registers the button for both `MessageReadCommandSurface` and `MessageComposeCommandSurface`.
