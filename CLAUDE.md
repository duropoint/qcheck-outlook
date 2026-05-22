# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Q Check is a Microsoft Outlook add-in (task pane) that lets users run PSC (Port State Control) quality checks on IMO numbers directly from Outlook. It is **vanilla JavaScript + Office.js** with no build system, no bundler, and no framework. Files are served directly from GitHub Pages at `https://duropoint.github.io/qcheck-outlook/`.

## Deployment

There is no build step. To deploy: push changed files to the `main` branch of the `duropoint/qcheck-outlook` GitHub repository. GitHub Pages serves them immediately (60–90 second CDN propagation).

When changing `manifest.xml`, bump the `<Version>` field (e.g. `1.0.0.0` → `1.0.0.1`) so Outlook picks up the change after re-sideloading.

## Architecture

The entire app lives in three files:

- **`taskpane.html`** — static template; all views are present in the DOM simultaneously as `<div class="view hidden">` elements
- **`taskpane.css`** — all styling; no preprocessor
- **`taskpane.js`** — all logic; a single flat script loaded after `</body>`

### Single-page view switching

`showView(el)` hides every view in `[formView, loadingView, errorView, confirmView, companyResult, vesselResult, settingsView]` then removes `hidden` from the target. There is no routing — always call `showView()` to switch screens.

### Settings persistence

Three settings are stored in `Office.context.roamingSettings` (syncs across user's devices):

| Constant | Key | Purpose |
|---|---|---|
| `SETTING_API_BASE` | `qcheck_api_base` | Backend base URL |
| `SETTING_API_KEY` | `qcheck_api_key` | Auth key for Q Check API |
| `SETTING_COMPANIES_KEY` | `qcheck_companies_key` | Auth key for Companies Search API |

`setSetting()` wraps `saveAsync` in a Promise. `getSetting()` returns a fallback if the key is absent. `getConfig()` is a convenience that returns `{ apiBase, apiKey }` for the Q Check endpoints.

**Settings gate**: `initApp()` checks all three settings at startup. If any are missing, only the settings view is shown (`firstRun = true`) and the toggle/gear button are hidden. They are unhidden after a successful `saveSettings()`.

### Two modes: Company and Vessel

The `mode` variable (`"company"` | `"vessel"`) controls which fields are visible. `setMode()` swaps active classes on the toggle buttons and toggles `.hidden` on `#companyFields` / `#vesselFields`. The IMO value is carried over when switching.

In Vessel mode, the ISM Company section (company IMO + name) is in a collapsible `#ismCollapseBody` div, toggled by `toggleIsmCollapse()`.

### API calls

Two backend endpoints (both on the same `apiBase`):

- **Q Check**: `POST /api/v1/qcheck/company` and `POST /api/v1/qcheck/vessel`
  - Header: `X-API-Key: <apiKey>`
  - 120-second `AbortController` timeout
  - Response for company: `{ global_performance, shareable_url }`
  - Response for vessel: `{ vessel_name, shareable_url, assessment: { global, age, psc, company } }`

- **Companies Search**: `GET /api/companies/search?name=…&imo=…`
  - Header: `X-API-Key: <companiesKey>`
  - Response: `{ success: true, results: [{ company_imo, company_name }, …] }`
  - Used for autocomplete and for the company-exists confirmation check

### Company-exists confirmation flow

Before running a Company Q Check, `runQCheck()` calls `searchCompanies({ imo })` and uses `isSameImo()` to check if the IMO is already in the DB. If found, it stores the `proceed` function in `pendingQCheck` and calls `showView(confirmView)`. The confirm proceed button calls `pendingQCheck()`.

`isSameImo(a, b)` uses `parseInt(a, 10) === parseInt(b, 10)` to normalize leading zeros (e.g. `"0012345" === "12345"`).

### Autocomplete

`setupAutocomplete({ inputEl, pairedEl, searchParam, dropdownEl })` wires a debounced (300 ms) search on a text input. Selecting an item fills both the typed field and `pairedEl` (bidirectional IMO ↔ Name sync). Four instances are created in `Office.onReady`:

- `#companyImo` ↔ `#companyName` (Company mode)
- `#vesselCompanyImo` ↔ `#vesselCompanyName` (ISM Company section in Vessel mode)

The `.ts-dropdown` is `position: fixed` — its coordinates are JS-calculated in `position()` using `getBoundingClientRect()` on each render. This avoids clipping by the `overflow-y: auto` on `.view-body`. Do not change it to `position: absolute` without addressing the overflow clipping.

Paste buttons use `mousedown → preventDefault()` to prevent the input losing focus before the `click` handler reads the clipboard.

### Color mapping

`colorClass(value)` maps API string values to `"green"` / `"amber"` / `"red"` CSS classes. Order matters — check `"not acceptable"` and `"very low"` before `"low"`, and `"high"` before the default fallback. Applied to company banner, vessel global banner, and each vessel factor pill via `setPill()`.

## CORS requirement

The task pane runs from `https://duropoint.github.io`. The backend (Render) must include this origin in its CORS allowed list with `Content-Type` and `X-API-Key` headers permitted. Without this, all `fetch` calls fail.

## manifest.xml notes

- The `<Id>` GUID uniquely identifies the add-in to Office. Only change it when forcing a full reinstall.
- `functionfile.html` is required by the manifest; it only calls `Office.onReady()` and does nothing else.
- The manifest registers the button for both `MessageReadCommandSurface` and `MessageComposeCommandSurface`.
