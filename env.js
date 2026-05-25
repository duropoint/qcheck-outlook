// Environment abstraction — works in Outlook, Chrome Extension, and plain browser.
//
// Detection order matters:
//   1. Extension first — uses chrome.storage.sync (present in extension context
//      regardless of whether office.js loaded or was blocked by CSP)
//   2. Outlook — Office.context.requirements is only set inside a real Office host
//   3. Browser fallback
//
// For Chrome Extension packaging: office.js CDN may be blocked by default CSP.
// That is harmless — extension is detected via chrome.storage, not via Office.
// If you want to suppress the CSP console warning, add the CDN to
// content_security_policy.extension_pages in your manifest.json.

const Env = (() => {
  // ---- Environment detection ----

  const _isExtension = typeof chrome !== "undefined"
    && typeof chrome.storage !== "undefined"
    && typeof chrome.storage.sync !== "undefined";

  function _detectEnv() {
    if (_isExtension) return "extension";
    if (
      typeof Office !== "undefined" &&
      Office.context != null &&
      Office.context.requirements != null
    ) return "outlook";
    return "browser";
  }

  const env      = _detectEnv();
  const isOffice = env === "outlook";

  // Settings cache for extension mode.
  // Populated synchronously before the ready() callback fires, so getSetting()
  // can remain synchronous everywhere in taskpane.js.
  let _cache = {};

  // ---- getSetting ----

  function getSetting(key, fallback) {
    let v;
    if (env === "outlook") {
      v = Office.context.roamingSettings ? Office.context.roamingSettings.get(key) : undefined;
    } else if (env === "extension") {
      v = _cache[key];
    } else {
      v = localStorage.getItem(key);
    }
    return (v == null) ? (fallback !== undefined ? fallback : null) : v;
  }

  // ---- setSetting (always returns a Promise) ----

  function setSetting(key, value) {
    if (env === "outlook") {
      Office.context.roamingSettings.set(key, value);
      return new Promise((resolve, reject) => {
        Office.context.roamingSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(result.error);
        });
      });
    }

    if (env === "extension") {
      _cache[key] = value;
      return new Promise((resolve) => chrome.storage.sync.set({ [key]: value }, resolve));
    }

    // browser — localStorage is synchronous
    if (value == null || value === "") {
      localStorage.removeItem(key);
    } else {
      localStorage.setItem(key, value);
    }
    return Promise.resolve();
  }

  // ---- openUrl ----
  // Outlook: Office browser window API
  // Extension popup: chrome.tabs.create (window.open stays inside the popup)
  // Browser: standard window.open

  function openUrl(url) {
    if (env === "outlook") {
      try {
        Office.context.ui.openBrowserWindow(url);
      } catch (_) {
        window.open(url, "_blank");
      }
    } else if (env === "extension") {
      chrome.tabs.create({ url });
    } else {
      window.open(url, "_blank");
    }
  }

  // ---- getUserProfile ----
  // Outlook: reads from Office mailbox (automatic, no user input needed)
  // Extension / Browser: reads email from stored setting; name is always blank

  function getUserProfile() {
    if (env === "outlook") {
      const p = Office.context.mailbox.userProfile;
      return {
        displayName:  p.displayName  || "",
        emailAddress: p.emailAddress || ""
      };
    }
    return {
      displayName:  "",
      emailAddress: getSetting("qcheck_user_email", "")
    };
  }

  // ---- ready ----
  // Extension: pre-loads ALL chrome.storage.sync entries into _cache before
  //   firing the callback, so getSetting() is synchronous afterwards.
  // Outlook: delegates to Office.onReady().
  // Browser: fires when DOM is ready (handles script-at-end-of-body case).

  function _callWhenReady(cb) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", cb);
    } else {
      cb();
    }
  }

  function ready(cb) {
    if (env === "outlook") {
      Office.onReady(cb);
      return;
    }
    if (env === "extension") {
      chrome.storage.sync.get(null, (items) => {
        _cache = items || {};
        _callWhenReady(cb);
      });
      return;
    }
    _callWhenReady(cb);
  }

  return { env, isOffice, ready, getSetting, setSetting, openUrl, getUserProfile };
})();
