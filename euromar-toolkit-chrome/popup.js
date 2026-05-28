// popup.js — message bridge between the hosted toolkit iframe and the
// extension's Chrome APIs. Handles all 8 generic operations directly so
// future tools can be added without touching this file.

const TOOLKIT_URL    = "https://duropoint.github.io/qcheck-outlook/taskpane.html";
const TOOLKIT_ORIGIN = new URL(TOOLKIT_URL).origin;
const SCRIPTS_BASE   = "https://duropoint.github.io/qcheck-outlook/scripts/";

const iframe   = document.getElementById("contentFrame");
const BASE_SRC = `${TOOLKIT_URL}?context=extension`;
iframe.src = BASE_SRC;

chrome.storage.session.get("pendingSelection")
  .then(({ pendingSelection }) => {
    if (!pendingSelection) return;
    chrome.storage.session.remove("pendingSelection");
    iframe.src = `${BASE_SRC}&selection=${encodeURIComponent(pendingSelection)}`;
  })
  .catch(() => {});

chrome.storage.session.onChanged.addListener((changes) => {
  const newValue = changes.pendingSelection?.newValue;
  if (!newValue) return;
  chrome.storage.session.remove("pendingSelection");
  iframe.src = `${BASE_SRC}&selection=${encodeURIComponent(newValue)}`;
});

// ── Helpers ─────────────────────────────────────────────────────────────────

function reply(payload) {
  try { iframe.contentWindow.postMessage(payload, "*"); }
  catch (err) { console.error("[Toolkit] reply failed:", err); }
}

async function findActiveTab() {
  let tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tabs.length) tabs = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  return tabs[0] || null;
}

async function resolveTabId(tabId) {
  if (tabId && tabId !== "active") return tabId;
  const tab = await findActiveTab();
  if (!tab) throw new Error("No active tab found.");
  return tab.id;
}

async function fetchText(url) {
  const fullUrl = url.startsWith("http") ? url : SCRIPTS_BASE + url;
  const resp = await fetch(fullUrl, { cache: "no-cache" });
  if (!resp.ok) throw new Error(`Fetch ${fullUrl} failed: ${resp.status}`);
  return await resp.text();
}

function waitForTabComplete(tabId) {
  return new Promise((resolve) => {
    const timeout = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      resolve();
    }, 25000);
    const listener = (id, changeInfo) => {
      if (id === tabId && changeInfo.status === "complete") {
        clearTimeout(timeout);
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    };
    chrome.tabs.onUpdated.addListener(listener);
  });
}

// ── Generic operations ──────────────────────────────────────────────────────

async function opExecOnTab({ tabId, scriptUrl, world }) {
  const id = await resolveTabId(tabId);
  const code = await fetchText(scriptUrl);

  // Always inject the generic relay (ISOLATED) so the hosted script can call back.
  // Idempotent — guarded by window.__euromarRelayInstalled inside relay.js.
  await chrome.scripting.executeScript({
    target: { tabId: id },
    files:  ["relay.js"],
    world:  "ISOLATED"
  });

  // Inject the hosted code in the requested world via new Function().
  // For MAIN world this is subject to page CSP (script-src 'unsafe-eval').
  // For ISOLATED world the extension's own CSP allows it.
  const targetWorld = world === "ISOLATED" ? "ISOLATED" : "MAIN";
  await chrome.scripting.executeScript({
    target: { tabId: id },
    world:  targetWorld,
    func: (code, url) => {
      try { new Function(code).call(window); }
      catch (e) { console.error("[EUROMAR] script error", url, e); throw e; }
    },
    args: [code, scriptUrl]
  });

  return { ok: true, tabId: id };
}

async function opCssOnTab({ tabId, cssUrl }) {
  const id = await resolveTabId(tabId);
  const css = await fetchText(cssUrl);
  await chrome.scripting.insertCSS({ target: { tabId: id }, css });
  return { ok: true, tabId: id };
}

async function opOpenTab({ url, ops, delay }) {
  const tab = await chrome.tabs.create({ url });
  await waitForTabComplete(tab.id);
  if (delay) await new Promise(r => setTimeout(r, delay));
  if (Array.isArray(ops)) {
    for (const op of ops) {
      await dispatchOp(op.type, { ...(op.payload || {}), tabId: tab.id });
    }
  }
  return { ok: true, tabId: tab.id };
}

async function opGetTabInfo({ tabId }) {
  if (!tabId || tabId === "active") {
    const tab = await findActiveTab();
    if (!tab) throw new Error("No active tab found.");
    return { ok: true, id: tab.id, url: tab.url || "", title: tab.title || "" };
  }
  const tab = await chrome.tabs.get(tabId);
  return { ok: true, id: tab.id, url: tab.url || "", title: tab.title || "" };
}

async function opCloseTab({ tabId }) {
  await chrome.tabs.remove(tabId);
  return { ok: true };
}

async function opNotify({ title, message, iconUrl }) {
  await chrome.notifications.create({
    type:    "basic",
    iconUrl: iconUrl || "icons/icon-128.png",
    title:   title   || "EUROMAR Toolkit",
    message: message || ""
  });
  return { ok: true };
}

async function opBadge({ text, color }) {
  await chrome.action.setBadgeText({ text: text == null ? "" : String(text) });
  if (color) await chrome.action.setBadgeBackgroundColor({ color });
  return { ok: true };
}

async function opDownload({ url, filename, saveAs }) {
  const downloadId = await chrome.downloads.download({
    url,
    filename: filename || undefined,
    saveAs:   !!saveAs
  });
  return { ok: true, downloadId };
}

async function dispatchOp(type, payload) {
  switch (type) {
    case "exec-on-tab":  return opExecOnTab(payload);
    case "css-on-tab":   return opCssOnTab(payload);
    case "open-tab":     return opOpenTab(payload);
    case "get-tab-info": return opGetTabInfo(payload);
    case "close-tab":    return opCloseTab(payload);
    case "notify":       return opNotify(payload);
    case "badge":        return opBadge(payload);
    case "download":     return opDownload(payload);
    default: throw new Error(`Unknown operation: ${type}`);
  }
}

const GENERIC_OPS = new Set([
  "exec-on-tab", "css-on-tab", "open-tab", "get-tab-info",
  "close-tab", "notify", "badge", "download"
]);

// ── Message bridge ──────────────────────────────────────────────────────────

window.addEventListener("message", async (event) => {
  if (event.origin !== TOOLKIT_ORIGIN) return;
  const msg = event.data;
  if (!msg || typeof msg.type !== "string") return;

  console.log("[Toolkit/popup]", msg.type, msg.requestId || "");

  // Diagnostic ping
  if (msg.type === "ping" || msg.type === "sfbr_ping") {
    reply({
      type:      msg.type === "ping" ? "ping-response" : "sfbr_ping_response",
      requestId: msg.requestId,
      ok:        true,
      version:   chrome.runtime.getManifest().version,
      tabInfo:   await safeTabInfo()
    });
    return;
  }

  // Generic shell operations
  if (GENERIC_OPS.has(msg.type)) {
    try {
      const result = await dispatchOp(msg.type, msg.payload || {});
      reply({ type: `${msg.type}-response`, requestId: msg.requestId, ...result });
    } catch (err) {
      console.error(`[Toolkit] ${msg.type} failed:`, err);
      reply({ type: `${msg.type}-response`, requestId: msg.requestId, ok: false, error: err.message });
    }
    return;
  }

  // ── Backwards-compat handlers (existing toolkit features) ────────────────

  if (msg.type === "close") { window.close(); return; }
  if (msg.type === "resize") return; // reserved

  if (msg.type === "toolkit-api") {
    const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
    try {
      const response = await chrome.runtime.sendMessage({
        action: "toolkit-api", endpoint: msg.endpoint, body: msg.body
      });
      reply({ type: "toolkit-api-response", requestId, ...response });
    } catch (err) {
      reply({ type: "toolkit-api-response", requestId, ok: false, error: err.message });
    }
    return;
  }

  if (msg.type === "zvl_fill") {
    const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
    try {
      const response = await chrome.runtime.sendMessage({
        action: "zvl_fill", vessel: msg.vessel
      });
      reply({ type: "zvl_fill_response", requestId, ...response });
    } catch (err) {
      reply({ type: "zvl_fill_response", requestId, ok: false, error: err.message });
    }
    return;
  }

  // ── Legacy sfbr_inject — kept for one transition release, just delegates ─
  if (msg.type === "sfbr_inject") {
    const requestId = msg.requestId;
    try {
      await dispatchOp("css-on-tab",  { tabId: "active", cssUrl:    "sfbr-styles.css" });
      await dispatchOp("exec-on-tab", { tabId: "active", scriptUrl: "sfbr-seafarers.js", world: "MAIN" });
      reply({ type: "sfbr_inject_response", requestId, ok: true });
    } catch (err) {
      reply({ type: "sfbr_inject_response", requestId, ok: false, error: err.message });
    }
    return;
  }
});

async function safeTabInfo() {
  try {
    const tab = await findActiveTab();
    if (!tab) return null;
    return { id: tab.id, url: tab.url || "(url not available)", title: tab.title || "" };
  } catch (err) {
    return { error: err.message };
  }
}
