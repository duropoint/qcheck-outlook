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

  // ── ZOHOtoBMar: extract seafarer data from the active Zoho CRM tab ─────────
  if (msg.type === "zoho-extract") {
    const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
    try {
      const tabs = await chrome.tabs.query({ currentWindow: true });
      const zohoTab = tabs.find(t => t.url && t.url.includes("zoho.com"));
      if (!zohoTab) throw new Error("No Zoho tab found. Open the seafarer record in Zoho first.");

      const results = await chrome.scripting.executeScript({
        target: { tabId: zohoTab.id, allFrames: true },
        func: ztbExtractZohoData
      });

      let fields = {};
      for (const frame of results) {
        const r = frame?.result || {};
        const filledR   = Object.values(r).filter(v => v && String(v).length).length;
        const filledCur = Object.values(fields).filter(v => v && String(v).length).length;
        if (filledR > filledCur) fields = r;
      }
      if (!fields["First Name"] && !fields["Last Name"]) {
        throw new Error("No seafarer data found. Make sure the Zoho record is open.");
      }
      reply({ type: "zoho-extract-response", requestId, ok: true, fields });
    } catch (err) {
      reply({ type: "zoho-extract-response", requestId, ok: false, error: err.message });
    }
    return;
  }

  // ── ZOHOtoBMar: drive the BMAR form automation ───────────────────────────────
  if (msg.type === "bmar-automate") {
    const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
    const { fields, documents, options } = msg.payload || {};
    try {
      const tabs = await chrome.tabs.query({ currentWindow: true });
      const bmarTab = tabs.find(t => t.url && t.url.includes("bmar.pt"));
      if (!bmarTab) throw new Error("No BMAR tab found. Open the BMAR application form first.");

      // Ping content script; inject if absent
      let scriptReady = false;
      try {
        const pong = await new Promise((resolve, reject) => {
          chrome.tabs.sendMessage(bmarTab.id, { action: "ping" }, resp => {
            if (chrome.runtime.lastError) reject(chrome.runtime.lastError);
            else resolve(resp);
          });
        });
        scriptReady = pong?.status === "pong";
      } catch (_) {}

      if (!scriptReady) {
        await chrome.scripting.executeScript({ target: { tabId: bmarTab.id }, files: ["zohotobmar-content.js"] });
        await new Promise(r => setTimeout(r, 1000));
      }

      chrome.tabs.sendMessage(bmarTab.id, { action: "automate", fields, documents, options });
      reply({ type: "bmar-automate-response", requestId, ok: true });
    } catch (err) {
      reply({ type: "bmar-automate-response", requestId, ok: false, error: err.message });
    }
    return;
  }

  // ── ZOHOtoBMar: retry failed document uploads ────────────────────────────────
  if (msg.type === "bmar-upload-retry") {
    const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
    const { documents } = msg.payload || {};
    try {
      const tabs = await chrome.tabs.query({ currentWindow: true });
      const bmarTab = tabs.find(t => t.url && t.url.includes("bmar.pt"));
      if (!bmarTab) throw new Error("No BMAR tab found.");
      chrome.tabs.sendMessage(bmarTab.id, { action: "uploadOnly", documents });
      reply({ type: "bmar-upload-retry-response", requestId, ok: true });
    } catch (err) {
      reply({ type: "bmar-upload-retry-response", requestId, ok: false, error: err.message });
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

// ── ZOHOtoBMar: progress relay ───────────────────────────────────────────────
// Progress events are sent by zohotobmar-content.js → background.js (which
// tags them _relayed: true) → here → iframe.
chrome.runtime.onMessage.addListener((msg) => {
  if (msg._relayed && msg.type && msg.type.startsWith("bmar-")) {
    reply(msg);
  }
});

// ── ZOHOtoBMar: data extraction function (injected into Zoho tab) ────────────
function ztbExtractZohoData() {
  const q      = (s, root = document) => root.querySelector(s);
  const getVal = (...sels) => { for (const s of sels) { const el = q(s); if (el) return (el.value ?? el.textContent ?? "").trim(); } return ""; };
  const getEl  = (...sels) => { for (const s of sels) { const el = q(s); if (el) return el; } return null; };
  const getSelect = (...sels) => {
    const el = getEl(...sels);
    if (!el) return "";
    const optText = el.selectedOptions && el.selectedOptions[0] ? el.selectedOptions[0].text : "";
    return (optText || el.value || "").trim();
  };
  const getDateBy = (legacyName, ...hints) => {
    let v = getVal(`input[name="${legacyName}"]`);
    if (v) return v;
    v = getVal('input[type="date"]');
    if (v) return v;
    const tokens = hints.join(" ").toLowerCase();
    for (const inp of document.querySelectorAll("input")) {
      const pl = (inp.getAttribute("placeholder") || "").toLowerCase();
      const al = (inp.getAttribute("aria-label")  || "").toLowerCase();
      if (tokens && tokens.split(/\s+/).every(t => al.includes(t) || pl.includes(t))) {
        if (inp.value) return inp.value.trim();
      }
    }
    return "";
  };

  const data = {};

  data["First Name"] = getVal('input[name="Name"][complink="Name_First"]', 'input[complink="Name_First"]', 'input[aria-label*="first" i]');
  data["Last Name"]  = getVal('input[name="Name"][complink="Name_Last"]',  'input[complink="Name_Last"]',  'input[aria-label*="last" i]');

  data["Ship"]              = getVal('input[name="SingleLine"]',  'input[placeholder*="Ship" i]',     'input[aria-label*="Ship" i]');
  data["Passport Number"]   = getVal('input[name="SingleLine1"]', 'input[placeholder*="Passport" i]', 'input[aria-label*="Passport" i]');
  data["Person in Charge"]  = getVal('input[name="SingleLine2"]', 'input[aria-label*="person in charge" i]');
  data["STCW Rule"]         = getVal('input[name="SingleLine3"]', 'input[aria-label*="stcw" i]');
  data["COC Number"]        = getVal('input[name="SingleLine4"]', 'input[placeholder*="COC" i]');
  data["COC Endorsement Number"] = getVal('input[name="SingleLine5"]', 'input[aria-label*="endorsement" i]');
  data["GOC Number"]        = getVal('input[name="SingleLine6"]', 'input[placeholder*="GMDSS" i], input[placeholder*="GOC" i]');
  data["GOC Endorsement Number"] = getVal('input[name="SingleLine7"]');
  data["COP1 Number"]       = getVal('input[name="SingleLine8"]', 'input[placeholder*="COP1" i]');
  data["COP2 Number"]       = getVal('input[name="SingleLine9"]', 'input[placeholder*="COP2" i]');
  data["Invoice"]           = "Applicant";

  data["Capacity"]              = getSelect('select[name="Dropdown"]',   'select[aria-label*="capacity" i]');
  data["Sex"]                   = getSelect('select[name="Dropdown1"]',  'select[aria-label*="sex" i]');
  data["Country of Origin"]     = getSelect('select[name="Dropdown2"]',  'select[aria-label*="country" i]');
  data["COC Issued By"]         = getSelect('select[name="Dropdown8"]');
  data["COC Endorsement Issued By"] = getSelect('select[name="Dropdown9"]');
  data["GOC Issued By"]         = getSelect('select[name="Dropdown10"]');
  data["GOC Endorsement Issued By"] = getSelect('select[name="Dropdown11"]');
  data["COP1 Issued By"]        = getSelect('select[name="Dropdown12"]');
  data["COP2 Issued By"]        = getSelect('select[name="Dropdown13"]');

  const dateMap = {
    "Date": "Birthdate", "Date2": "Passport Validity",
    "Date4": "Medical Issuance date", "Date5": "Medical Expiry",
    "Date11": "COC Issuance", "Date12": "COC Expiry",
    "Date13": "COC Endorsement Issuance", "Date14": "COC Endorsement Expiry",
    "Date15": "GOC Issuance", "Date16": "GOC Expiry",
    "Date17": "GOC Endorsement Issuance", "Date18": "GOC Endorsement Expiry",
    "Date19": "COP1 Issuance", "Date20": "COP1 Expiry",
    "Date21": "COP2 Issuance", "Date22": "COP2 Expiry",
    "Date33": "COC Revalidation date", "Date34": "COC Revalidation Expiry date",
    "Date36": "GOC Revalidation date", "Date35": "GOC Revalidation Expiry date"
  };
  const dateHints = {
    "Birthdate": ["birth", "date"], "Passport Validity": ["passport", "valid"],
    "COC Issuance": ["coc", "issu"], "COC Expiry": ["coc", "expiry"],
    "GOC Issuance": ["goc", "gmdss", "issu"], "GOC Expiry": ["goc", "gmdss", "expiry"],
    "COP1 Issuance": ["cop1", "issu"], "COP1 Expiry": ["cop1", "expiry"],
    "COP2 Issuance": ["cop2", "issu"], "COP2 Expiry": ["cop2", "expiry"],
    "Medical Issuance date": ["medical", "issu"], "Medical Expiry": ["medical", "expiry"]
  };
  for (const legacy in dateMap) {
    const label = dateMap[legacy];
    data[label] = getDateBy(legacy, ...(dateHints[label] || []));
  }

  data["Invoice Address"]  = getVal('textarea[name="MultiLine"]',  'textarea[aria-label*="invoice" i]');
  data["Dispatch Address"] = getVal('textarea[name="MultiLine1"]', 'textarea[aria-label*="dispatch" i]');
  data["PIC Email"]        = getVal('input[name="Email"]', 'input[type="email"]');

  return data;
}
