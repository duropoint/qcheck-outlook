const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";

// ── Context menu + side panel behaviour on install ───────────────────────────

chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });

  chrome.contextMenus.create({
    id: "euromar-toolkit-menu",
    title: "EUROMAR Toolkit",
    contexts: ["selection", "page"]
  });
});

// ── Context menu click → extract number, open side panel ─────────────────────

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  const selection = info.selectionText ?? "";
  const match = selection.match(/\d{4,7}/);

  if (match) {
    await chrome.storage.session.set({ pendingSelection: match[0] });
  }

  await chrome.sidePanel.open({ windowId: tab.windowId });
});

// ── Message router ────────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === "toolkit-api") {
    handleApiCall(msg).then(sendResponse);
    return true;
  }
  if (msg.action === "zvl_fill") {
    handleZvlFill(msg.vessel).then(sendResponse);
    return true;
  }
  if (msg.action === "sfbr_open_zoho") {
    // Called by the relay script (ISOLATED world) in the Seafarers tab
    handleSfbrOpenZoho(msg.url).catch(console.error);
    sendResponse({ ok: true });
    return false;
  }
});

// ── Zammad fill ───────────────────────────────────────────────────────────────
// Injects a self-contained fill function directly into the active Zammad tab.
// No content script registration needed — works on any domain via <all_urls>.

async function handleZvlFill(vessel) {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url?.startsWith("https://euromar.zammad.com/#ticket/")) {
    return { ok: false, error: "Open a Zammad case first" };
  }
  try {
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: injectFillVessel,
      args: [vessel]
    });
    return result ?? { ok: false, error: "No result from page" };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Injected into the Zammad page — must be fully self-contained (no closure refs).
function injectFillVessel(vessel) {
  function setNativeValue(el, value) {
    const proto  = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
    setter.call(el, value);
    el.dispatchEvent(new Event("input",  { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.dispatchEvent(new Event("blur",   { bubbles: true }));
  }

  function fillField(names, value) {
    if (!value) return false;
    for (const n of names) {
      const el = document.querySelector(
        `input[name="${n}"], input[data-name="${n}"], textarea[name="${n}"], textarea[data-name="${n}"]`
      );
      if (el) { setNativeValue(el, value); return true; }
    }
    const labels = document.querySelectorAll("label, .form-group label, .controls label");
    for (const lbl of labels) {
      const txt = (lbl.textContent || "").trim().toLowerCase();
      const wantsName    = names.some(n => n.includes("name"))    && txt.includes("vessel") && txt.includes("name");
      const wantsImo     = names.some(n => n.includes("imo"))     && txt.includes("imo")    && !txt.includes("manager");
      const wantsDetails = names.some(n => n.includes("details")) && txt.includes("vessel") && txt.includes("details");
      if (wantsName || wantsImo || wantsDetails) {
        const container = lbl.closest(".form-group, .controls, .row") || lbl.parentElement;
        const input = container?.querySelector("input, textarea");
        if (input) { setNativeValue(input, value); return true; }
      }
    }
    return false;
  }

  function formatNumber(v) {
    const n = Number(String(v).replace(/[^\d.-]/g, ""));
    return Number.isFinite(n) ? n.toLocaleString("en-US") : String(v);
  }

  function buildVesselDetails(r) {
    const today = new Date().toISOString().slice(0, 10);
    const gt    = r.gross_tonnage ? formatNumber(r.gross_tonnage) : "";
    return [
      r.vessel_type   ? `Vessel Type: ${r.vessel_type}`                                                           : null,
      r.class_society ? `Class Society: ${r.class_society}`                                                       : null,
      (gt || r.year_built) ? `GT: ${gt || "—"} | Built: ${r.year_built || "—"}`                                  : null,
      r.ism_manager   ? `ISM Manager: ${r.ism_manager}${r.ism_manager_imo ? ` (IMO ${r.ism_manager_imo})` : ""}` : null,
      `Updated: ${today}`
    ].filter(Boolean).join("\n");
  }

  const details = buildVesselDetails(vessel);
  const filledName    = fillField(["vessel_name",    "vesselname",    "vessel-name"],             vessel.vessel_name  || "");
  const filledImo     = fillField(["vessel_imo",     "vesselimo",     "vessel-imo",     "imo"],   String(vessel.vessel_imo || ""));
  const filledDetails = fillField(["vessel_details", "vesseldetails", "vessel-details"],          details);

  return { ok: filledName || filledImo || filledDetails };
}

// ── Seafarers → Zoho BMAR Bridge ─────────────────────────────────────────────
//
// Injection (sfbr_inject) is now handled directly in popup.js to avoid the
// MV3 service-worker response-dropping issue.
//
// Remaining flow handled here:
//   1. User clicks "Send to Zoho BMAR" on the Seafarers Panel.
//      sfbr-seafarers.js (MAIN) writes the URL to a DOM attribute and
//      sfbr-relay.js (ISOLATED) picks it up via MutationObserver, then calls
//      chrome.runtime.sendMessage({ action: "sfbr_open_zoho", url }).
//   2. sfbr_open_zoho handler creates the Zoho tab, waits for it to load, then
//      injects sfbr-zoho.js (MAIN) + sfbr-styles.css — the fill banner appears.

async function handleSfbrOpenZoho(url) {
  // Open the Zoho form in a new tab (URL already contains the base64 data in the hash)
  const tab = await chrome.tabs.create({ url });
  // Wait for the tab to fully load before injecting
  await waitForTabComplete(tab.id);
  // Small extra buffer for JS-heavy forms to finish rendering
  await new Promise(r => setTimeout(r, 1200));
  try {
    await chrome.scripting.insertCSS({
      target: { tabId: tab.id },
      files: ["sfbr-styles.css"]
    });
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["sfbr-zoho.js"],
      world: "MAIN"
    });
  } catch (err) {
    console.error("[SFBR] Failed to inject into Zoho tab:", err);
  }
}

/** Resolves when the given tab reaches status "complete", or after a timeout. */
function waitForTabComplete(tabId) {
  return new Promise((resolve) => {
    const timeout = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      resolve();
    }, 20000);

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

// ── Generic API bridge ────────────────────────────────────────────────────────

async function handleApiCall(msg) {
  const stored = await chrome.storage.local.get(["apiBase", "apiKey"]);
  const apiBase = stored.apiBase ?? "https://pscplatformalpha.onrender.com";
  const apiKey  = stored.apiKey;

  if (!apiKey) {
    return { ok: false, error: "API key not set. Configure it in the toolkit settings." };
  }

  const controller = new AbortController();
  const timeoutId  = setTimeout(() => controller.abort(), 120_000);

  try {
    const response = await fetch(`${apiBase}${msg.endpoint}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-API-Key": apiKey
      },
      body: JSON.stringify(msg.body),
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      let errorMsg = response.statusText;
      try {
        const errBody = await response.json();
        if (errBody.error) errorMsg = errBody.error;
      } catch { /* ignore parse failures */ }
      return { ok: false, error: errorMsg, status: response.status };
    }

    const data = await response.json();
    return { ok: true, data };
  } catch (err) {
    clearTimeout(timeoutId);
    if (err.name === "AbortError") {
      return { ok: false, error: "Request timed out after 120 seconds" };
    }
    return { ok: false, error: err.message };
  }
}
