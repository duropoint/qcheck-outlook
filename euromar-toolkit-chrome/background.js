// background.js — service worker.
// Handles:
//   - Side-panel install + context menu (unchanged)
//   - The legacy toolkit-api and zvl_fill paths (called by popup.js)
//   - Relay-originated messages from hosted tool scripts (via relay.js)
//
// The relay path uses the same 8 generic ops as popup.js so a script
// injected into any page can request follow-up actions (open another tab,
// fill it, badge, notify, download, etc.) without going back to popup.

const SCRIPTS_BASE = "https://duropoint.github.io/qcheck-outlook/scripts/";

chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  chrome.contextMenus.create({
    id: "euromar-toolkit-menu",
    title: "EUROMAR Toolkit",
    contexts: ["selection", "page"]
  });
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  const selection = info.selectionText ?? "";
  const match = selection.match(/\d{4,7}/);
  if (match) await chrome.storage.session.set({ pendingSelection: match[0] });
  await chrome.sidePanel.open({ windowId: tab.windowId });
});

// ── Message router ──────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === "toolkit-api") {
    handleApiCall(msg).then(sendResponse);
    return true;
  }
  if (msg.action === "zvl_fill") {
    handleZvlFill(msg.vessel).then(sendResponse);
    return true;
  }
  if (msg.source === "relay" && typeof msg.action === "string") {
    handleRelay(msg.action, msg.payload || {}, sender)
      .then(sendResponse)
      .catch(err => {
        console.error(`[Toolkit/bg] relay action ${msg.action} failed:`, err);
        sendResponse({ ok: false, error: err.message });
      });
    return true;
  }
  // Forward bmar-* progress events from the BMAR content script to the side panel.
  // sender.tab is set when the message originates from a content script.
  if (msg._fromContent && msg.type && msg.type.startsWith("bmar-") && sender?.tab) {
    chrome.runtime.sendMessage({ ...msg, _fromContent: false, _relayed: true }).catch(() => {});
    sendResponse({ ok: true });
    return true;
  }
});

// ── Helpers ─────────────────────────────────────────────────────────────────

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

// ── Generic operations (same set as popup.js) ───────────────────────────────

async function opExecOnTab({ tabId, scriptUrl, world }) {
  if (!tabId) throw new Error("exec-on-tab requires tabId");
  const code = await fetchText(scriptUrl);
  await chrome.scripting.executeScript({
    target: { tabId }, files: ["relay.js"], world: "ISOLATED"
  });
  await chrome.scripting.executeScript({
    target: { tabId },
    world: world === "ISOLATED" ? "ISOLATED" : "MAIN",
    func: (code, url) => {
      try { new Function(code).call(window); }
      catch (e) { console.error("[EUROMAR] script error", url, e); throw e; }
    },
    args: [code, scriptUrl]
  });
  return { ok: true, tabId };
}

async function opCssOnTab({ tabId, cssUrl }) {
  if (!tabId) throw new Error("css-on-tab requires tabId");
  const css = await fetchText(cssUrl);
  await chrome.scripting.insertCSS({ target: { tabId }, css });
  return { ok: true, tabId };
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
  if (!tabId) throw new Error("get-tab-info requires tabId");
  const tab = await chrome.tabs.get(tabId);
  return { ok: true, id: tab.id, url: tab.url || "", title: tab.title || "" };
}

async function opCloseTab({ tabId }, sender) {
  const id = tabId || sender?.tab?.id;
  if (!id) throw new Error("close-tab requires tabId");
  await chrome.tabs.remove(id);
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

async function dispatchOp(type, payload, sender) {
  switch (type) {
    case "exec-on-tab":  return opExecOnTab(payload);
    case "css-on-tab":   return opCssOnTab(payload);
    case "open-tab":     return opOpenTab(payload);
    case "get-tab-info": return opGetTabInfo(payload);
    case "close-tab":    return opCloseTab(payload, sender);
    case "notify":       return opNotify(payload);
    case "badge":        return opBadge(payload);
    case "download":     return opDownload(payload);
    default: throw new Error(`Unknown operation: ${type}`);
  }
}

// ── Relay handler ───────────────────────────────────────────────────────────
// A hosted tool script (running in MAIN world on some target page) emitted
// a signal element; relay.js (ISOLATED) caught it and forwarded the action
// here. We just dispatch.

async function handleRelay(action, payload, sender) {
  return await dispatchOp(action, payload, sender);
}

// ── Backwards-compat (existing toolkit features) ────────────────────────────

async function handleApiCall(msg) {
  const stored = await chrome.storage.local.get(["apiBase", "apiKey"]);
  const apiBase = stored.apiBase ?? "https://pscplatformalpha.onrender.com";
  const apiKey  = stored.apiKey;
  if (!apiKey) return { ok: false, error: "API key not set. Configure it in the toolkit settings." };

  const controller = new AbortController();
  const timeoutId  = setTimeout(() => controller.abort(), 120_000);
  try {
    const response = await fetch(`${apiBase}${msg.endpoint}`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-API-Key": apiKey },
      body: JSON.stringify(msg.body),
      signal: controller.signal
    });
    clearTimeout(timeoutId);
    if (!response.ok) {
      let errorMsg = response.statusText;
      try { const j = await response.json(); if (j.error) errorMsg = j.error; } catch {}
      return { ok: false, error: errorMsg, status: response.status };
    }
    return { ok: true, data: await response.json() };
  } catch (err) {
    clearTimeout(timeoutId);
    if (err.name === "AbortError") return { ok: false, error: "Request timed out after 120 seconds" };
    return { ok: false, error: err.message };
  }
}

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
  const filledName    = fillField(["vessel_name", "vesselname", "vessel-name"],            vessel.vessel_name  || "");
  const filledImo     = fillField(["vessel_imo",  "vesselimo",  "vessel-imo",  "imo"],     String(vessel.vessel_imo || ""));
  const filledDetails = fillField(["vessel_details", "vesseldetails", "vessel-details"],   details);
  return { ok: filledName || filledImo || filledDetails };
}
