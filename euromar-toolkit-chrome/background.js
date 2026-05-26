const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";

const POPUP_WIDTH = 500;
const POPUP_HEIGHT = 720;

// ── Context menu ─────────────────────────────────────────────────────────────

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: "euromar-toolkit-menu",
    title: "EUROMAR Toolkit",
    contexts: ["selection", "page"]
  });
});

// ── Icon click → open popup window ───────────────────────────────────────────

chrome.action.onClicked.addListener(async (tab) => {
  const { left, top } = await getPopupPosition(tab);
  chrome.windows.create({
    url: chrome.runtime.getURL("popup.html"),
    type: "popup",
    width: POPUP_WIDTH,
    height: POPUP_HEIGHT,
    left,
    top
  });
});

// ── Context menu click → extract number, open popup ──────────────────────────

chrome.contextMenus.onClicked.addListener((info) => {
  const selection = info.selectionText ?? "";
  const match = selection.match(/\d{4,7}/);
  const base = chrome.runtime.getURL("popup.html");
  const url = match ? `${base}?selection=${encodeURIComponent(match[0])}` : base;
  chrome.windows.create({
    url,
    type: "popup",
    width: POPUP_WIDTH,
    height: POPUP_HEIGHT
  });
});

// ── Position helper: top-right of the focused browser window ─────────────────

async function getPopupPosition(tab) {
  try {
    const win = await chrome.windows.get(tab.windowId);
    const left = Math.max(0, (win.left ?? 0) + (win.width ?? 1280) - POPUP_WIDTH - 20);
    const top = (win.top ?? 0) + 60;
    return { left, top };
  } catch {
    return { left: undefined, top: undefined };
  }
}

// ── API bridge: forwards requests from the popup iframe ──────────────────────
// Generic handler — any toolkit tool can POST to any endpoint via this bridge.

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === "toolkit-api") {
    handleApiCall(msg).then(sendResponse);
    return true; // keep channel open for the async response
  }
});

async function handleApiCall(msg) {
  const stored = await chrome.storage.local.get(["apiBase", "apiKey"]);
  const apiBase = stored.apiBase ?? "https://pscplatformalpha.onrender.com";
  const apiKey = stored.apiKey;

  if (!apiKey) {
    return { ok: false, error: "API key not set. Configure it in the toolkit settings." };
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 120_000);

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
