const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";

// ── Context menu + side panel behaviour on install ───────────────────────────

chrome.runtime.onInstalled.addListener(() => {
  // Toolbar icon click toggles the side panel automatically
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
    // Store temporarily so popup.js can read it (handles both fresh open and already-open panel)
    await chrome.storage.session.set({ pendingSelection: match[0] });
  }

  await chrome.sidePanel.open({ windowId: tab.windowId });
});

// ── API bridge: forwards requests from the panel iframe ──────────────────────
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
