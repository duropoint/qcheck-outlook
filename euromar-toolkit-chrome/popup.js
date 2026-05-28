const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";
const TOOLKIT_ORIGIN = new URL(TOOLKIT_URL).origin;
const iframe = document.getElementById("contentFrame");

// Always append ?context=extension so the hosted toolkit knows it's in the shell
const BASE_SRC = `${TOOLKIT_URL}?context=extension`;

// Set src synchronously so the panel never opens blank
iframe.src = BASE_SRC;

// Then check if a context-menu selection is waiting and append it
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

// ── Message bridge ────────────────────────────────────────────────────────────
// Accept messages from the toolkit origin (more reliable than comparing
// against iframe.contentWindow, which can become stale across re-navigations).

function reply(payload) {
  // Use "*" as the target origin: the hosted iframe is at TOOLKIT_ORIGIN but we
  // want the response to always be delivered, even if the iframe's effective
  // origin appears different after navigation. The payload itself is harmless.
  try {
    iframe.contentWindow.postMessage(payload, "*");
  } catch (err) {
    console.error("[Toolkit] Failed to post response back to iframe:", err);
  }
}

async function findActiveTab() {
  // Side panels run in a window context. Try currentWindow first, then fall
  // back to lastFocusedWindow (handles edge cases where the side panel itself
  // is the focused surface).
  let tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tabs.length) tabs = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  return tabs[0] || null;
}

window.addEventListener("message", async (event) => {
  if (event.origin !== TOOLKIT_ORIGIN) return;

  const msg = event.data;
  if (!msg || typeof msg.type !== "string") return;

  console.log("[Toolkit/popup] received:", msg.type, msg.requestId || "");

  switch (msg.type) {
    case "close":
      window.close();
      break;

    case "resize":
      // Reserved for future use
      break;

    // Diagnostic: tells the toolkit which extension version is responding,
    // so we can confirm the new code is loaded.
    case "sfbr_ping": {
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      const version = chrome.runtime.getManifest().version;
      let tabInfo = null;
      try {
        const tab = await findActiveTab();
        if (tab) {
          tabInfo = {
            id:    tab.id,
            url:   tab.url || "(url not available — host_permissions may be missing)",
            title: tab.title || ""
          };
        }
      } catch (err) {
        tabInfo = { error: err.message };
      }
      reply({ type: "sfbr_ping_response", requestId, ok: true, version, tabInfo });
      break;
    }

    case "toolkit-api": {
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      try {
        const response = await chrome.runtime.sendMessage({
          action: "toolkit-api",
          endpoint: msg.endpoint,
          body: msg.body
        });
        reply({ type: "toolkit-api-response", requestId, ...response });
      } catch (err) {
        reply({ type: "toolkit-api-response", requestId, ok: false, error: err.message });
      }
      break;
    }

    case "zvl_fill": {
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      try {
        const response = await chrome.runtime.sendMessage({
          action: "zvl_fill",
          vessel: msg.vessel
        });
        reply({ type: "zvl_fill_response", requestId, ...response });
      } catch (err) {
        reply({ type: "zvl_fill_response", requestId, ok: false, error: err.message });
      }
      break;
    }

    case "sfbr_inject": {
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      const SFBR_ORIGINS = [
        "https://seafarers.eu-registry.com",
        "https://seafarers-web-test.idego.io"
      ];
      let step = "find tab";
      try {
        const tab = await findActiveTab();
        if (!tab) throw new Error("No active tab found in the current window.");
        if (!tab.id) throw new Error("Active tab has no usable tab ID.");

        // If url is available (host_permissions match), verify it's Seafarers.
        // If url is hidden by Chrome, we still try — the user told us to.
        if (tab.url) {
          const origin = new URL(tab.url).origin;
          if (!SFBR_ORIGINS.includes(origin)) {
            throw new Error(
              `Active tab is "${tab.url}". Switch to the Seafarers Panel tab first, then click Inject.`
            );
          }
        }

        step = "insert CSS";
        await chrome.scripting.insertCSS({
          target: { tabId: tab.id },
          files:  ["sfbr-styles.css"]
        });

        step = "inject MAIN-world script";
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files:  ["sfbr-seafarers.js"],
          world:  "MAIN"
        });

        step = "inject ISOLATED-world relay";
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files:  ["sfbr-relay.js"],
          world:  "ISOLATED"
        });

        reply({
          type: "sfbr_inject_response",
          requestId,
          ok: true,
          tabId: tab.id,
          tabUrl: tab.url || "(hidden)"
        });
      } catch (err) {
        console.error(`[Toolkit/popup] sfbr_inject failed at step "${step}":`, err);
        reply({
          type: "sfbr_inject_response",
          requestId,
          ok: false,
          error: `${err.message} (step: ${step})`
        });
      }
      break;
    }
  }
});
