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

// ── Live selection updates ────────────────────────────────────────────────────

chrome.storage.session.onChanged.addListener((changes) => {
  const newValue = changes.pendingSelection?.newValue;
  if (!newValue) return;
  chrome.storage.session.remove("pendingSelection");
  iframe.src = `${BASE_SRC}&selection=${encodeURIComponent(newValue)}`;
});

// ── Message bridge ────────────────────────────────────────────────────────────
// Only accept messages originating from our iframe; reply only to its origin.

window.addEventListener("message", async (event) => {
  if (event.source !== iframe.contentWindow) return;

  const msg = event.data;
  if (!msg || typeof msg.type !== "string") return;

  switch (msg.type) {
    case "close":
      window.close();
      break;

    case "resize":
      // Reserved for future use — toolkit can suggest a preferred panel size.
      break;

    case "toolkit-api": {
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      try {
        const response = await chrome.runtime.sendMessage({
          action: "toolkit-api",
          endpoint: msg.endpoint,
          body: msg.body
        });
        iframe.contentWindow.postMessage(
          { type: "toolkit-api-response", requestId, ...response },
          TOOLKIT_ORIGIN
        );
      } catch (err) {
        iframe.contentWindow.postMessage(
          { type: "toolkit-api-response", requestId, ok: false, error: err.message },
          TOOLKIT_ORIGIN
        );
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
        iframe.contentWindow.postMessage(
          { type: "zvl_fill_response", requestId, ...response },
          TOOLKIT_ORIGIN
        );
      } catch (err) {
        iframe.contentWindow.postMessage(
          { type: "zvl_fill_response", requestId, ok: false, error: err.message },
          TOOLKIT_ORIGIN
        );
      }
      break;
    }

    case "chkbox_inject": {
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      const SFBR_ORIGINS = [
        "https://seafarers.eu-registry.com",
        "https://seafarers-web-test.idego.io"
      ];
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) throw new Error("No active tab found.");
        if (tab.url) {
          const origin = new URL(tab.url).origin;
          if (!SFBR_ORIGINS.includes(origin)) {
            throw new Error("Active tab is not the Seafarers Panel. Open seafarers.eu-registry.com and try again.");
          }
        }
        const [{ result }] = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: () => {
            const boxes = document.querySelectorAll(".ui-chkbox-box");
            let count = 0;
            for (const box of boxes) {
              if (!box.querySelector(".ui-icon-check")) {
                box.click();
                count++;
              }
            }
            return { ok: true, count };
          }
        });
        iframe.contentWindow.postMessage(
          { type: "chkbox_inject_response", requestId, ...(result ?? { ok: true, count: 0 }) },
          TOOLKIT_ORIGIN
        );
      } catch (err) {
        iframe.contentWindow.postMessage(
          { type: "chkbox_inject_response", requestId, ok: false, error: err.message },
          TOOLKIT_ORIGIN
        );
      }
      break;
    }

    case "sfbr_inject": {
      // Handled directly here — no background.js hop needed.
      // Extension pages (including side panels) can call chrome.scripting directly,
      // and doing it here avoids the MV3 service-worker response-dropping issue.
      const requestId = msg.requestId ?? `${Date.now()}-${Math.random()}`;
      const SFBR_ORIGINS = [
        "https://seafarers.eu-registry.com",
        "https://seafarers-web-test.idego.io"
      ];
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab) throw new Error("No active tab found.");

        // tab.url is only populated when the manifest includes the 'tabs' permission.
        // If it is available, verify we're on the Seafarers Panel; otherwise trust the user.
        if (tab.url) {
          const origin = new URL(tab.url).origin;
          if (!SFBR_ORIGINS.includes(origin)) {
            throw new Error(
              "Active tab is not the Seafarers Panel. Open seafarers.eu-registry.com and try again."
            );
          }
        }

        await chrome.scripting.insertCSS({
          target: { tabId: tab.id },
          files:  ["sfbr-styles.css"]
        });
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files:  ["sfbr-seafarers.js"],
          world:  "MAIN"
        });
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files:  ["sfbr-relay.js"],
          world:  "ISOLATED"
        });

        iframe.contentWindow.postMessage(
          { type: "sfbr_inject_response", requestId, ok: true },
          TOOLKIT_ORIGIN
        );
      } catch (err) {
        iframe.contentWindow.postMessage(
          { type: "sfbr_inject_response", requestId, ok: false, error: err.message },
          TOOLKIT_ORIGIN
        );
      }
      break;
    }
  }
});
