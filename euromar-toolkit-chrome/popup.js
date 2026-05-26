const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";

const TOOLKIT_ORIGIN = new URL(TOOLKIT_URL).origin;
const iframe = document.getElementById("contentFrame");

// Set src synchronously so the panel never opens blank
iframe.src = TOOLKIT_URL;

// Then check if a context-menu selection is waiting and update src if so
chrome.storage.session.get("pendingSelection")
  .then(({ pendingSelection }) => {
    if (!pendingSelection) return;
    chrome.storage.session.remove("pendingSelection");
    iframe.src = `${TOOLKIT_URL}?selection=${encodeURIComponent(pendingSelection)}`;
  })
  .catch(() => {}); // src is already set above if storage is unavailable

// ── Live selection updates ────────────────────────────────────────────────────
// If the panel is already open when a context-menu click arrives, react here
// rather than waiting for a full reload.

chrome.storage.session.onChanged.addListener((changes) => {
  const newValue = changes.pendingSelection?.newValue;
  if (!newValue) return;
  chrome.storage.session.remove("pendingSelection");
  iframe.src = `${TOOLKIT_URL}?selection=${encodeURIComponent(newValue)}`;
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
  }
});
