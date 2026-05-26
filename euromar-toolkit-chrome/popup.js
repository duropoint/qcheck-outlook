const TOOLKIT_URL = "https://duropoint.github.io/qcheck-outlook/taskpane.html";

const TOOLKIT_ORIGIN = new URL(TOOLKIT_URL).origin;
const iframe = document.getElementById("contentFrame");

// Forward all URL params from popup.html?… to the hosted toolkit
const params = new URLSearchParams(window.location.search);
iframe.src = params.toString() ? `${TOOLKIT_URL}?${params.toString()}` : TOOLKIT_URL;

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
      // Reserved for future use — toolkit can suggest a preferred window size.
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
