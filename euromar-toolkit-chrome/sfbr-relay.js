// sfbr-relay.js — injected in ISOLATED world on the Seafarers Panel tab.
//
// Bridges CustomEvents fired by sfbr-seafarers.js (MAIN world) across the
// world boundary to the extension background service worker.
// CustomEvents are visible across both worlds per the Chrome extension model.

document.addEventListener("sfbr:open_zoho", (e) => {
  if (e.detail && e.detail.url) {
    chrome.runtime.sendMessage({ action: "sfbr_open_zoho", url: e.detail.url });
  }
});
