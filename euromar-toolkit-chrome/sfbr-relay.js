// sfbr-relay.js — injected in ISOLATED world on the Seafarers Panel tab.
//
// Bridges the Zoho URL from sfbr-seafarers.js (MAIN world) to the extension
// background service worker via chrome.runtime.sendMessage.
//
// Uses a MutationObserver watching for a temporary DOM element written by
// sfbr-seafarers.js, rather than CustomEvent.detail, because Chrome does not
// reliably forward CustomEvent detail across the MAIN ↔ ISOLATED world boundary.

(function () {
  'use strict';

  const SIGNAL_ID = '__sfbr_zoho_signal__';

  function handleSignal(node) {
    if (node.nodeType !== 1 || node.id !== SIGNAL_ID) return;
    const url = node.getAttribute('data-zoho-url');
    node.remove(); // consume the signal
    if (url) {
      chrome.runtime.sendMessage({ action: 'sfbr_open_zoho', url });
    }
  }

  // Watch for the signal element being added to the DOM
  const observer = new MutationObserver((mutations) => {
    for (const m of mutations) {
      m.addedNodes.forEach(handleSignal);
    }
  });

  observer.observe(document.documentElement, { childList: true, subtree: true });

  // Also handle the case where the element is already present at inject time
  const existing = document.getElementById(SIGNAL_ID);
  if (existing) handleSignal(existing);
})();
