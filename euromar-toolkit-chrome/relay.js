// relay.js — GENERIC relay, injected in ISOLATED world.
//
// Watches the DOM for relay-signal elements written by hosted tool scripts
// (which run in MAIN world and have no chrome.runtime access) and forwards
// them to background.js.
//
// Signal element format:
//   <span class="__euromar_relay__"
//         data-euromar-action="op-name"
//         data-euromar-payload='{...JSON...}'></span>
//
// Once injected, this file never needs to change. Future tools just use the
// same signaling pattern.

(function () {
  'use strict';
  if (window.__euromarRelayInstalled) return;
  window.__euromarRelayInstalled = true;

  const SIGNAL_CLASS = '__euromar_relay__';

  function handleSignal(node) {
    if (!node || node.nodeType !== 1) return;
    if (!node.classList || !node.classList.contains(SIGNAL_CLASS)) return;

    const action = node.getAttribute('data-euromar-action');
    let payload = {};
    try {
      const raw = node.getAttribute('data-euromar-payload');
      if (raw) payload = JSON.parse(raw);
    } catch (e) {
      console.warn('[EUROMAR relay] invalid payload JSON', e);
    }
    node.remove();
    if (action) {
      try {
        chrome.runtime.sendMessage({ source: 'relay', action, payload });
      } catch (e) {
        console.error('[EUROMAR relay] sendMessage failed', e);
      }
    }
  }

  const observer = new MutationObserver((muts) => {
    for (const m of muts) m.addedNodes.forEach(handleSignal);
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });

  // Process any signal already present at injection time
  document.querySelectorAll('.' + SIGNAL_CLASS).forEach(handleSignal);
})();
