// sfbr-seafarers.js — injected in MAIN world on the Seafarers Panel tab.
//
// Adds a "Send to Zoho BMAR" button next to the form's submit button.
// When clicked, serialises the visible seafarer fields and signals the
// extension (via a DOM attribute that sfbr-relay.js watches) to open the
// Zoho form in a new tab and auto-fill it.
//
// Uses DOM-attribute signalling instead of CustomEvent.detail because
// Chrome does not reliably forward CustomEvent detail across the
// MAIN ↔ ISOLATED world boundary in all versions.

(function () {
  'use strict';

  // Guard — prevent double-injection
  if (window.__sfbrInjected) return;
  window.__sfbrInjected = true;

  // ── Button injection ────────────────────────────────────────────────────────

  function init() {
    // Try to add the button whenever there is a submit button on the page.
    // addTransferButton() already guards against duplicates.
    addTransferButton();
  }

  function addTransferButton() {
    if (document.getElementById('zoho-transfer-btn')) return; // already present

    // Find submit button or any prominent action button
    const submitBtn = document.querySelector('button[type="submit"]')
      || Array.from(document.querySelectorAll('button'))
           .find(b => /confirm|submit|save/i.test((b.textContent || '').trim()));

    if (!submitBtn) return;

    const transferBtn = document.createElement('button');
    transferBtn.id = 'zoho-transfer-btn';
    transferBtn.type = 'button';
    transferBtn.className = 'zoho-transfer-btn';
    transferBtn.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
           stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M5 12h14"></path><path d="m12 5 7 7-7 7"></path>
      </svg>
      Send to Zoho BMAR
    `;
    transferBtn.addEventListener('click', handleTransfer);
    submitBtn.parentNode.insertBefore(transferBtn, submitBtn);
  }

  // ── Data extraction ─────────────────────────────────────────────────────────

  function extractData() {
    const getValue = (id) => { const el = document.getElementById(id); return el ? (el.value || '') : ''; };

    const formatDateForZoho = (dateStr) => {
      if (!dateStr || dateStr === '' || dateStr.includes('YYYY')) return '';
      const normalized = dateStr.replace(/\//g, '-');
      const parts = normalized.split('-');
      if (parts.length !== 3) return '';
      const [year, month, day] = parts;
      return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
    };

    const shipInput    = document.getElementById('ship_search');
    const invoiceInput = document.getElementById('invoice_recipient_search');

    return {
      ship:            shipInput    ? shipInput.value    : '',
      invoiceRecipient: invoiceInput ? invoiceInput.value : '',
      fullName:        getValue('passport_name'),
      sex:             getValue('passport_sex'),
      birthdate:       formatDateForZoho(getValue('passport_birthdate')),
      countryOfOrigin: getValue('passport_country'),
      passportNumber:  getValue('passport_number'),
      passportValidity: formatDateForZoho(getValue('passport_validity')),
      medicalIssuanceDate: formatDateForZoho(getValue('mc_issue_date')),
      medicalExpiryDate:   formatDateForZoho(getValue('mc_expiry_date')),
      cocNumber:           getValue('coc_number'),
      cocIssuanceDate:     formatDateForZoho(getValue('coc_issuance_date')),
      cocExpiryDate:       formatDateForZoho(getValue('coc_expiry_date')),
      cocIssuedBy:         getValue('coc_issued_by'),
      cocRevalidationIssuanceDate: formatDateForZoho(getValue('coc_revalidation_issuance_date')),
      cocRevalidationExpiryDate:   formatDateForZoho(getValue('coc_revalidation_expiry_date')),
      cocEndorsementNumber:      getValue('coc_endorsement_number'),
      cocEndorsementIssuanceDate: formatDateForZoho(getValue('coc_endorsement_issuance_date')),
      cocEndorsementExpiryDate:   formatDateForZoho(getValue('coc_endorsement_expiry_date')),
      cocEndorsementIssuedBy:     getValue('coc_endorsement_issued_by'),
      gocNumber:           getValue('goc_number'),
      gocIssuanceDate:     formatDateForZoho(getValue('goc_issuance_date')),
      gocExpiryDate:       formatDateForZoho(getValue('goc_expiry_date')),
      gocIssuedBy:         getValue('goc_issued_by'),
      gocRevalidationIssuanceDate: formatDateForZoho(getValue('goc_revalidation_issuance_date')),
      gocRevalidationExpiryDate:   formatDateForZoho(getValue('goc_revalidation_expiry_date')),
      gocEndorsementNumber:      getValue('goc_endorsement_number'),
      gocEndorsementIssuanceDate: formatDateForZoho(getValue('goc_endorsement_issuance_date')),
      gocEndorsementExpiryDate:   formatDateForZoho(getValue('goc_endorsement_expiry_date')),
      gocEndorsementIssuedBy:     getValue('goc_endorsement_issued_by'),
      cop1Number:      getValue('cop1_number'),
      cop1IssuanceDate: formatDateForZoho(getValue('cop1_issuance_date')),
      cop1ExpiryDate:   formatDateForZoho(getValue('cop1_expiry_date')),
      cop1IssuedBy:     getValue('cop1_issued_by'),
      cop2Number:      getValue('cop2_number'),
      cop2IssuanceDate: formatDateForZoho(getValue('cop2_issuance_date')),
      cop2ExpiryDate:   formatDateForZoho(getValue('cop2_expiry_date')),
      cop2IssuedBy:     getValue('cop2_issued_by'),
      extractedAt: new Date().toISOString()
    };
  }

  // ── Transfer handler ────────────────────────────────────────────────────────

  function handleTransfer() {
    const btn = document.getElementById('zoho-transfer-btn');
    try {
      const data = extractData();
      if (!data.fullName || data.fullName.trim() === '') {
        showNotification('Error: Seafarer name not found. Are you on the Review Extracted Data page?', 'error');
        return;
      }

      const encodedData = btoa(encodeURIComponent(JSON.stringify(data)));
      const zohoUrl = `https://forms.zoho.com/europeanmarlda/form/BMAR#seafarerData=${encodedData}`;

      btn.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
             stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="20 6 9 17 4 12"></polyline>
        </svg>
        Data sent!
      `;
      btn.classList.add('success');
      showNotification(`Data for ${data.fullName} sent to Zoho`, 'success');

      // Signal sfbr-relay.js (ISOLATED world) via a DOM attribute on a temporary
      // element. Using DOM attributes is reliable across MAIN ↔ ISOLATED worlds,
      // unlike CustomEvent.detail which Chrome can silently strip.
      setTimeout(() => {
        const signal = document.createElement('span');
        signal.id = '__sfbr_zoho_signal__';
        signal.setAttribute('data-zoho-url', zohoUrl);
        signal.style.display = 'none';
        document.documentElement.appendChild(signal);
        // sfbr-relay.js removes the element after reading it
      }, 400);

      // Auto-submit the seafarers form after transfer
      setTimeout(() => {
        const confirmBtn = document.querySelector('button[type="submit"]')
          || Array.from(document.querySelectorAll('button'))
               .find(b => b.textContent.toLowerCase().includes('confirm'));
        if (confirmBtn) confirmBtn.click();
      }, 800);

      setTimeout(() => {
        btn.innerHTML = `
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none"
               stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M5 12h14"></path><path d="m12 5 7 7-7 7"></path>
          </svg>
          Send to Zoho BMAR
        `;
        btn.classList.remove('success');
      }, 3000);

    } catch (error) {
      console.error('[SFBR] Error extracting data:', error);
      showNotification('Error extracting data: ' + error.message, 'error');
    }
  }

  // ── Notification ────────────────────────────────────────────────────────────

  function showNotification(message, type) {
    const existing = document.querySelector('.seafarer-notification');
    if (existing) existing.remove();
    const n = document.createElement('div');
    n.className = `seafarer-notification ${type}`;
    n.textContent = message;
    document.body.appendChild(n);
    setTimeout(() => { n.classList.add('fade-out'); setTimeout(() => n.remove(), 300); }, 4000);
  }

  // ── Boot ────────────────────────────────────────────────────────────────────

  init();
  setTimeout(init, 800);   // second pass for slow-rendering SPAs

  // Re-init on SPA navigation so the button survives page transitions
  const _origPush    = history.pushState.bind(history);
  const _origReplace = history.replaceState.bind(history);
  history.pushState    = function (...a) { _origPush(...a);    setTimeout(init, 600); };
  history.replaceState = function (...a) { _origReplace(...a); setTimeout(init, 600); };
  window.addEventListener('popstate', () => setTimeout(init, 600));
})();
