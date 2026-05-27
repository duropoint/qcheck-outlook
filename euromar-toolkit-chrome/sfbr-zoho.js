// Content script for Zoho BMAR Form
// Automatically fills form fields with data from Seafarers Panel

(function() {
  'use strict';

  // Field mapping based on DIAGNOSTIC output from Zoho form
  // Format: 'ZohoFieldName': 'seafarerDataProperty'
  const FIELD_MAPPING = {
    // === BASIC INFO ===
    'SingleLine': 'ship',                    // Ship
    // Name is handled separately (First + Last) - both have name="Name"
    'Dropdown1': 'sex',                      // Sex (Female, Male)
    'Date': 'birthdate',                     // Birthdate
    'Dropdown2': 'countryOfOrigin',          // Country of origin
    'SingleLine1': 'passportNumber',         // Passport Number
    'Date2': 'passportValidity',             // Passport Validity
    
    // === COMPANY DATA ===
    'MultiLine': 'invoiceRecipient',         // Invoice Address
    
    // === MEDICAL ===
    'Date4': 'medicalIssuanceDate',          // Medical issuance date
    'Date5': 'medicalExpiryDate',            // Medical expiry date
    
    // === COC ===
    'SingleLine4': 'cocNumber',              // COC Number
    'Date11': 'cocIssuanceDate',             // COC Issuance date
    'Date12': 'cocExpiryDate',               // COC Expiry date
    'Date33': 'cocRevalidationIssuanceDate', // COC Revalidation date
    'Date34': 'cocRevalidationExpiryDate',   // COC Revalidation Expiry date
    'Dropdown8': 'cocIssuedBy',              // COC Issued by
    
    // === COC ENDORSEMENT ===
    'SingleLine5': 'cocEndorsementNumber',   // COC Endorsement Number
    'Date13': 'cocEndorsementIssuanceDate',  // COC Endorsement Issuance date
    'Date14': 'cocEndorsementExpiryDate',    // COC Endorsement Expiry date
    'Dropdown9': 'cocEndorsementIssuedBy',   // COC Endorsement Issued by
    
    // === GOC ===
    'SingleLine6': 'gocNumber',              // GOC Number
    'Date15': 'gocIssuanceDate',             // GOC Issuance date
    'Date16': 'gocExpiryDate',               // GOC Expiry date
    'Date36': 'gocRevalidationIssuanceDate', // GOC Revalidation date
    'Date35': 'gocRevalidationExpiryDate',   // GOC Revalidation expiry date
    'Dropdown10': 'gocIssuedBy',             // GOC Issued by
    
    // === GOC ENDORSEMENT ===
    'SingleLine7': 'gocEndorsementNumber',   // GOC Endorsement Number
    'Date17': 'gocEndorsementIssuanceDate',  // GOC Endorsement Issuance date
    'Date18': 'gocEndorsementExpiryDate',    // GOC Endorsement Expiry date
    'Dropdown11': 'gocEndorsementIssuedBy',  // GOC Endorsement Issued by
    
    // === COP-1 ===
    'SingleLine8': 'cop1Number',             // COP-1 Number
    'Date19': 'cop1IssuanceDate',            // COP-1 Issuance date
    'Date20': 'cop1ExpiryDate',              // COP-1 Expiry date
    'Dropdown12': 'cop1IssuedBy',            // COP-1 Issued by
    
    // === COP-2 ===
    'SingleLine9': 'cop2Number',             // COP-2 Number
    'Date21': 'cop2IssuanceDate',            // COP-2 Issuance date
    'Date22': 'cop2ExpiryDate',              // COP-2 Expiry date
    'Dropdown13': 'cop2IssuedBy',            // COP-2 Issued by
  };

  let seafarerData = null;

  function init() {
    // Check for data in URL hash
    const hash = window.location.hash;
    if (hash && hash.includes('seafarerData=')) {
      try {
        const encodedData = hash.split('seafarerData=')[1];
        const jsonString = decodeURIComponent(atob(encodedData));
        seafarerData = JSON.parse(jsonString);
        
        console.log('[Zoho Bridge] Found seafarer data:', seafarerData.fullName);
        console.log('[Zoho Bridge] Data:', JSON.stringify(seafarerData, null, 2));
        
        // Clean the URL (remove the hash) to avoid confusion
        history.replaceState(null, '', window.location.pathname + window.location.search);
        
        // Show fill button/banner
        showFillBanner();
      } catch (e) {
        console.error('[Zoho Bridge] Error parsing data from URL:', e);
      }
    }
  }

  function showFillBanner() {
    // Remove existing banner
    const existing = document.querySelector('.zoho-fill-banner');
    if (existing) existing.remove();

    const banner = document.createElement('div');
    banner.className = 'zoho-fill-banner';
    banner.innerHTML = `
      <div class="banner-content">
        <div class="banner-info">
          <strong>📋 Dados disponíveis:</strong> ${seafarerData.fullName}
          <span class="banner-details">(${seafarerData.countryOfOrigin || 'País não especificado'})</span>
        </div>
        <div class="banner-actions">
          <button id="fill-form-btn" class="banner-btn fill-btn">
            ✨ Preencher Formulário
          </button>
          <button id="diagnose-btn" class="banner-btn" style="background:#17a2b8;">
            🔍 Diagnosticar Campos
          </button>
          <button id="clear-data-btn" class="banner-btn clear-btn">
            ✕ Limpar
          </button>
        </div>
      </div>
    `;

    document.body.insertBefore(banner, document.body.firstChild);

    // Add event listeners
    document.getElementById('fill-form-btn').addEventListener('click', fillForm);
    document.getElementById('diagnose-btn').addEventListener('click', diagnoseFields);
    document.getElementById('clear-data-btn').addEventListener('click', clearData);
  }

  function diagnoseFields() {
    console.log('[Zoho Bridge] ========== DIAGNÓSTICO DE CAMPOS ==========');
    
    // Get all form fields
    const inputs = document.querySelectorAll('input, select, textarea');
    let fieldIndex = 0;
    
    inputs.forEach((el) => {
      const name = el.name || '(sem nome)';
      const id = el.id || '(sem id)';
      const type = el.type || el.tagName.toLowerCase();
      const label = getFieldLabel(el);
      
      // Skip hidden fields
      if (type === 'hidden') return;
      
      fieldIndex++;
      
      if (el.tagName.toLowerCase() === 'select') {
        const options = Array.from(el.options)
          .filter(o => o.text.trim() !== '')
          .map(o => o.text)
          .slice(0, 5);
        console.log(`${fieldIndex}. [SELECT] name="${name}" | label="${label}" | options: ${options.join(', ')}${el.options.length > 5 ? '...' : ''}`);
      } else {
        const value = el.value || '(vazio)';
        const placeholder = el.placeholder || '';
        console.log(`${fieldIndex}. [${type.toUpperCase()}] name="${name}" | label="${label}" | placeholder="${placeholder}" | value="${value}"`);
      }
    });
    
    console.log('[Zoho Bridge] ==========================================');
    console.log('[Zoho Bridge] Copia este output e envia ao Claude para corrigir o mapeamento!');
    
    showNotification('Diagnóstico completo - ver consola (F12)', 'info');
  }

  function fillForm() {
    if (!seafarerData) {
      showNotification('Sem dados para preencher', 'error');
      return;
    }

    let filledCount = 0;
    const errors = [];

    // Fill Name fields first (special handling - full name goes to First only)
    const nameResult = fillNameFields();
    if (nameResult) filledCount += nameResult;

    // Fill all mapped fields
    for (const [fieldName, dataKey] of Object.entries(FIELD_MAPPING)) {
      const value = seafarerData[dataKey];
      
      if (!value || value === '') {
        continue;
      }

      try {
        const filled = fillField(fieldName, value);
        if (filled) {
          filledCount++;
          console.log(`[Zoho Bridge] ✓ Filled ${fieldName} (${dataKey}) with "${value}"`);
        } else {
          console.log(`[Zoho Bridge] ✗ Could not fill ${fieldName} (${dataKey}) with "${value}"`);
        }
      } catch (e) {
        errors.push(`${fieldName}: ${e.message}`);
        console.error(`[Zoho Bridge] ❌ Error filling ${fieldName}:`, e);
      }
    }

    // Update banner
    const btn = document.getElementById('fill-form-btn');
    if (btn) {
      btn.innerHTML = `✓ ${filledCount} campos preenchidos`;
      btn.classList.add('success');
    }

    showNotification(`Formulário preenchido com ${filledCount} campos`, 'success');

    if (errors.length > 0) {
      console.warn('[Zoho Bridge] Some fields could not be filled:', errors);
    }
  }

  function fillField(fieldName, value) {
    // Try multiple selector strategies
    const selectors = [
      `[name="${fieldName}"]`,
      `#${fieldName}`,
      `input[name="${fieldName}"]`,
      `select[name="${fieldName}"]`,
      `textarea[name="${fieldName}"]`,
      `[data-name="${fieldName}"]`,
    ];

    let element = null;
    for (const selector of selectors) {
      try {
        element = document.querySelector(selector);
        if (element) break;
      } catch (e) {
        // Invalid selector, continue
      }
    }

    if (!element) {
      console.log(`[Zoho Bridge] Field "${fieldName}" not found in DOM`);
      return false;
    }

    // Handle different input types
    const tagName = element.tagName.toLowerCase();

    if (tagName === 'select') {
      return fillSelect(element, value);
    } else if (tagName === 'input' || tagName === 'textarea') {
      return fillInput(element, value);
    }

    return false;
  }

  function fillInput(element, value) {
    // Focus the element first
    element.focus();
    
    // Clear existing value
    element.value = '';
    
    // Set new value
    element.value = value;
    
    // Trigger all necessary events for Zoho to recognize the change
    const events = ['input', 'change', 'blur', 'keyup'];
    events.forEach(eventType => {
      element.dispatchEvent(new Event(eventType, { bubbles: true }));
    });

    // Also try input event with data
    element.dispatchEvent(new InputEvent('input', {
      bubbles: true,
      inputType: 'insertText',
      data: value
    }));

    return true;
  }

  function fillSelect(element, value) {
    const options = Array.from(element.options);
    const fieldName = element.name || element.id || 'unknown';
    
    console.log(`[Zoho Bridge] Trying to select "${value}" in dropdown ${fieldName}`);
    
    // Filter out empty options
    const validOptions = options.filter(o => o.value && o.value.trim() !== '' && o.text && o.text.trim() !== '');
    
    // Normalize value for comparison
    const normalizeStr = (str) => str.toLowerCase().trim().replace(/[^a-z0-9]/g, '');
    const valueNorm = normalizeStr(value);
    
    // 1. Exact match (case-insensitive)
    let option = validOptions.find(o => 
      normalizeStr(o.value) === valueNorm || 
      normalizeStr(o.text) === valueNorm
    );
    
    // 2. Starts with match
    if (!option) {
      option = validOptions.find(o => 
        normalizeStr(o.text).startsWith(valueNorm) ||
        valueNorm.startsWith(normalizeStr(o.text))
      );
    }
    
    // 3. Contains match
    if (!option) {
      option = validOptions.find(o => 
        normalizeStr(o.text).includes(valueNorm) || 
        valueNorm.includes(normalizeStr(o.text))
      );
    }
    
    // 4. Word-based match (for country names with different formats)
    if (!option) {
      const valueWords = value.toLowerCase().split(/\s+/);
      option = validOptions.find(o => {
        const optionText = o.text.toLowerCase();
        return valueWords.some(word => word.length > 2 && optionText.includes(word));
      });
    }

    if (option) {
      element.value = option.value;
      element.dispatchEvent(new Event('change', { bubbles: true }));
      element.dispatchEvent(new Event('input', { bubbles: true }));
      console.log(`[Zoho Bridge] ✓ Selected "${option.text}" for value "${value}"`);
      return true;
    }

    // DO NOT select empty option - just fail
    console.log(`[Zoho Bridge] ✗ No matching option for "${value}" in ${fieldName}. NOT selecting anything.`);
    return false;
  }

  function fillNameFields() {
    let filled = 0;
    
    // Strategy 1: Find by name attribute "Name" (Zoho uses this for name fields)
    const nameInputs = document.querySelectorAll('input[name="Name"]');
    
    if (nameInputs.length >= 1) {
      // Put FULL NAME in First Name field only, leave Last Name empty
      if (seafarerData.fullName) {
        fillInput(nameInputs[0], seafarerData.fullName);
        filled++;
        console.log(`[Zoho Bridge] ✓ Filled First Name with full name: "${seafarerData.fullName}"`);
      }
      // Last Name (nameInputs[1]) intentionally left empty
      return filled;
    }

    // Strategy 2: Find by placeholder text - only fill First name field
    const allInputs = document.querySelectorAll('input[type="text"]');
    allInputs.forEach(input => {
      const placeholder = (input.placeholder || '').toLowerCase();
      const label = getFieldLabel(input).toLowerCase();
      
      if ((placeholder.includes('first') || label.includes('first')) && seafarerData.fullName) {
        fillInput(input, seafarerData.fullName);
        filled++;
      }
      // Last name field intentionally not filled
    });

    return filled;
  }

  function getFieldLabel(input) {
    // Try to find associated label
    const id = input.id;
    if (id) {
      const label = document.querySelector(`label[for="${id}"]`);
      if (label) return label.textContent;
    }
    
    // Try parent elements
    let parent = input.parentElement;
    for (let i = 0; i < 3 && parent; i++) {
      const label = parent.querySelector('label');
      if (label) return label.textContent;
      parent = parent.parentElement;
    }
    
    return '';
  }

  function clearData() {
    seafarerData = null;
    const banner = document.querySelector('.zoho-fill-banner');
    if (banner) {
      banner.classList.add('fade-out');
      setTimeout(() => banner.remove(), 300);
    }
    showNotification('Dados limpos', 'info');
  }

  function showNotification(message, type) {
    const existing = document.querySelector('.seafarer-notification');
    if (existing) existing.remove();

    const notification = document.createElement('div');
    notification.className = `seafarer-notification ${type}`;
    notification.textContent = message;
    document.body.appendChild(notification);

    setTimeout(() => {
      notification.classList.add('fade-out');
      setTimeout(() => notification.remove(), 300);
    }, 4000);
  }

  // Initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    // Wait a bit for Zoho form to fully render
    setTimeout(init, 1500);
  }
})();
