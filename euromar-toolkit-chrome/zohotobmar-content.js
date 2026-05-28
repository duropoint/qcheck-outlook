// BMAR Auto Complete v2.1.1 — content script for BMAR certification portal.
// Listens for automation commands from the toolkit side panel (via popup.js)
// and drives the BMAR multi-step form: Step 2 (applicant), Step 3 (certificates
// + medical), Step 4 (document uploads). Progress events are sent back via
// chrome.runtime.sendMessage so popup.js can forward them to the iframe.

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === "ping") {
    sendResponse({ status: "pong" });
    return true;
  }
  if (message.action === "automate") {
    executeFullAutomation(message.fields, message.documents, message.options || {});
    sendResponse({ status: "started" });
    return true;
  }
  if (message.action === "uploadOnly") {
    sendProgress(75, "Retrying failed uploads…");
    uploadDocuments(message.documents);
    sendResponse({ status: "started" });
    return true;
  }
  return true;
});

async function executeFullAutomation(fields, documents, options) {
  try {
    sendProgress(5, "Checking current page step…");

    const currentStep = detectCurrentStep();

    // Redirect 04_coc_* file to tanker category when there is no COC number in the data
    const hasCOCNumber = fields["COC Number"] && fields["COC Number"].trim() !== "";
    if (!hasCOCNumber && documents.coc) {
      documents.tankerCOC = documents.coc;
      delete documents.coc;
    }

    if (currentStep === 2) {
      sendProgress(10, "Filling Step 2 (Applicant)…");
      await fillStep2(fields);
      await wait(1000);
      sendProgress(30, "Advancing to Step 3…");
      await clickNextButton();
      await wait(2000);
    }

    if (currentStep === 3 || currentStep === 2) {
      sendProgress(35, "Filling Step 3 (Characterization)…");
      await fillStep3(fields, options);
      await wait(1000);
      sendProgress(70, "Advancing to Step 4…");
      await clickNextButton();
      await wait(3000);
    }

    if (currentStep === 4 || currentStep === 3 || currentStep === 2) {
      sendProgress(75, "Uploading documents…");
      await uploadDocuments(documents);
      sendProgress(100, "Automation complete.");
    }

    chrome.runtime.sendMessage({ type: "bmar-complete", _fromContent: true });
  } catch (error) {
    console.error("[BMAR] Automation error:", error);
    chrome.runtime.sendMessage({ type: "bmar-error", message: error.message, _fromContent: true });
  }
}

function detectCurrentStep() {
  const steps = document.querySelectorAll('[class*="step"], [class*="Step"]');
  for (const step of steps) {
    if (step.textContent.includes("4") && step.className.includes("active")) return 4;
    if (step.textContent.includes("3") && step.className.includes("active")) return 3;
    if (step.textContent.includes("2") && step.className.includes("active")) return 2;
  }
  if (document.querySelector('label[id*="AnexoBean"]'))         return 4;
  if (document.querySelector('button[id*="EndorsmentDetail"]')) return 3;
  if (document.querySelector('input[id*="PedidoPessoaTit"]'))   return 2;
  return 2;
}

async function fillStep2(fields) {
  const mappings = {
    "Birthdate":        'input[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_datNascimento_fld_input"]',
    "First Name":       'input[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_nomPessoa_fld"]',
    "Passport Number":  'input[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_numPassaporte_fld"]',
    "Passport Validity":'input[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_datValidadePassap_fld_input"]'
  };
  const dropdownMappings = {
    "Sex": {
      select: 'select[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_codGenero_fld_input"]',
      label:  'label[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_codGenero_fld_label"]'
    },
    "Country of Origin": {
      select: 'select[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_xfkPaisNacionalid_fld_input"]',
      label:  'label[id="detailFormId:PedidoPessoaTit:PedidoPessoaTitBean_xfkPaisNacionalid_fld_label"]'
    },
    "Invoice": {
      select: 'select[id="detailFormId:PedidoRequerenteBean_codDestFatura_fld_input"]',
      label:  'label[id="detailFormId:PedidoRequerenteBean_codDestFatura_fld_label"]'
    }
  };

  for (const [key, selector] of Object.entries(mappings)) {
    if (fields[key]) { pasteInput(selector, fields[key]); await wait(200); }
  }
  for (const [key, sels] of Object.entries(dropdownMappings)) {
    if (fields[key] && fields[key] !== "-Select-") {
      pasteDropdown(sels.select, sels.label, fields[key]);
      await wait(200);
    }
  }
}

async function fillStep3(fields, options = {}) {
  await fillMedicalInfo(fields, options);
  await wait(500);

  const hasCOC  = !!(fields["COC Number"]  && fields["COC Number"].trim());
  const hasGOC  = !!(fields["GOC Number"]  && fields["GOC Number"].trim());
  const hasCOP1 = !!(fields["COP1 Number"] && fields["COP1 Number"].trim());
  const hasCOP2 = !!(fields["COP2 Number"] && fields["COP2 Number"].trim());

  if (hasCOC)  { await addCertificateWithRules("STCW",          fields, "COC",  options); await wait(1000); }
  if (hasGOC)  { await addCertificateWithRules("STCW (GMDSS)",  fields, "GOC",  options); await wait(1000); }
  if (hasCOP1) { await addCertificateWithRules("STCW (Tankers)", fields, "COP1", options); await wait(1000); }
  if (hasCOP2) { await addCertificateWithRules("STCW (Tankers)", fields, "COP2", options); await wait(1000); }
}

async function fillMedicalInfo(fields, options = {}) {
  if (fields["Medical Issuance date"]) {
    pasteInput('input[id="detailFormId:EndorsmentDetail:EndorsmentBean_datEmissaoDm_fld_input"]', fields["Medical Issuance date"]);
    await wait(500);
  }
  if (fields["Medical Expiry"]) {
    pasteInput('input[id="detailFormId:EndorsmentDetail:EndorsmentBean_datValidadeDm_fld_input"]', fields["Medical Expiry"]);
    await wait(500);
  }
  if (options.medicalNationality && options.medicalNationality.trim()) {
    const input = document.querySelector('input[id="detailFormId:EndorsmentDetail:EndorsmentBean_txtLocEmissaoDm_fld"]');
    if (input) {
      input.focus();
      input.value = options.medicalNationality;
      input.dispatchEvent(new Event("input",  { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      input.dispatchEvent(new Event("blur",   { bubbles: true }));
    }
    await wait(500);
  }
  if (fields["Ship"]) {
    pasteInput('input[id="detailFormId:EndorsmentDetail:EndorsmentBean_txtLocEmiDeclar_fld"]', fields["Ship"]);
    await wait(300);
  }

  // Auto-select "Digital Document"
  const digitalDocSelect = document.querySelector('select[id="detailFormId:EndorsmentDetail:EndorsmentBean_flgDocFisico_fld_input"]');
  const digitalDocLabel  = document.querySelector('label[id="detailFormId:EndorsmentDetail:EndorsmentBean_flgDocFisico_fld_label"]');
  if (digitalDocSelect) {
    const option = Array.from(digitalDocSelect.options).find(o => o.value === "NAO");
    if (option) {
      digitalDocSelect.value = option.value;
      digitalDocSelect.dispatchEvent(new Event("change", { bubbles: true }));
      if (digitalDocLabel) digitalDocLabel.textContent = "Digital Document";
      await wait(300);
    }
  }
}

async function addCertificateWithRules(certType, fields, prefix, options) {
  await addCertificateOnly(certType, fields, prefix);

  let shouldAddRule = false;
  let ruleToAdd     = null;
  let capacityToAdd = null;

  if (prefix === "COC" && options.addRuleCapacity && options.stcwRule && options.capacity) {
    shouldAddRule = true;
    ruleToAdd     = options.stcwRule;
    capacityToAdd = options.capacity;
  }
  if (prefix === "GOC") {
    shouldAddRule = true;
    ruleToAdd     = "IV/2";
    capacityToAdd = "GMDSS Radio Operator";
  }
  if (prefix === "COP1" && options.addTankerCapacity && options.cop1Rule && options.tankerCapacity) {
    shouldAddRule = true;
    ruleToAdd     = options.cop1Rule;
    capacityToAdd = options.tankerCapacity;
  }
  if (prefix === "COP2" && options.addTankerCapacity && options.cop2Rule && options.tankerCapacity) {
    shouldAddRule = true;
    ruleToAdd     = options.cop2Rule;
    capacityToAdd = options.tankerCapacity;
  }

  if (shouldAddRule) await addRuleInsideModal(ruleToAdd, capacityToAdd, prefix);

  await wait(500);

  const confirmBtn = Array.from(document.querySelectorAll("button")).find(b =>
    (b.innerText.includes("Confirm") || b.innerText.includes("Confirmar")) &&
    b.id.includes("detailFormId") &&
    !b.id.includes("EndorsCertifRegra")
  );
  if (confirmBtn) {
    confirmBtn.click();
    await wait(3000);
  }
}

async function addCertificateOnly(certType, fields, prefix) {
  sendProgress(
    40 + (prefix === "COC" ? 0 : prefix === "GOC" ? 10 : prefix === "COP1" ? 20 : 30),
    `Adding ${prefix} certificate…`
  );

  const addButton = Array.from(document.querySelectorAll("button")).find(btn =>
    btn.id.includes("EndorsmentDetail") &&
    (btn.innerText.includes("Add New") || btn.innerText.includes("Adicionar"))
  );
  if (!addButton) { console.error("[BMAR] Add New button not found"); return; }
  addButton.click();
  await wait(2000);

  const typeDropdown = document.querySelector('select[id*="xfkTpCertif_fld_input"]');
  const typeLabel    = document.querySelector('label[id*="xfkTpCertif_fld_label"]');
  if (typeDropdown) {
    const option = Array.from(typeDropdown.options).find(opt => {
      const t = opt.text.trim();
      if (certType === "STCW"         && t === "STCW") return true;
      if (certType === "STCW (GMDSS)" && (t.includes("GMDSS") || t.includes("Radiocommunications"))) return true;
      if (certType === "STCW (Tankers)" && t.includes("Tanker")) return true;
      return t.includes(certType);
    });
    if (option) {
      typeDropdown.value = option.value;
      typeDropdown.dispatchEvent(new Event("change", { bubbles: true }));
      if (typeLabel) typeLabel.textContent = option.text;
      await wait(1000);
    }
  }

  const certMappings = {
    [`${prefix} Number`]:               'input[id="detailFormId:EndorsmentCertifBean_numCertificado_fld"]',
    [`${prefix} Issuance`]:             'input[id="detailFormId:EndorsmentCertifBean_datEmissaoCert_fld_input"]',
    [`${prefix} Expiry`]:               'input[id="detailFormId:EndorsmentCertifBean_datValidadeCert_fld_input"]',
    [`${prefix} Endorsement Number`]:   'input[id="detailFormId:EndorsmentCertifBean_numEndorsment_fld"]',
    [`${prefix} Endorsement Issuance`]: 'input[id="detailFormId:EndorsmentCertifBean_datEmissaoEndors_fld_input"]',
    [`${prefix} Endorsement Expiry`]:   'input[id="detailFormId:EndorsmentCertifBean_datValidadeEndors_fld_input"]',
    [`${prefix} Revalidation date`]:    'input[id="detailFormId:EndorsmentCertifBean_datRevalidCert_fld_input"]',
    [`${prefix} Revalidation Expiry date`]: 'input[id="detailFormId:EndorsmentCertifBean_datValRevalCert_fld_input"]'
  };
  for (const [key, sel] of Object.entries(certMappings)) {
    if (fields[key]) { pasteInput(sel, fields[key]); await wait(300); }
  }

  if (fields[`${prefix} Issued By`]) {
    pasteDropdown(
      'select[id="detailFormId:EndorsmentCertifBean_xfkPaisEmiCert_fld_input"]',
      'label[id="detailFormId:EndorsmentCertifBean_xfkPaisEmiCert_fld_label"]',
      fields[`${prefix} Issued By`]
    );
    await wait(300);
  }
  if (fields[`${prefix} Endorsement Issued By`]) {
    pasteDropdown(
      'select[id="detailFormId:EndorsmentCertifBean_xfkPaisEmiEndors_fld_input"]',
      'label[id="detailFormId:EndorsmentCertifBean_xfkPaisEmiEndors_fld_label"]',
      fields[`${prefix} Endorsement Issued By`]
    );
    await wait(300);
  }
}

async function addRuleInsideModal(stcwRule, capacity, certType) {
  try {
    await wait(2000);

    const modal = document.querySelector('.ui-dialog[aria-hidden="false"]');
    if (modal) {
      const modalContent = modal.querySelector('.ui-dialog-content');
      if (modalContent) { modalContent.scrollTop = modalContent.scrollHeight; await wait(1000); }
    }

    let addRuleBtn = null;
    const rulesSection = Array.from(document.querySelectorAll('.ui-fieldset-legend, legend, .ui-panel-title'))
      .find(el => {
        const t = el.textContent.trim().toLowerCase();
        return t.includes("certificate's rules") || t.includes("rules and capacities") || t.includes("stcw rule");
      });

    if (rulesSection) {
      const container = rulesSection.closest('fieldset, .ui-panel, .ui-fieldset');
      if (container) {
        const btn = container.querySelector('button');
        if (btn && btn.innerText.includes("Add New")) addRuleBtn = btn;
      }
    }

    if (!addRuleBtn) {
      for (const fieldset of document.querySelectorAll('fieldset')) {
        const legend = fieldset.querySelector('legend');
        if (legend && legend.textContent.toLowerCase().includes('rule')) {
          const btn = fieldset.querySelector('button');
          if (btn && btn.innerText.includes("Add New") && btn.offsetParent !== null) {
            addRuleBtn = btn;
            break;
          }
        }
      }
    }

    if (!addRuleBtn) return;
    addRuleBtn.click();
    await wait(2000);
    await addRuleAndCapacity(stcwRule, capacity, certType);
  } catch (error) {
    console.error("[BMAR] Error adding rule:", error);
  }
}

async function addRuleAndCapacity(stcwRule, capacity, certType) {
  const ruleTextMap = {
    "II/1":   "Rule II/1 - Officer in charge of a navigational watch on ships of 500 GT or more",
    "II/2":   "Rule II/2 - Master and Chief Mate on ships of 500 GT or more",
    "II/3":   "Rule II/3 - Officer in charge of a navigational watch on ships of less than 500 GT, engaged on near coastal voyages",
    "III/1":  "Rule III/1 - Officer in Charge of an Engineering Watch",
    "III/2":  "Rule III/2 - Chief Engineer Officer and Second Engineer Officer on ships powered by main propulsion machinery of 3000 KW propulsion power or more",
    "III/3":  "Rule III/3 - Chief Engineer Officer and Second Engineer Officeron ships powered by main propulsion machinery of between 750 KW and 3000 KW propulsion power)",
    "III/6":  "Rule III/6 - Electrotechnical Officer",
    "IV/2":   "Rule IV/2 - GMDSS Radio Operator",
    "V/1-1-1":"Rule V/1-1-1 - Basic training for oil and chemical tanker cargo operations",
    "V/1-1-2":"Rule V/1-1-2 - Advanced training for oil tanker cargo operations",
    "V/1-1-3":"Rule V/1-1-3 - Advanced training for chemical tanker cargo operations",
    "V/1-2-1":"Rule V/1-2-1 - Basic training for liquified gas tanker cargo operations",
    "V/1-2-2":"Rule V/1-2-2 - Advanced training for liquified gas tanker cargo operations"
  };

  try {
    await wait(1000);
    const ruleFullText = ruleTextMap[stcwRule];
    const ruleSelect = document.querySelector('select[id*="xfkRegraStcw_fld_input"]');
    const ruleLabel  = document.querySelector('label[id*="xfkRegraStcw_fld_label"]');

    if (ruleSelect && ruleFullText) {
      const option = Array.from(ruleSelect.options).find(o =>
        o.text.includes(ruleFullText) || o.text.includes(stcwRule) ||
        (stcwRule.startsWith("V/1") && o.text.includes(stcwRule))
      );
      if (option) {
        ruleSelect.value = option.value;
        ruleSelect.dispatchEvent(new Event("change", { bubbles: true }));
        if (ruleLabel) ruleLabel.textContent = option.text;
        await wait(1000);
      } else {
        return;
      }
    } else {
      return;
    }

    const capacitySelect = document.querySelector('select[id*="xfkCargoStcw_fld_input"]');
    const capacityLabel  = document.querySelector('label[id*="xfkCargoStcw_fld_label"]');
    if (capacitySelect && capacity) {
      const option = Array.from(capacitySelect.options).find(o =>
        o.text.trim() === capacity || o.text.includes(capacity)
      );
      if (option) {
        capacitySelect.value = option.value;
        capacitySelect.dispatchEvent(new Event("change", { bubbles: true }));
        if (capacityLabel) capacityLabel.textContent = option.text;
        await wait(500);
      } else {
        return;
      }
    } else {
      return;
    }

    await wait(500);
    const confirmRuleBtn = Array.from(document.querySelectorAll("button")).find(b =>
      (b.innerText.includes("Confirm") || b.innerText.includes("Confirmar")) &&
      b.id.match(/j_idt\d{4,5}/) &&
      b.offsetParent !== null
    );
    if (confirmRuleBtn) {
      confirmRuleBtn.click();
      await wait(1500);
    }
  } catch (error) {
    console.error("[BMAR] Error adding rule & capacity:", error);
  }
}

async function uploadDocuments(documents) {
  await wait(1500);

  const labelMap = {
    seamans:   "Seaman's Book",
    coc:       "Certificate of Competency",
    gmdss:     "Radiocommunications Certificate",
    med:       "Medical Certificate",
    rok:       "Statement",
    tanker:    "Certificate of Proficiency",
    tankerChem:"Certificate of Proficiency",
    tankerOil: "Certificate of Proficiency",
    tankerCOC: "Certificate of Proficiency",
    pass:      "Passport or Other Seafarer Identity Document",
    loc:       "Other document"
  };

  const totalDocs = Object.keys(documents).filter(k => documents[k]).length;
  let uploadCount = 0;
  const uploadResults = { success: [], failed: [] };

  for (const [tipo, ficheiro] of Object.entries(documents)) {
    if (!ficheiro) continue;
    uploadCount++;
    sendProgress(
      75 + (uploadCount * Math.floor(20 / totalDocs)),
      `Uploading ${uploadCount}/${totalDocs}: ${ficheiro.name}`
    );
    try {
      const ok = await uploadSingleDocument(tipo, ficheiro, labelMap[tipo]);
      if (ok) uploadResults.success.push({ name: ficheiro.name, type: tipo });
      else     uploadResults.failed.push({  name: ficheiro.name, type: tipo, reason: "Upload failed" });
    } catch (error) {
      uploadResults.failed.push({ name: ficheiro.name, type: tipo, reason: error.message });
    }
    await wait(300);
  }

  chrome.runtime.sendMessage({ type: "bmar-upload-complete", results: uploadResults, _fromContent: true });
}

async function uploadSingleDocument(tipo, ficheiro, labelText) {
  let addBtn = null;
  const label = Array.from(document.querySelectorAll("label.ui-outputlabel"))
    .find(l => l.textContent.trim() === labelText);

  if (label) {
    const panelContent = label.closest(".ui-panel-content");
    if (panelContent) {
      const buttons = Array.from(panelContent.querySelectorAll("button"));
      addBtn = buttons.find(b => b.innerText.trim() === "Add New");
    }
  }
  if (!addBtn) throw new Error(`Add New button not found for label "${labelText}"`);

  addBtn.click();
  await wait(1500);

  let fileInput = null;
  for (let i = 0; i < 15 && !fileInput; i++) {
    const allFileInputs = document.querySelectorAll('input[type="file"]');
    fileInput = Array.from(allFileInputs).find(inp => {
      const s = window.getComputedStyle(inp);
      return s.display !== "none" && s.visibility !== "hidden";
    });
    if (!fileInput) {
      fileInput = document.getElementById("detailFormId:AnexoBean_binFicheiro_fld_input");
    }
    if (!fileInput) await wait(500);
  }
  if (!fileInput) throw new Error("File input not found");

  const blob = b64toBlob(ficheiro.data, "application/pdf");
  const file = new File([blob], ficheiro.name, { type: "application/pdf" });
  const dt   = new DataTransfer();
  dt.items.add(file);
  fileInput.files = dt.files;
  fileInput.dispatchEvent(new Event("change", { bubbles: true }));
  fileInput.dispatchEvent(new Event("input",  { bubbles: true }));
  await wait(1500);

  // Check for errors
  const errorMessages = Array.from(document.querySelectorAll('.ui-message-error, .ui-messages-error, .ui-messages-error-detail'));
  const errorTexts    = errorMessages.map(e => e.textContent.trim()).filter(t => t);
  if (errorTexts.length > 0) {
    await dismissAndBack();
    throw new Error(errorTexts.join("; "));
  }

  const visibleButtons = Array.from(document.querySelectorAll("button")).filter(b => b.offsetParent !== null);
  const confirmBtn = visibleButtons.find(b => {
    const t = b.innerText.trim();
    return t === "Confirm" || t === "Confirmar" || t.includes("Confirm") || t.includes("Confirmar");
  });
  if (!confirmBtn) throw new Error("Confirm button not found");
  confirmBtn.click();
  await wait(1500);

  // Post-confirm error check
  await wait(300);
  const postErrors = Array.from(document.querySelectorAll('.ui-message-error, .ui-messages-error'))
    .map(e => e.textContent.trim()).filter(t => t);
  if (postErrors.length > 0) {
    await dismissAndBack();
    throw new Error(postErrors.join("; "));
  }
  return true;
}

async function dismissAndBack() {
  const visibleButtons = Array.from(document.querySelectorAll("button")).filter(b => b.offsetParent !== null);
  const backBtn = visibleButtons.find(b => {
    const t = b.innerText.trim();
    return t === "Back" || t === "Voltar" || t.includes("Back") || t.includes("Voltar");
  });
  if (backBtn) {
    backBtn.click();
    await wait(1000);
    const yesBtn = Array.from(document.querySelectorAll("button")).find(b =>
      b.offsetParent !== null && (b.innerText.trim() === "Yes" || b.innerText.trim() === "Sim")
    );
    if (yesBtn) { yesBtn.click(); await wait(1000); }
    return;
  }
  const cancelBtn = visibleButtons.find(b =>
    b.innerText.includes("Cancel") || b.innerText.includes("Cancelar")
  );
  if (cancelBtn) { cancelBtn.click(); await wait(1000); return; }
  const closeBtn = document.querySelector('button.ui-dialog-titlebar-close, a.ui-dialog-titlebar-close');
  if (closeBtn)  { closeBtn.click();  await wait(1000); }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function pasteInput(selector, value) {
  const field = document.querySelector(selector);
  if (!field) return;
  field.focus();
  field.value = value;
  field.dispatchEvent(new Event("input",  { bubbles: true }));
  field.dispatchEvent(new Event("change", { bubbles: true }));
}

function pasteDropdown(selectSel, labelSel, value) {
  if (value === "-Select-") return;
  const sel = document.querySelector(selectSel);
  const lbl = document.querySelector(labelSel);
  if (!sel) return;
  const option = Array.from(sel.options).find(o => o.text.trim() === value || o.text.trim().includes(value));
  if (option) {
    sel.value = option.value;
    sel.dispatchEvent(new Event("change", { bubbles: true }));
    if (lbl) lbl.textContent = value;
  }
}

async function clickNextButton() {
  const btn = Array.from(document.querySelectorAll("button")).find(b =>
    b.id.includes("j_idt") &&
    (b.innerText.includes("Next") || b.innerText.includes("Seguinte") || b.className.includes("next"))
  );
  if (btn) btn.click();
}

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

function b64toBlob(b64Data, contentType = "", sliceSize = 512) {
  const byteCharacters = atob(b64Data);
  const byteArrays = [];
  for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
    const slice      = byteCharacters.slice(offset, offset + sliceSize);
    const byteNumbers = new Array(slice.length);
    for (let i = 0; i < slice.length; i++) byteNumbers[i] = slice.charCodeAt(i);
    byteArrays.push(new Uint8Array(byteNumbers));
  }
  return new Blob(byteArrays, { type: contentType });
}

function sendProgress(percent, message) {
  chrome.runtime.sendMessage({ type: "bmar-progress", percent, message, _fromContent: true });
}
