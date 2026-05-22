// Q Check — Outlook task pane logic

const DEFAULT_API_BASE = "https://pscplatformalpha.onrender.com";

const SETTING_API_BASE      = "qcheck_api_base";
const SETTING_API_KEY       = "qcheck_api_key";
const SETTING_COMPANIES_KEY = "qcheck_companies_key";
const SETTING_ZAMMAD_URL    = "zammad_url";
const SETTING_ZAMMAD_TOKEN  = "zammad_token";

// ---------- DOM ----------
const $ = (id) => document.getElementById(id);

const toggleWrap        = $("toggleWrap");
const toggleCompanyBtn  = $("toggleCompany");
const toggleVesselBtn   = $("toggleVessel");
const companyFields     = $("companyFields");
const vesselFields      = $("vesselFields");
const formView          = $("formView");
const loadingView       = $("loadingView");
const errorView         = $("errorView");
const errorMsg          = $("errorMsg");
const confirmView       = $("confirmView");
const companyResult     = $("companyResult");
const vesselResult      = $("vesselResult");
const settingsView      = $("settingsView");

const companyImo        = $("companyImo");
const companyName       = $("companyName");
const vesselImo         = $("vesselImo");
const vesselCompanyImo  = $("vesselCompanyImo");
const vesselName        = $("vesselName");
const vesselCompanyName = $("vesselCompanyName");

const ismCollapseBtn    = $("ismCollapseBtn");
const ismCollapseBody   = $("ismCollapseBody");
const ismCollapseIcon   = $("ismCollapseIcon");

const runBtn            = $("runBtn");
const backToFormBtn     = $("backToForm");
const settingsBtn       = $("settingsBtn");
const sendBtn           = $("sendBtn");
const closeSettingsBtn  = $("closeSettings");

const apiBaseInput          = $("apiBase");
const apiKeyInput           = $("apiKey");
const companiesApiKeyInput  = $("companiesApiKey");
const zammadUrlInput        = $("zammadUrl");
const zammadTokenInput      = $("zammadToken");
const saveBtn               = $("saveBtn");
const testBtn               = $("testBtn");
const settingsStatus        = $("settingsStatus");

const ticketView          = $("ticketView");
const ticketSuccessView   = $("ticketSuccessView");
const loadingText         = $("loadingText");
const loadingSub          = $("loadingSub");

let mode          = "company";
let firstRun      = false;
let pendingQCheck = null;

// ---------- Zammad category tree (hardcoded; mirrors Zammad admin config) ----------
const CATEGORY_TREE = {
  "Technical": {
    "Hull & Structure":   ["Corrosion", "Cracks & Fractures", "Coating", "Other"],
    "Main Engine":        ["Breakdown", "Performance Loss", "Maintenance", "Other"],
    "Auxiliary Engine":   ["Breakdown", "Performance Loss", "Maintenance", "Other"],
    "Navigation Systems": ["Radar", "GPS / GNSS", "AIS", "ECDIS", "Other"],
    "Electrical":         ["Power Failure", "Lighting", "Switchboard", "Other"],
    "Pumps & Piping":     ["Leakage", "Failure", "Maintenance", "Other"],
    "Other":              []
  },
  "Operations": {
    "Port State Control": ["Deficiency", "Detention", "Pre-Inspection", "Other"],
    "Voyage":             ["Delay", "Deviation", "Weather Routing", "Other"],
    "Cargo":              ["Damage", "Shortage", "Stowage Issue", "Other"],
    "Bunkering":          ["Quality Issue", "Quantity Dispute", "Delay", "Other"],
    "Other":              []
  },
  "Crew": {
    "Manning":            ["Shortage", "Qualification Issue", "Crew Change", "Other"],
    "Medical":            ["Illness", "Injury", "Medical Evacuation", "Other"],
    "Training":           ["Certificate Renewal", "Drills", "Courses", "Other"],
    "Other":              []
  },
  "Safety & Compliance": {
    "ISM":                ["Non-conformity", "Audit", "Drill", "Other"],
    "ISPS":               ["Security Alert", "Drill", "Audit", "Other"],
    "MARPOL":             ["Pollution Incident", "Oil Record Book", "Garbage", "Other"],
    "MLC":                ["Wages", "Rest Hours", "Accommodation", "Other"],
    "Other":              []
  },
  "Commercial": {
    "Charter Party":      ["Dispute", "Claim", "Amendment", "Other"],
    "Hire & Freight":     ["Payment", "Dispute", "Off-hire", "Other"],
    "Port Costs":         ["Invoice", "Dispute", "Disbursement Account", "Other"],
    "Other":              []
  },
  "Other": {
    "Other":              []
  }
};

// ---------- Office.js init ----------
Office.onReady(() => {
  bindEvents();
  initApp();
  setupAutocomplete({ inputEl: companyImo,       pairedEl: companyName,       searchParam: "imo",  dropdownEl: $("companyImoDropdown") });
  setupAutocomplete({ inputEl: companyName,       pairedEl: companyImo,        searchParam: "name", dropdownEl: $("companyNameDropdown") });
  setupAutocomplete({ inputEl: vesselCompanyImo,  pairedEl: vesselCompanyName, searchParam: "imo",  dropdownEl: $("vesselCompanyImoDropdown") });
  setupAutocomplete({ inputEl: vesselCompanyName, pairedEl: vesselCompanyImo,  searchParam: "name", dropdownEl: $("vesselCompanyNameDropdown") });
  // Paste buttons
  document.querySelectorAll(".paste-btn").forEach(btn => {
    btn.addEventListener("mousedown", e => e.preventDefault()); // keep input focused
    btn.addEventListener("click", () => handlePaste($(btn.dataset.target)));
  });
});

// ---------- App init ----------
// Gate: all three settings (URL + both keys) must be saved before accessing the app.
function initApp() {
  const base         = getSetting(SETTING_API_BASE, "");
  const key          = getSetting(SETTING_API_KEY, "");
  const companiesKey = getSetting(SETTING_COMPANIES_KEY, "");

  if (!base || !key || !companiesKey) {
    firstRun = true;
    toggleWrap.classList.add("hidden");
    settingsBtn.classList.add("hidden");
    sendBtn.classList.add("hidden");
    closeSettingsBtn.classList.add("hidden");
    apiBaseInput.value         = getSetting(SETTING_API_BASE, DEFAULT_API_BASE);
    apiKeyInput.value          = getSetting(SETTING_API_KEY, "");
    companiesApiKeyInput.value = getSetting(SETTING_COMPANIES_KEY, "");
    zammadUrlInput.value       = getSetting(SETTING_ZAMMAD_URL, "");
    zammadTokenInput.value     = getSetting(SETTING_ZAMMAD_TOKEN, "");
    settingsStatus.textContent = "";
    settingsStatus.className   = "";
    showView(settingsView);
  } else {
    closeSettingsBtn.classList.remove("hidden");
    sendBtn.classList.remove("hidden");
    showView(formView);
  }
}

// ---------- Event binding ----------
function bindEvents() {
  toggleCompanyBtn.addEventListener("click", () => setMode("company"));
  toggleVesselBtn.addEventListener("click", () => setMode("vessel"));
  ismCollapseBtn.addEventListener("click", toggleIsmCollapse);
  runBtn.addEventListener("click", runQCheck);
  backToFormBtn.addEventListener("click", () => showView(formView));
  settingsBtn.addEventListener("click", openSettings);
  sendBtn.addEventListener("click", openTicketForm);
  closeSettingsBtn.addEventListener("click", () => showView(formView));
  saveBtn.addEventListener("click", saveSettings);
  testBtn.addEventListener("click", testConnection);
  $("submitTicketBtn").addEventListener("click", submitTicket);
  $("cancelTicketBtn").addEventListener("click", () => showView(formView));
  $("ticketBackBtn").addEventListener("click", () => showView(formView));
  $("ticketCatL1").addEventListener("change", updateCatL2);
  $("ticketCatL2").addEventListener("change", updateCatL3);
  $("confirmProceedBtn").addEventListener("click", () => {
    if (pendingQCheck) { const fn = pendingQCheck; pendingQCheck = null; fn(); }
  });
  $("confirmCancelBtn").addEventListener("click", () => {
    pendingQCheck = null;
    showView(formView);
  });
}

// ---------- Paste helper ----------
async function handlePaste(inputEl) {
  if (!inputEl) return;
  try {
    const text = await navigator.clipboard.readText();
    if (text) {
      inputEl.value = text.trim();
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
    }
  } catch (_) {
    // Clipboard access denied — focus the field so user can paste with keyboard
  }
  inputEl.focus();
}

// ---------- ISM Company collapse ----------
function toggleIsmCollapse() {
  const expanded = !ismCollapseBody.classList.contains("hidden");
  if (expanded) {
    ismCollapseBody.classList.add("hidden");
    ismCollapseIcon.textContent = "▶";
  } else {
    ismCollapseBody.classList.remove("hidden");
    ismCollapseIcon.textContent = "▼";
  }
}

// ---------- Toggle ----------
function setMode(newMode) {
  mode = newMode;
  if (mode === "company") {
    toggleCompanyBtn.classList.add("active");
    toggleVesselBtn.classList.remove("active");
    companyFields.classList.remove("hidden");
    vesselFields.classList.add("hidden");
    if (vesselImo.value && !companyImo.value) companyImo.value = vesselImo.value;
  } else {
    toggleCompanyBtn.classList.remove("active");
    toggleVesselBtn.classList.add("active");
    companyFields.classList.add("hidden");
    vesselFields.classList.remove("hidden");
    if (companyImo.value && !vesselImo.value) vesselImo.value = companyImo.value;
  }
}

// ---------- Settings storage ----------
function getSetting(key, fallback) {
  if (!Office.context.roamingSettings) return fallback;
  const v = Office.context.roamingSettings.get(key);
  return v == null ? fallback : v;
}

function setSetting(key, value) {
  if (!Office.context.roamingSettings) return Promise.resolve();
  Office.context.roamingSettings.set(key, value);
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

function getConfig() {
  return {
    apiBase: (getSetting(SETTING_API_BASE, DEFAULT_API_BASE) || DEFAULT_API_BASE).replace(/\/$/, ""),
    apiKey:  getSetting(SETTING_API_KEY, "") || ""
  };
}

// ---------- Settings UI ----------
function openSettings() {
  apiBaseInput.value         = getSetting(SETTING_API_BASE, DEFAULT_API_BASE);
  apiKeyInput.value          = getSetting(SETTING_API_KEY, "");
  companiesApiKeyInput.value = getSetting(SETTING_COMPANIES_KEY, "");
  zammadUrlInput.value       = getSetting(SETTING_ZAMMAD_URL, "");
  zammadTokenInput.value     = getSetting(SETTING_ZAMMAD_TOKEN, "");
  settingsStatus.textContent = "";
  settingsStatus.className   = "";
  showView(settingsView);
}

async function saveSettings() {
  const base         = apiBaseInput.value.trim();
  const key          = apiKeyInput.value.trim();
  const companiesKey = companiesApiKeyInput.value.trim();
  const zUrl         = zammadUrlInput.value.trim();
  const zToken       = zammadTokenInput.value.trim();

  if (!base) {
    settingsStatus.textContent = "Please enter the API Base URL.";
    settingsStatus.className   = "status-msg error";
    return;
  }
  if (!key) {
    settingsStatus.textContent = "Please enter a Q Check API key.";
    settingsStatus.className   = "status-msg error";
    return;
  }
  if (!companiesKey) {
    settingsStatus.textContent = "Please enter a Companies Search key.";
    settingsStatus.className   = "status-msg error";
    return;
  }

  try {
    await setSetting(SETTING_API_BASE, base.replace(/\/$/, ""));
    await setSetting(SETTING_API_KEY, key);
    await setSetting(SETTING_COMPANIES_KEY, companiesKey);
    await setSetting(SETTING_ZAMMAD_URL, zUrl.replace(/\/$/, ""));
    await setSetting(SETTING_ZAMMAD_TOKEN, zToken);
    settingsStatus.textContent = "Saved.";
    settingsStatus.className   = "status-msg ok";

    if (firstRun) {
      firstRun = false;
      toggleWrap.classList.remove("hidden");
      settingsBtn.classList.remove("hidden");
      sendBtn.classList.remove("hidden");
      closeSettingsBtn.classList.remove("hidden");
      setTimeout(() => showView(formView), 800);
    }
  } catch (err) {
    settingsStatus.textContent = "Save failed: " + (err.message || err);
    settingsStatus.className   = "status-msg error";
  }
}

async function testConnection() {
  const base         = apiBaseInput.value.trim().replace(/\/$/, "") || DEFAULT_API_BASE;
  const key          = apiKeyInput.value.trim();
  const companiesKey = companiesApiKeyInput.value.trim();

  settingsStatus.textContent = "Testing…";
  settingsStatus.className   = "status-msg";

  const lines = [];

  // Test Q Check API
  if (!key) {
    lines.push("Q Check API: no key entered");
  } else {
    try {
      const resp = await fetch(`${base}/api/v1/qcheck/company`, {
        method:  "POST",
        headers: { "Content-Type": "application/json", "X-API-Key": key },
        body:    JSON.stringify({ company_imo: "" })
      });
      if (resp.status === 401)      lines.push("Q Check API: key rejected (401)");
      else if (resp.status === 503) lines.push("Q Check API: server not configured (503)");
      else                          lines.push(`Q Check API: OK (${resp.status}) ✓`);
    } catch {
      lines.push("Q Check API: connection failed");
    }
  }

  // Test Companies Search API
  if (!companiesKey) {
    lines.push("Companies API: no key entered");
  } else {
    try {
      const resp = await fetch(`${base}/api/companies/search?imo=0000001`, {
        headers: { "X-API-Key": companiesKey }
      });
      if (resp.status === 401) lines.push("Companies API: key rejected (401)");
      else                     lines.push(`Companies API: OK (${resp.status}) ✓`);
    } catch {
      lines.push("Companies API: connection failed");
    }
  }

  const hasError = lines.some(l => !l.includes("✓") && !l.includes("no key"));
  settingsStatus.innerHTML = lines.join("<br>");
  settingsStatus.className = "status-msg " + (hasError ? "error" : "ok");
}

// ---------- View switching ----------
function showView(view) {
  [formView, loadingView, errorView, confirmView, companyResult, vesselResult, settingsView, ticketView, ticketSuccessView]
    .forEach(v => v.classList.add("hidden"));
  view.classList.remove("hidden");
}

// ---------- Validation ----------
function isValidImo(v) {
  return /^\d{1,7}$/.test((v || "").trim());
}

function isSameImo(a, b) {
  return parseInt(a || 0, 10) === parseInt(b || 0, 10) && parseInt(a || 0, 10) !== 0;
}

// ---------- Run ----------
async function runQCheck() {
  const { apiBase, apiKey } = getConfig();
  if (!apiKey) {
    return showError("API key not set. Open Settings (⚙ top right) and paste your API key.");
  }

  if (mode === "company") {
    const imo  = companyImo.value.trim();
    const name = companyName.value.trim();
    if (!isValidImo(imo)) return showError("Please enter a valid Company IMO.");

    const proceed = () => callApi({
      url:  `${apiBase}/api/v1/qcheck/company`,
      apiKey,
      body: { company_imo: imo, company_name: name },
      onOk: (data) => renderCompanyResult({ imo, name, data })
    });

    if (getSetting(SETTING_COMPANIES_KEY, "")) {
      const results = await searchCompanies({ imo });
      if (results.find(r => isSameImo(r.company_imo, imo))) {
        pendingQCheck = proceed;
        showView(confirmView);
        return;
      }
    }

    await proceed();

  } else {
    const vImo = vesselImo.value.trim();
    const cImo = vesselCompanyImo.value.trim();
    const vNm  = vesselName.value.trim();
    const cNm  = vesselCompanyName.value.trim();
    if (!isValidImo(vImo)) return showError("Please enter a valid Vessel IMO.");
    if (cImo && !isValidImo(cImo)) return showError("Please enter a valid Company IMO.");

    const body = { vessel_imo: vImo, vessel_name: vNm };
    if (cImo) body.company_imo  = cImo;
    if (cNm)  body.company_name = cNm;

    await callApi({
      url:  `${apiBase}/api/v1/qcheck/vessel`,
      apiKey,
      body,
      onOk: (data) => renderVesselResult({ vImo, vNm, data })
    });
  }
}

async function callApi({ url, apiKey, body, onOk }) {
  loadingText.textContent = "Running Q Check…";
  loadingSub.textContent  = "This can take up to 90 seconds";
  showView(loadingView);
  try {
    const controller = new AbortController();
    const timeoutId  = setTimeout(() => controller.abort(), 120000);
    const resp = await fetch(url, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "X-API-Key": apiKey },
      body:    JSON.stringify(body),
      signal:  controller.signal
    });
    clearTimeout(timeoutId);
    if (!resp.ok) {
      let errText = `HTTP ${resp.status}`;
      try { const j = await resp.json(); if (j.error) errText = j.error; }
      catch (_) { try { errText = await resp.text(); } catch (_) {} }
      return showError(errText);
    }
    const data = await resp.json();
    onOk(data);
  } catch (err) {
    if (err.name === "AbortError") return showError("Request timed out after 120 seconds.");
    return showError(err.message || "Network error (your API may need CORS enabled for this origin).");
  }
}

function showError(message) {
  errorMsg.textContent = message;
  showView(errorView);
}

// ---------- Colour mapping ----------
function colorClass(v) {
  if (!v) return "amber";
  const s = v.toLowerCase();
  if (s.includes("not acceptable")) return "red";
  if (s.includes("very low"))       return "red";
  if (s.includes("review"))         return "amber";
  if (s.includes("low"))            return "amber";
  if (s.includes("high"))           return "green";
  if (s.includes("acceptable"))     return "green";
  return "amber";
}

// ---------- Render ----------
function renderCompanyResult({ imo, name, data }) {
  $("companyResultName").textContent = name || "Company";
  $("companyResultImo").textContent  = `IMO: ${imo}`;
  const banner = $("companyResultBanner");
  banner.textContent = data.global_performance || "Unknown";
  banner.className   = "result-banner " + colorClass(data.global_performance);
  const url = data.shareable_url || "";
  $("companyShareUrl").textContent = url;
  $("companyCopyBtn").onclick = () => copyToClipboard(url, $("companyCopyBtn"));
  $("companyOpenBtn").onclick = () => openUrl(url);
  $("companyNewBtn").onclick  = () => showView(formView);
  showView(companyResult);
}

function renderVesselResult({ vImo, vNm, data }) {
  $("vesselResultName").textContent = vNm || data.vessel_name || "Vessel";
  $("vesselResultImo").textContent  = `IMO: ${vImo}`;
  const a = data.assessment || {};
  const gBanner = $("vesselGlobalBanner");
  gBanner.textContent = a.global || "Unknown";
  gBanner.className   = "global-banner " + colorClass(a.global);
  setPill($("vesselAgePill"),     a.age);
  setPill($("vesselPscPill"),     a.psc);
  setPill($("vesselCompanyPill"), a.company);
  const url = data.shareable_url || "";
  $("vesselShareUrl").textContent = url;
  $("vesselCopyBtn").onclick = () => copyToClipboard(url, $("vesselCopyBtn"));
  $("vesselOpenBtn").onclick = () => openUrl(url);
  $("vesselNewBtn").onclick  = () => showView(formView);
  showView(vesselResult);
}

function setPill(el, value) {
  el.textContent = value || "Unknown";
  el.className   = "factor-pill " + colorClass(value);
}

// ---------- Company autocomplete ----------

function debounce(fn, ms) {
  let t;
  return function (...args) {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), ms);
  };
}

function escHtml(s) {
  return String(s || "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

async function searchCompanies({ name, imo }) {
  const { apiBase } = getConfig();
  const key = getSetting(SETTING_COMPANIES_KEY, "");
  if (!key) return [];

  const params = new URLSearchParams();
  if (name) params.set("name", name);
  if (imo)  params.set("imo", imo);

  try {
    const resp = await fetch(`${apiBase}/api/companies/search?${params}`, {
      headers: { "X-API-Key": key }
    });
    if (!resp.ok) return [];
    const data = await resp.json();
    return (data.success && Array.isArray(data.results)) ? data.results : [];
  } catch {
    return [];
  }
}

function setupAutocomplete({ inputEl, pairedEl, searchParam, dropdownEl }) {
  const selfField  = searchParam === "imo" ? "company_imo"  : "company_name";
  const otherField = searchParam === "imo" ? "company_name" : "company_imo";

  let activeIdx = -1;
  let hits      = [];

  const doSearch = debounce(async (q) => {
    if (!q || q.length < 2) { hide(); return; }
    showLoading();
    const opts = {};
    opts[searchParam] = q;
    hits = await searchCompanies(opts);
    render(q, hits);
  }, 300);

  inputEl.addEventListener("input", () => {
    activeIdx = -1;
    doSearch(inputEl.value.trim());
  });

  inputEl.addEventListener("keydown", (e) => {
    if (dropdownEl.classList.contains("hidden")) return;
    const items = dropdownEl.querySelectorAll(".ts-option[data-idx]");
    if (e.key === "ArrowDown") {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, items.length - 1);
      updateActive(items);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      updateActive(items);
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (activeIdx >= 0 && hits[activeIdx]) pick(hits[activeIdx]);
      else hide();
    } else if (e.key === "Escape") {
      hide();
    }
  });

  inputEl.addEventListener("blur", () => setTimeout(hide, 160));

  function updateActive(items) {
    items.forEach((el, i) => el.classList.toggle("active", i === activeIdx));
    if (activeIdx >= 0 && items[activeIdx]) items[activeIdx].scrollIntoView({ block: "nearest" });
  }

  function pick(item) {
    inputEl.value  = item[selfField]  || "";
    pairedEl.value = item[otherField] || "";
    hide();
  }

  function position() {
    const r = inputEl.getBoundingClientRect();
    dropdownEl.style.left  = r.left  + "px";
    dropdownEl.style.top   = (r.bottom + 2) + "px";
    dropdownEl.style.width = r.width + "px";
  }

  function showLoading() {
    dropdownEl.innerHTML = '<div class="ts-option-loading">Searching…</div>';
    position();
    dropdownEl.classList.remove("hidden");
  }

  function render(q, items) {
    dropdownEl.innerHTML = "";
    activeIdx = -1;

    items.forEach((item, i) => {
      const el = document.createElement("div");
      el.className   = "ts-option";
      el.dataset.idx = i;
      el.innerHTML   = `<div class="ts-opt-primary">${escHtml(item.company_imo)}</div>`
                     + `<div class="ts-opt-secondary">${escHtml(item.company_name)}</div>`;
      el.addEventListener("mousedown", (e) => { e.preventDefault(); pick(item); });
      dropdownEl.appendChild(el);
    });

    const hasExact = items.some(r => (r[selfField] || "").toLowerCase() === q.toLowerCase());
    if (!hasExact) {
      const footer = document.createElement("div");
      footer.className = "ts-option-create";
      footer.innerHTML = `No exact match — press <kbd>Enter</kbd> to use "${escHtml(q)}" as-is`;
      dropdownEl.appendChild(footer);
    }

    position();
    dropdownEl.classList.remove("hidden");
  }

  function hide() {
    dropdownEl.classList.add("hidden");
    dropdownEl.innerHTML = "";
    hits      = [];
    activeIdx = -1;
  }
}

// ---------- Helpers ----------
function copyToClipboard(text, btn) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => flashBtn(btn, "Copied!")).catch(() => fallbackCopy(text, btn));
  } else {
    fallbackCopy(text, btn);
  }
}

function fallbackCopy(text, btn) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.opacity  = "0";
  document.body.appendChild(ta);
  ta.select();
  try { document.execCommand("copy"); flashBtn(btn, "Copied!"); }
  catch (_) { flashBtn(btn, "Copy failed"); }
  document.body.removeChild(ta);
}

function flashBtn(btn, msg) {
  const original = btn.textContent;
  btn.textContent = msg;
  setTimeout(() => { btn.textContent = original; }, 1500);
}

function openUrl(url) {
  try {
    window.open(url, "_blank");
  } catch (_) {
    Office.context.ui.openBrowserWindow(url);
  }
}

// ---------- Send to Maritime Team ----------

function openTicketForm() {
  const zUrl   = getSetting(SETTING_ZAMMAD_URL, "");
  const zToken = getSetting(SETTING_ZAMMAD_TOKEN, "");
  if (!zUrl || !zToken) {
    openSettings();
    settingsStatus.textContent = "Please configure Zammad URL and Token to use Send to Maritime Team.";
    settingsStatus.className   = "status-msg error";
    return;
  }

  $("ticketTitle").value      = "";
  $("ticketText").value       = "";
  $("ticketVesselName").value = mode === "vessel" ? vesselName.value.trim() : "";
  $("ticketVesselImo").value  = mode === "vessel" ? vesselImo.value.trim()  : "";
  $("ticketErrorMsg").textContent = "";
  $("ticketErrorMsg").classList.add("hidden");
  initCategorySelects();
  showView(ticketView);
}

function initCategorySelects() {
  const l1 = $("ticketCatL1");
  l1.innerHTML = '<option value="">— Select category —</option>';
  Object.keys(CATEGORY_TREE).forEach(key => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = key;
    l1.appendChild(opt);
  });
  const l2wrap = $("catL2Wrap");
  const l3wrap = $("catL3Wrap");
  $("ticketCatL2").innerHTML = '<option value="">— Select sub-category —</option>';
  $("ticketCatL3").innerHTML = '<option value="">— Select type —</option>';
  l2wrap.classList.add("hidden");
  l3wrap.classList.add("hidden");
}

function updateCatL2() {
  const l1Val  = $("ticketCatL1").value;
  const l2     = $("ticketCatL2");
  const l2wrap = $("catL2Wrap");
  const l3wrap = $("catL3Wrap");

  l2.innerHTML = '<option value="">— Select sub-category —</option>';
  $("ticketCatL3").innerHTML = '<option value="">— Select type —</option>';
  l3wrap.classList.add("hidden");

  if (!l1Val || !CATEGORY_TREE[l1Val]) {
    l2wrap.classList.add("hidden");
    return;
  }

  Object.keys(CATEGORY_TREE[l1Val]).forEach(key => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = key;
    l2.appendChild(opt);
  });
  l2wrap.classList.remove("hidden");
}

function updateCatL3() {
  const l1Val  = $("ticketCatL1").value;
  const l2Val  = $("ticketCatL2").value;
  const l3     = $("ticketCatL3");
  const l3wrap = $("catL3Wrap");

  l3.innerHTML = '<option value="">— Select type (optional) —</option>';
  l3wrap.classList.add("hidden");

  if (!l1Val || !l2Val) return;
  const children = (CATEGORY_TREE[l1Val] || {})[l2Val];
  if (!children || children.length === 0) return;

  children.forEach(item => {
    const opt = document.createElement("option");
    opt.value = item;
    opt.textContent = item;
    l3.appendChild(opt);
  });
  l3wrap.classList.remove("hidden");
}

function showTicketError(msg) {
  const el = $("ticketErrorMsg");
  el.textContent = msg;
  el.classList.remove("hidden");
}

async function submitTicket() {
  const zUrl   = getSetting(SETTING_ZAMMAD_URL, "").replace(/\/$/, "");
  const zToken = getSetting(SETTING_ZAMMAD_TOKEN, "");

  const title  = $("ticketTitle").value.trim();
  const text   = $("ticketText").value.trim();
  const vName  = $("ticketVesselName").value.trim();
  const vImo   = $("ticketVesselImo").value.trim();
  const catL1  = $("ticketCatL1").value;
  const catL2  = $("ticketCatL2").value;
  const catL3  = $("ticketCatL3").value;

  $("ticketErrorMsg").classList.add("hidden");

  if (!title)  return showTicketError("Please enter a title.");
  if (!text)   return showTicketError("Please enter a description.");
  if (vImo && !isValidImo(vImo)) return showTicketError("Vessel IMO must be up to 7 digits.");
  if (!catL1)  return showTicketError("Please select a category.");

  const category = catL3 ? [catL1, catL2, catL3] : (catL2 ? [catL1, catL2] : [catL1]);

  const customer = Office.context.mailbox.userProfile.emailAddress;

  const body = {
    title,
    state:    "new",
    group:    "Maritime Operations Department",
    customer,
    article:  { subject: title, body: text, type: "note", internal: false },
    category
  };
  if (vName) body.vessel_name = vName;
  if (vImo)  body.vessel_imo  = parseInt(vImo, 10);

  loadingText.textContent = "Submitting ticket…";
  loadingSub.textContent  = "Please wait";
  showView(loadingView);

  try {
    const resp = await fetch(`${zUrl}/api/v1/tickets`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Token token=${zToken}` },
      body:    JSON.stringify(body)
    });

    if (resp.status === 201) {
      const data = await resp.json();
      renderTicketSuccess(data);
      return;
    }

    if (resp.status === 401) {
      showView(ticketView);
      showTicketError("Authentication failed (401). Please check your Zammad Token in Settings.");
      return;
    }

    let errMsg = `HTTP ${resp.status}`;
    try {
      const j = await resp.json();
      errMsg = j.error || (typeof j === "string" ? j : errMsg);
    } catch (_) {
      try { errMsg = await resp.text() || errMsg; } catch (_) {}
    }
    showView(ticketView);
    showTicketError(errMsg);
  } catch (err) {
    showView(ticketView);
    showTicketError(err.message || "Network error. Please check your connection.");
  }
}

function renderTicketSuccess(data) {
  $("ticketSuccessNumber").textContent = data.number || data.id || "–";
  showView(ticketSuccessView);
}
