// Q Check — Outlook task pane logic

const DEFAULT_API_BASE = "https://pscplatformalpha.onrender.com";

const SETTING_API_BASE        = "qcheck_api_base";
const SETTING_API_KEY         = "qcheck_api_key";
const SETTING_COMPANIES_KEY   = "qcheck_companies_key";
const SETTING_ZAMMAD_PROXY    = "qcheck_zammad_proxy";
const SETTING_ZAMMAD_TOKEN    = "qcheck_zammad_token";

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
const closeSettingsBtn  = $("closeSettings");

const apiBaseInput          = $("apiBase");
const apiKeyInput           = $("apiKey");
const companiesApiKeyInput  = $("companiesApiKey");
const zammadProxyInput      = $("zammadProxyInput");
const zammadTokenInput      = $("zammadTokenInput");
const saveBtn               = $("saveBtn");
const testBtn               = $("testBtn");
const settingsStatus        = $("settingsStatus");

const companyEscalateBtn = $("companyEscalateBtn");
const vesselEscalateBtn  = $("vesselEscalateBtn");

let mode           = "company";
let firstRun       = false;
let pendingQCheck  = null;
let lastResultData = null;

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
  const proxyUrl     = getSetting(SETTING_ZAMMAD_PROXY, "");
  const zammadToken  = getSetting(SETTING_ZAMMAD_TOKEN, "");

  if (!base || !key || !companiesKey || !proxyUrl || !zammadToken) {
    firstRun = true;
    toggleWrap.classList.add("hidden");
    settingsBtn.classList.add("hidden");
    closeSettingsBtn.classList.add("hidden");
    apiBaseInput.value         = getSetting(SETTING_API_BASE, DEFAULT_API_BASE);
    apiKeyInput.value          = getSetting(SETTING_API_KEY, "");
    companiesApiKeyInput.value = getSetting(SETTING_COMPANIES_KEY, "");
    zammadProxyInput.value     = getSetting(SETTING_ZAMMAD_PROXY, "");
    zammadTokenInput.value     = getSetting(SETTING_ZAMMAD_TOKEN, "");
    settingsStatus.textContent = "";
    settingsStatus.className   = "";
    showView(settingsView);
  } else {
    closeSettingsBtn.classList.remove("hidden");
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
  closeSettingsBtn.addEventListener("click", () => showView(formView));
  saveBtn.addEventListener("click", saveSettings);
  testBtn.addEventListener("click", testConnection);
  $("confirmProceedBtn").addEventListener("click", () => {
    if (pendingQCheck) { const fn = pendingQCheck; pendingQCheck = null; fn(); }
  });
  $("confirmCancelBtn").addEventListener("click", () => {
    pendingQCheck = null;
    showView(formView);
  });
  companyEscalateBtn.addEventListener("click", () =>
    submitZammadTicket(companyEscalateBtn, $("companyZammadStatus"))
  );
  vesselEscalateBtn.addEventListener("click", () =>
    submitZammadTicket(vesselEscalateBtn, $("vesselZammadStatus"))
  );
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
  zammadProxyInput.value     = getSetting(SETTING_ZAMMAD_PROXY, "");
  zammadTokenInput.value     = getSetting(SETTING_ZAMMAD_TOKEN, "");
  settingsStatus.textContent = "";
  settingsStatus.className   = "";
  showView(settingsView);
}

async function saveSettings() {
  const base         = apiBaseInput.value.trim();
  const key          = apiKeyInput.value.trim();
  const companiesKey = companiesApiKeyInput.value.trim();
  const proxyUrl     = zammadProxyInput.value.trim();
  const zammadToken  = zammadTokenInput.value.trim();

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
  if (!proxyUrl) {
    settingsStatus.textContent = "Please enter the Zammad Proxy URL.";
    settingsStatus.className   = "status-msg error";
    return;
  }
  if (!zammadToken) {
    settingsStatus.textContent = "Please enter a Zammad token.";
    settingsStatus.className   = "status-msg error";
    return;
  }

  try {
    await setSetting(SETTING_API_BASE, base.replace(/\/$/, ""));
    await setSetting(SETTING_API_KEY, key);
    await setSetting(SETTING_COMPANIES_KEY, companiesKey);
    await setSetting(SETTING_ZAMMAD_PROXY, proxyUrl.replace(/\/$/, ""));
    await setSetting(SETTING_ZAMMAD_TOKEN, zammadToken);
    settingsStatus.textContent = "Saved.";
    settingsStatus.className   = "status-msg ok";

    if (firstRun) {
      firstRun = false;
      toggleWrap.classList.remove("hidden");
      settingsBtn.classList.remove("hidden");
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
  const proxyUrl     = zammadProxyInput.value.trim().replace(/\/$/, "");
  const zammadToken  = zammadTokenInput.value.trim();

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

  // Test Zammad proxy (Cloudflare Worker)
  if (!proxyUrl || !zammadToken) {
    lines.push("Zammad proxy: proxy URL and Zammad token required");
  } else {
    try {
      const resp = await fetch(`${proxyUrl}/groups`, {
        headers: { "X-Zammad-Token": zammadToken }
      });
      if (resp.status === 401)      lines.push("Zammad proxy: Zammad token rejected (401)");
      else if (resp.status === 404) lines.push("Zammad proxy: endpoint not found — check Worker URL");
      else if (!resp.ok)            lines.push(`Zammad proxy: error (${resp.status})`);
      else                          lines.push(`Zammad proxy: OK (${resp.status}) ✓`);
    } catch {
      lines.push("Zammad proxy: connection failed");
    }
  }

  const hasError = lines.some(l => !l.includes("✓") && !l.includes("no key") && !l.includes("required"));
  settingsStatus.innerHTML = lines.join("<br>");
  settingsStatus.className = "status-msg " + (hasError ? "error" : "ok");
}

// ---------- View switching ----------
function showView(view) {
  [formView, loadingView, errorView, confirmView, companyResult, vesselResult, settingsView]
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

  lastResultData = { mode: "company", imo, name, data };

  const notAcceptable = colorClass(data.global_performance) !== "green";
  companyEscalateBtn.classList.toggle("hidden", !notAcceptable);
  companyEscalateBtn.disabled    = false;
  companyEscalateBtn.textContent = "Send to Maritime Team";
  const companyZammadStatus = $("companyZammadStatus");
  companyZammadStatus.classList.add("hidden");
  companyZammadStatus.textContent = "";

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

  lastResultData = {
    mode: "vessel",
    vImo,
    vNm:  vNm || data.vessel_name || "",
    cImo: vesselCompanyImo.value.trim(),
    cNm:  vesselCompanyName.value.trim(),
    data
  };

  const notAcceptable = colorClass(a.global) !== "green";
  vesselEscalateBtn.classList.toggle("hidden", !notAcceptable);
  vesselEscalateBtn.disabled    = false;
  vesselEscalateBtn.textContent = "Send to Maritime Team";
  const vesselZammadStatus = $("vesselZammadStatus");
  vesselZammadStatus.classList.add("hidden");
  vesselZammadStatus.textContent = "";

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

// ---------- Zammad ticket ----------

function buildZammadTitle(d) {
  if (d.mode === "company") {
    return `Q Check Review Required : Company${d.name ? " " + d.name : ""} ${d.imo}`;
  }
  return `Q Check Review Required : Vessel${d.vNm ? " " + d.vNm : ""} ${d.vImo}`;
}

function buildZammadDescription(d) {
  const profile   = Office.context.mailbox.userProfile;
  const userName  = profile.displayName  || "";
  const userEmail = profile.emailAddress || "";

  if (d.mode === "company") {
    const global = d.data.global_performance || "Unknown";
    const url    = d.data.shareable_url      || "";
    return [
      "Dear colleagues,",
      "",
      "Please note that the below Company requires a further Review on the Eligibility to MAR before we proceed further.",
      `Company IMO Number : ${d.imo}`,
      `Company Name : ${d.name || "Unknown"}`,
      "",
      `Global Assessment : ${global}`,
      `Report Link : ${url}`,
      "",
      "Thank you,",
      userName,
      userEmail
    ].join("\n");
  }

  const a      = d.data.assessment || {};
  const global = a.global  || "Unknown";
  const url    = d.data.shareable_url || "";
  const lines  = [
    "Dear colleagues,",
    "",
    "Please note that the below Vessel requires a further Review on the Eligibility to MAR before we proceed further.",
    `Vessel IMO Number : ${d.vImo}`,
    `Vessel Name : ${d.vNm || "Unknown"}`
  ];
  if (d.cImo) lines.push(`Company IMO Number : ${d.cImo}`);
  if (d.cNm)  lines.push(`Company Name : ${d.cNm}`);
  lines.push(
    "",
    `Global Assessment : ${global}`,
    `Age Criteria : ${a.age     || "Unknown"}`,
    `PSC Performance : ${a.psc  || "Unknown"}`,
    `Company Status : ${a.company || "Unknown"}`,
    `Report Link : ${url}`,
    "",
    "Thank you,",
    userName,
    userEmail
  );
  return lines.join("\n");
}

async function submitZammadTicket(btnEl, statusEl) {
  const proxyUrl    = getSetting(SETTING_ZAMMAD_PROXY, "").replace(/\/$/, "");
  const zammadToken = getSetting(SETTING_ZAMMAD_TOKEN, "");

  if (!proxyUrl || !zammadToken) {
    statusEl.textContent = "Zammad proxy URL or token not configured — open Settings.";
    statusEl.className   = "zammad-status error";
    statusEl.classList.remove("hidden");
    return;
  }

  const d         = lastResultData;
  const title     = buildZammadTitle(d);
  const body      = buildZammadDescription(d);
  const userEmail = Office.context.mailbox.userProfile.emailAddress;

  const ticket = {
    title,
    state:    "new",
    group:    "Maritime Operations Department",
    customer: userEmail,
    article:  { subject: title, body, type: "note", internal: false },
    category: ["Maritime Department", "Registrations / Deletions", "New Registrations"]
  };

  if (d.mode === "vessel") {
    ticket.vessel_name = d.vNm || "";
    ticket.vessel_imo  = parseInt(d.vImo, 10);
  }

  btnEl.disabled    = true;
  btnEl.textContent = "Sending…";
  statusEl.classList.add("hidden");

  try {
    const resp = await fetch(`${proxyUrl}/tickets`, {
      method:  "POST",
      headers: {
        "Content-Type":   "application/json",
        "X-Zammad-Token": zammadToken
      },
      body: JSON.stringify(ticket)
    });

    if (resp.status === 201) {
      const json = await resp.json();
      const num  = json.number || json.id;
      btnEl.textContent    = "Sent ✓";
      statusEl.textContent = `Ticket #${num} created successfully.`;
      statusEl.className   = "zammad-status ok";
      statusEl.classList.remove("hidden");
    } else if (resp.status === 401) {
      btnEl.disabled       = false;
      btnEl.textContent    = "Send to Maritime Team";
      statusEl.textContent = "Authentication failed (401). Check your Zammad token in Settings.";
      statusEl.className   = "zammad-status error";
      statusEl.classList.remove("hidden");
    } else {
      let errText = `Error ${resp.status}`;
      try { const j = await resp.json(); if (j.error) errText = j.error; } catch (_) {}
      btnEl.disabled       = false;
      btnEl.textContent    = "Send to Maritime Team";
      statusEl.textContent = errText;
      statusEl.className   = "zammad-status error";
      statusEl.classList.remove("hidden");
    }
  } catch (err) {
    btnEl.disabled       = false;
    btnEl.textContent    = "Send to Maritime Team";
    statusEl.textContent = err.message || "Network error. Please try again.";
    statusEl.className   = "zammad-status error";
    statusEl.classList.remove("hidden");
  }
}
