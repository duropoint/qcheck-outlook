// Q Check — Outlook task pane logic

const DEFAULT_API_BASE = "https://pscplatformalpha.onrender.com";

// Storage keys (using Office's roaming settings = persists per user across devices)
const SETTING_API_BASE = "qcheck_api_base";
const SETTING_API_KEY  = "qcheck_api_key";

// ---------- DOM ----------
const $ = (id) => document.getElementById(id);

const toggleCompanyBtn  = $("toggleCompany");
const toggleVesselBtn   = $("toggleVessel");
const companyFields     = $("companyFields");
const vesselFields      = $("vesselFields");
const formView          = $("formView");
const loadingView       = $("loadingView");
const errorView         = $("errorView");
const errorMsg          = $("errorMsg");
const companyResult     = $("companyResult");
const vesselResult      = $("vesselResult");
const settingsView      = $("settingsView");

const companyImo        = $("companyImo");
const companyName       = $("companyName");
const vesselImo         = $("vesselImo");
const vesselCompanyImo  = $("vesselCompanyImo");
const vesselName        = $("vesselName");
const vesselCompanyName = $("vesselCompanyName");

const runBtn            = $("runBtn");
const backToFormBtn     = $("backToForm");
const useSelectionBtn   = $("useSelectionBtn");
const settingsBtn       = $("settingsBtn");
const closeSettingsBtn  = $("closeSettings");

const apiBaseInput      = $("apiBase");
const apiKeyInput       = $("apiKey");
const saveBtn           = $("saveBtn");
const testBtn           = $("testBtn");
const settingsStatus    = $("settingsStatus");

let mode = "company";

// ---------- Office.js init ----------
Office.onReady(() => {
  console.log("Q Check task pane ready");
  bindEvents();
});

// ---------- Event binding ----------
function bindEvents() {
  toggleCompanyBtn.addEventListener("click", () => setMode("company"));
  toggleVesselBtn.addEventListener("click", () => setMode("vessel"));
  useSelectionBtn.addEventListener("click", grabSelectedText);
  runBtn.addEventListener("click", runQCheck);
  backToFormBtn.addEventListener("click", () => showView(formView));
  settingsBtn.addEventListener("click", openSettings);
  closeSettingsBtn.addEventListener("click", () => showView(formView));
  saveBtn.addEventListener("click", saveSettings);
  testBtn.addEventListener("click", testConnection);
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

// ---------- Selected text ----------
function grabSelectedText() {
  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    return showError("Not running inside an Outlook email.");
  }
  const item = Office.context.mailbox.item;

  // getSelectedDataAsync works in both Read and Compose modes
  if (typeof item.getSelectedDataAsync === "function") {
    item.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const text = (result.value && result.value.data) || "";
        applySelectionToImo(text);
      } else {
        showError("Couldn't read selected text: " + (result.error?.message || "unknown error"));
      }
    });
  } else {
    showError("This Outlook version doesn't expose selected text.");
  }
}

function applySelectionToImo(text) {
  const match = (text || "").match(/\d{4,7}/);
  if (!match) {
    flashSelectionButton("No number found in selection");
    return;
  }
  const imo = match[0];
  if (mode === "company") {
    companyImo.value = imo;
  } else {
    vesselImo.value = imo;
  }
  flashSelectionButton("✓ Inserted " + imo);
}

function flashSelectionButton(msg) {
  const original = useSelectionBtn.textContent;
  useSelectionBtn.textContent = msg;
  setTimeout(() => { useSelectionBtn.textContent = original; }, 1800);
}

// ---------- Settings storage (uses Office roaming settings) ----------
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
  apiBaseInput.value = getSetting(SETTING_API_BASE, DEFAULT_API_BASE);
  apiKeyInput.value  = getSetting(SETTING_API_KEY, "");
  settingsStatus.textContent = "";
  settingsStatus.className = "";
  showView(settingsView);
}

async function saveSettings() {
  const base = (apiBaseInput.value.trim() || DEFAULT_API_BASE).replace(/\/$/, "");
  const key  = apiKeyInput.value.trim();

  if (!key) {
    settingsStatus.textContent = "Please enter an API key.";
    settingsStatus.className   = "status-msg error";
    return;
  }

  try {
    await setSetting(SETTING_API_BASE, base);
    await setSetting(SETTING_API_KEY, key);
    settingsStatus.textContent = "Saved.";
    settingsStatus.className   = "status-msg ok";
  } catch (err) {
    settingsStatus.textContent = "Save failed: " + (err.message || err);
    settingsStatus.className   = "status-msg error";
  }
}

async function testConnection() {
  const base = (apiBaseInput.value.trim() || DEFAULT_API_BASE).replace(/\/$/, "");
  const key  = apiKeyInput.value.trim();
  if (!key) {
    settingsStatus.textContent = "Please enter an API key first.";
    settingsStatus.className   = "status-msg error";
    return;
  }
  settingsStatus.textContent = "Testing…";
  settingsStatus.className   = "status-msg";

  try {
    const resp = await fetch(`${base}/api/v1/qcheck/company`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-API-Key": key },
      body: JSON.stringify({ company_imo: "" })
    });
    if (resp.status === 401) {
      settingsStatus.textContent = "API key rejected (401). Check the key.";
      settingsStatus.className   = "status-msg error";
    } else if (resp.status === 503) {
      settingsStatus.textContent = "Server not configured (503). Ask admin to set Q_CHECK_API_KEY on Render.";
      settingsStatus.className   = "status-msg error";
    } else {
      settingsStatus.textContent = `Connection OK (server responded ${resp.status}).`;
      settingsStatus.className   = "status-msg ok";
    }
  } catch (err) {
    settingsStatus.textContent = "Connection failed: " + (err.message || "network error");
    settingsStatus.className   = "status-msg error";
  }
}

// ---------- View switching ----------
function showView(view) {
  [formView, loadingView, errorView, companyResult, vesselResult, settingsView].forEach(v => v.classList.add("hidden"));
  view.classList.remove("hidden");
}

// ---------- Validation ----------
function isValidImo(v) {
  return /^\d{1,7}$/.test((v || "").trim());
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

    await callApi({
      url:    `${apiBase}/api/v1/qcheck/company`,
      apiKey,
      body:   { company_imo: imo, company_name: name },
      onOk:   (data) => renderCompanyResult({ imo, name, data })
    });
  } else {
    const vImo = vesselImo.value.trim();
    const cImo = vesselCompanyImo.value.trim();
    const vNm  = vesselName.value.trim();
    const cNm  = vesselCompanyName.value.trim();
    if (!isValidImo(vImo)) return showError("Please enter a valid Vessel IMO.");
    if (!isValidImo(cImo)) return showError("Please enter a valid Company IMO.");

    await callApi({
      url:    `${apiBase}/api/v1/qcheck/vessel`,
      apiKey,
      body:   {
        vessel_imo:   vImo,
        company_imo:  cImo,
        vessel_name:  vNm,
        company_name: cNm
      },
      onOk:   (data) => renderVesselResult({ vImo, vNm, data })
    });
  }
}

async function callApi({ url, apiKey, body, onOk }) {
  showView(loadingView);
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 120000);
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-API-Key": apiKey },
      body: JSON.stringify(body),
      signal: controller.signal
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
    if (err.name === "AbortError") {
      return showError("Request timed out after 120 seconds.");
    }
    return showError(err.message || "Network error (your API may need CORS enabled for this origin).");
  }
}

function showError(message) {
  errorMsg.textContent = message;
  showView(errorView);
}

// ---------- Render ----------
function renderCompanyResult({ imo, name, data }) {
  $("companyResultName").textContent = name || "Company";
  $("companyResultImo").textContent  = `IMO: ${imo}`;
  const banner = $("companyResultBanner");
  banner.textContent = data.global_performance || "Unknown";
  banner.className   = "result-banner " + colorClassForPerformance(data.global_performance);
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
  gBanner.className   = "global-banner " + colorClassForAssessment(a.global);
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
  el.className   = "factor-pill " + colorClassForAssessment(value);
}

function colorClassForPerformance(p) {
  if (!p) return "amber";
  const s = p.toLowerCase();
  if (s.includes("very low"))   return "red";
  if (s.includes("low"))        return "amber";
  if (s.includes("acceptable")) return "green";
  return "amber";
}
function colorClassForAssessment(a) {
  if (!a) return "amber";
  const s = a.toLowerCase();
  if (s.includes("not acceptable")) return "red";
  if (s.includes("review"))         return "amber";
  if (s.includes("acceptable"))     return "green";
  return "amber";
}

// ---------- Helpers ----------
function copyToClipboard(text, btn) {
  // Office task panes sometimes have restricted clipboard access; try modern API then fall back.
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
  ta.style.opacity = "0";
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand("copy");
    flashBtn(btn, "Copied!");
  } catch (_) {
    flashBtn(btn, "Copy failed");
  }
  document.body.removeChild(ta);
}

function flashBtn(btn, msg) {
  const original = btn.textContent;
  btn.textContent = msg;
  setTimeout(() => { btn.textContent = original; }, 1500);
}

function openUrl(url) {
  // Office task panes prefer Office.context.ui.displayDialogAsync, but window.open works for opening browsers in desktop.
  try {
    window.open(url, "_blank");
  } catch (_) {
    Office.context.ui.openBrowserWindow(url);
  }
}
