// Q Check — task pane logic (Outlook + Browser + Chrome Extension)

const DEFAULT_API_BASE  = "https://pscplatformalpha.onrender.com";
const ZAMMAD_PROXY_URL  = "https://zammad-dashboard.onrender.com/api/zammad";

const SETTING_API_BASE      = "qcheck_api_base";
const SETTING_API_KEY       = "qcheck_api_key";
const SETTING_COMPANIES_KEY = "qcheck_companies_key";
const SETTING_ZAMMAD_TOKEN   = "qcheck_zammad_token";
const SETTING_DASHBOARD_KEY  = "qcheck_dashboard_key";
const SETTING_USER_EMAIL     = "qcheck_user_email";
const SETTING_FAVORITES      = "qcheck_favorites";

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

const apiBaseInput         = $("apiBase");
const apiKeyInput          = $("apiKey");
const companiesApiKeyInput = $("companiesApiKey");
const zammadTokenInput      = $("zammadTokenInput");
const dashboardApiKeyInput  = $("dashboardApiKeyInput");
const userEmailInput        = $("userEmailInput");
const userEmailRow         = $("userEmailRow");
const saveBtn              = $("saveBtn");
const testBtn              = $("testBtn");
const settingsStatus       = $("settingsStatus");

const companyEscalateBtn = $("companyEscalateBtn");
const vesselEscalateBtn  = $("vesselEscalateBtn");

const backBtn           = $("backBtn");
const headerTitle       = $("headerTitle");
const mainView          = $("mainView");
const maritimeView      = $("maritimeView");
const salesView            = $("salesView");
const seafarersView        = $("seafarersView");
const seafarersBridgeView  = $("seafarersBridgeView");
const favoritesSection     = $("favoritesSection");
const favList              = $("favList");
const zammadSearchView     = $("zammadSearchView");
const zammadReportsView = $("zammadReportsView");
const zvlSearchInput    = $("zvlSearchInput");
const zvlStatus         = $("zvlStatus");
const zvlResultsList    = $("zvlResultsList");
const zvlDetail         = $("zvlDetail");
const zvlDetailCard     = $("zvlDetailCard");
const zvlBackToResults  = $("zvlBackToResults");
const companySearchView  = $("companySearchView");
const csSearchInput      = $("csSearchInput");
const csStatus           = $("csStatus");
const csResultsList      = $("csResultsList");
const csDetail           = $("csDetail");
const csDetailCard       = $("csDetailCard");
const csFleetStatus      = $("csFleetStatus");
const csFleetList        = $("csFleetList");
const csBackToSearch     = $("csBackToSearch");
const csVesselSection    = $("csVesselSection");
const csVesselDetailCard = $("csVesselDetailCard");
const csBackToCompany    = $("csBackToCompany");
const rptFrom           = $("rptFrom");
const rptTo             = $("rptTo");
const rptVesselImo      = $("rptVesselImo");
const rptIncludeClosed  = $("rptIncludeClosed");
const rptGenerateBtn    = $("rptGenerateBtn");
const rptStatus         = $("rptStatus");

let mode           = "company";
let firstRun       = false;
let pendingQCheck  = null;
let lastResultData = null;
let isComposeMode  = false;
let currentView        = null;
let settingsReturnView = null;
let zvlSearchTimer     = null;
let zvlResults         = [];

// Navigation-origin tracking — each sub-view remembers which view to return to
// (supports reaching tools from Favorites on mainView, or from a category view)
let formViewBack        = null; // back target for formView
let zammadSearchBack    = null; // back target for zammadSearchView
let zammadReportsBack   = null; // back target for zammadReportsView
let bridgeViewBack      = null; // back target for seafarersBridgeView
let companySearchBack   = null; // back target for companySearchView

let csSearchTimer    = null;
let csResults        = [];
let csCurrentCompany = null;

// Favorites — array of tool IDs, persisted via SETTING_FAVORITES
let favorites = [];

// ---------- Tool registry ----------
// Single source of truth for tool metadata + navigation actions.
// Adding a tool here automatically makes it available in Favorites.
const TOOL_DEFS = {
  "qcheck": {
    icon: "✓",        // ✓
    name: "Q Check",
    desc: "PSC quality check by vessel or company IMO",
    navigate(origin) {
      formViewBack = origin || maritimeView;
      showView(formView);
    }
  },
  "vessel-search": {
    icon: "🔍",  // 🔍
    name: "Vessel Search",
    desc: "Look up vessel data by name",
    navigate(origin) {
      zvlSearchInput.value = "";
      zvlStatus.textContent = "";
      zvlResultsList.innerHTML = "";
      zvlDetail.classList.add("hidden");
      zvlResultsList.classList.remove("hidden");
      zammadSearchBack = origin || maritimeView;
      showView(zammadSearchView);
      setTimeout(() => zvlSearchInput.focus(), 80);
    }
  },
  "zammad-reports": {
    icon: "📄",  // 📄
    name: "Zammad Reports",
    desc: "Generate maritime PDF reports",
    navigate(origin) {
      zammadReportsBack = origin || maritimeView;
      showView(zammadReportsView);
    }
  },
  "company-search": {
    icon: "🏢",  // 🏢
    name: "Company Search",
    desc: "Look up company data and fleet by name or IMO",
    navigate(origin) {
      csSearchInput.value = "";
      csStatus.textContent = "";
      csResultsList.innerHTML = "";
      csResultsList.classList.remove("hidden");
      csDetail.classList.add("hidden");
      csVesselSection.classList.add("hidden");
      companySearchBack = origin || maritimeView;
      showView(companySearchView);
      setTimeout(() => csSearchInput.focus(), 80);
    }
  },
  "zoho-bridge": {
    icon: "🔗",  // 🔗
    name: "Zoho BMAR Bridge",
    desc: "Transfer seafarer data to Zoho BMAR automatically",
    extOnly: true,   // hidden outside the Chrome Extension
    navigate(origin) {
      bridgeViewBack = origin || seafarersView;
      showSeafarersBridgeView();
    }
  }
};

// ---------- Environment helpers ----------

/** True when running inside the Chrome Extension side-panel (or the iframe it embeds). */
function isExtensionContext() {
  return Env.env === "extension"
    || new URLSearchParams(location.search).get("context") === "extension";
}

/**
 * Hide tool tiles that are marked extOnly in TOOL_DEFS when the UI is not
 * running inside the Chrome Extension.
 * The Seafarers *category* tile is always visible; only tools inside it that
 * are extension-only get hidden, and an empty-state message is shown instead.
 */
function applyExtOnlyVisibility() {
  if (isExtensionContext()) return; // nothing to hide

  // Hide extension-only tools inside the Seafarers sub-menu
  const sfbrBtn = $("tileSfbrBtn");
  if (sfbrBtn) sfbrBtn.classList.add("hidden");

  // Show the empty-state message so the sub-menu isn't blank
  const emptyNote = $("seafarersEmpty");
  if (emptyNote) emptyNote.classList.remove("hidden");
}

// ---------- Env init ----------
Env.ready(() => {
  if (!Env.isOffice) {
    userEmailRow.classList.remove("hidden");
    const note = $("footerNote");
    if (note) note.textContent = "Can take 10–90 seconds. Don't close this tab.";
  }

  isComposeMode = !!(
    typeof Office !== "undefined" &&
    Office.context &&
    Office.context.mailbox &&
    Office.context.mailbox.item &&
    Office.context.mailbox.item.body &&
    typeof Office.context.mailbox.item.body.setSelectedDataAsync === "function"
  );

  loadFavorites();   // populate favorites[] before rendering
  bindEvents();
  initApp();
  applyExtOnlyVisibility();
  renderFavStars();
  renderFavoritesSection();
  setupAutocomplete({ inputEl: companyImo,       pairedEl: companyName,       searchParam: "imo",  dropdownEl: $("companyImoDropdown") });
  setupAutocomplete({ inputEl: companyName,       pairedEl: companyImo,        searchParam: "name", dropdownEl: $("companyNameDropdown") });
  setupAutocomplete({ inputEl: vesselCompanyImo,  pairedEl: vesselCompanyName, searchParam: "imo",  dropdownEl: $("vesselCompanyImoDropdown") });
  setupAutocomplete({ inputEl: vesselCompanyName, pairedEl: vesselCompanyImo,  searchParam: "name", dropdownEl: $("vesselCompanyNameDropdown") });
  document.querySelectorAll(".paste-btn").forEach(btn => {
    btn.addEventListener("mousedown", e => e.preventDefault());
    btn.addEventListener("click", () => handlePaste($(btn.dataset.target)));
  });
});

// ---------- App init ----------
// Gate: all five settings must be saved before the form is accessible.
function initApp() {
  const base         = Env.getSetting(SETTING_API_BASE, "");
  const key          = Env.getSetting(SETTING_API_KEY, "");
  const companiesKey = Env.getSetting(SETTING_COMPANIES_KEY, "");
  const zammadToken  = Env.getSetting(SETTING_ZAMMAD_TOKEN, "");
  const dashboardKey = Env.getSetting(SETTING_DASHBOARD_KEY, "");

  if (!base || !key || !companiesKey || !zammadToken || !dashboardKey) {
    firstRun = true;
    toggleWrap.classList.add("hidden");
    settingsBtn.classList.add("hidden");
    closeSettingsBtn.classList.add("hidden");
    apiBaseInput.value         = Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE);
    apiKeyInput.value          = Env.getSetting(SETTING_API_KEY, "");
    companiesApiKeyInput.value = Env.getSetting(SETTING_COMPANIES_KEY, "");
    zammadTokenInput.value     = Env.getSetting(SETTING_ZAMMAD_TOKEN, "");
    dashboardApiKeyInput.value = Env.getSetting(SETTING_DASHBOARD_KEY, "");
    if (!Env.isOffice) userEmailInput.value = Env.getSetting(SETTING_USER_EMAIL, "");
    settingsStatus.textContent = "";
    settingsStatus.className   = "";
    showView(settingsView);
  } else {
    closeSettingsBtn.classList.remove("hidden");
    showView(mainView);
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
  closeSettingsBtn.addEventListener("click", () => showView(settingsReturnView || mainView));
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
  $("tileMaritimeBtn").addEventListener("click", () => showView(maritimeView));
  $("tileSalesBtn").addEventListener("click", () => showView(salesView));
  $("tileSeafarersBtn").addEventListener("click", () => showView(seafarersView));

  $("tileQCheckBtn").addEventListener("click", () => {
    formViewBack = maritimeView;
    showView(formView);
  });
  $("tileVesselSearchBtn").addEventListener("click", () => {
    TOOL_DEFS["vessel-search"].navigate(maritimeView);
  });
  $("tileCompanySearchBtn").addEventListener("click", () => {
    TOOL_DEFS["company-search"].navigate(maritimeView);
  });
  $("tileZammadReportsBtn").addEventListener("click", () => {
    TOOL_DEFS["zammad-reports"].navigate(maritimeView);
  });

  // Sales sub-menu — Q Check re-routes to the same form (no duplication)
  $("tileSalesQCheckBtn").addEventListener("click", () => {
    formViewBack = salesView;
    showView(formView);
  });

  // Seafarers sub-menu
  $("tileSfbrBtn").addEventListener("click", () => {
    TOOL_DEFS["zoho-bridge"].navigate(seafarersView);
  });

  // Seafarers Bridge — inject button
  $("sfbrInjectBtn").addEventListener("click", sendSfbrInject);

  // Star (favorites) buttons — stop propagation so tile click isn't also fired
  document.querySelectorAll(".fav-star-btn").forEach(star => {
    star.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleFavorite(star.dataset.tool);
    });
  });
  zvlSearchInput.addEventListener("input", onZvlInput);
  zvlBackToResults.addEventListener("click", () => {
    zvlDetail.classList.add("hidden");
    zvlResultsList.classList.remove("hidden");
  });
  csSearchInput.addEventListener("input", onCsInput);
  csBackToSearch.addEventListener("click", () => {
    csDetail.classList.add("hidden");
    csResultsList.classList.remove("hidden");
  });
  csBackToCompany.addEventListener("click", () => {
    csVesselSection.classList.add("hidden");
    csDetail.classList.remove("hidden");
  });
  rptGenerateBtn.addEventListener("click", generateReport);
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
    // Clipboard access denied — let user paste manually
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

// ---------- Config helper ----------
function getConfig() {
  return {
    apiBase: (Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE) || DEFAULT_API_BASE).replace(/\/$/, ""),
    apiKey:  Env.getSetting(SETTING_API_KEY, "") || ""
  };
}

// ---------- Settings UI ----------
function openSettings() {
  settingsReturnView = currentView;
  apiBaseInput.value         = Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE);
  apiKeyInput.value          = Env.getSetting(SETTING_API_KEY, "");
  companiesApiKeyInput.value = Env.getSetting(SETTING_COMPANIES_KEY, "");
  zammadTokenInput.value     = Env.getSetting(SETTING_ZAMMAD_TOKEN, "");
  dashboardApiKeyInput.value = Env.getSetting(SETTING_DASHBOARD_KEY, "");
  if (!Env.isOffice) userEmailInput.value = Env.getSetting(SETTING_USER_EMAIL, "");
  settingsStatus.textContent = "";
  settingsStatus.className   = "";
  showView(settingsView);
}

async function saveSettings() {
  const base         = apiBaseInput.value.trim();
  const key          = apiKeyInput.value.trim();
  const companiesKey = companiesApiKeyInput.value.trim();
  const zammadToken  = zammadTokenInput.value.trim();
  const dashboardKey = dashboardApiKeyInput.value.trim();

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
  if (!zammadToken) {
    settingsStatus.textContent = "Please enter a Zammad token.";
    settingsStatus.className   = "status-msg error";
    return;
  }
  if (!dashboardKey) {
    settingsStatus.textContent = "Please enter a Dashboard API key.";
    settingsStatus.className   = "status-msg error";
    return;
  }

  try {
    await Env.setSetting(SETTING_API_BASE, base.replace(/\/$/, ""));
    await Env.setSetting(SETTING_API_KEY, key);
    await Env.setSetting(SETTING_COMPANIES_KEY, companiesKey);
    await Env.setSetting(SETTING_ZAMMAD_TOKEN, zammadToken);
    await Env.setSetting(SETTING_DASHBOARD_KEY, dashboardKey);
    if (!Env.isOffice) {
      await Env.setSetting(SETTING_USER_EMAIL, userEmailInput.value.trim());
    }
    settingsStatus.textContent = "Saved.";
    settingsStatus.className   = "status-msg ok";

    if (firstRun) {
      firstRun = false;
      closeSettingsBtn.classList.remove("hidden");
      setTimeout(() => showView(mainView), 800);
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
  const zammadToken  = zammadTokenInput.value.trim();
  const dashboardKey = dashboardApiKeyInput.value.trim();

  settingsStatus.textContent = "Testing…";
  settingsStatus.className   = "status-msg";

  const lines = [];

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

  if (!zammadToken) {
    lines.push("Zammad proxy: token required");
  } else {
    try {
      const resp = await fetch(`${ZAMMAD_PROXY_URL}/groups`, {
        headers: { "X-Zammad-Token": zammadToken }
      });
      if (resp.status === 401)      lines.push("Zammad proxy: token rejected (401)");
      else if (resp.status === 404) lines.push("Zammad proxy: endpoint not found");
      else if (!resp.ok)            lines.push(`Zammad proxy: error (${resp.status})`);
      else                          lines.push(`Zammad proxy: OK (${resp.status}) ✓`);
    } catch {
      lines.push("Zammad proxy: connection failed");
    }
  }

  if (!dashboardKey) {
    lines.push("Dashboard API: no key entered");
  } else {
    try {
      const resp = await fetch("https://zammad-dashboard.onrender.com/api/v1/report", {
        method:  "POST",
        headers: { "Authorization": `Bearer ${dashboardKey}`, "Content-Type": "application/json" },
        body:    JSON.stringify({ format: "pdf", limit: 1, filters: { state: ["open"] } })
      });
      if (resp.status === 401)      lines.push("Dashboard API: key rejected (401)");
      else if (resp.status === 404) lines.push("Dashboard API: endpoint not found");
      else                          lines.push(`Dashboard API: OK (${resp.status}) ✓`);
    } catch {
      lines.push("Dashboard API: connection failed");
    }
  }

  const hasError = lines.some(l => !l.includes("✓") && !l.includes("no key") && !l.includes("required"));
  settingsStatus.innerHTML = lines.join("<br>");
  settingsStatus.className = "status-msg " + (hasError ? "error" : "ok");
}

// ---------- View switching ----------
function showView(view) {
  [mainView, maritimeView, salesView, seafarersView, seafarersBridgeView,
   zammadSearchView, zammadReportsView, companySearchView,
   formView, loadingView, errorView, confirmView, companyResult, vesselResult, settingsView]
    .filter(v => v)
    .forEach(v => v.classList.add("hidden"));
  view.classList.remove("hidden");
  currentView = view;
  updateNavChrome(view);
}

function updateNavChrome(view) {
  const id = view.id;
  let title        = "Q Check";
  let showSettings = false;
  let showToggle   = false;
  let backTarget   = null;

  if (id === "mainView") {
    title = "Euromar Toolkit";
    showSettings = !firstRun;
  } else if (id === "maritimeView") {
    title = "Maritime";
    showSettings = !firstRun;
    backTarget = mainView;
  } else if (id === "salesView") {
    title = "Sales";
    showSettings = !firstRun;
    backTarget = mainView;
  } else if (id === "seafarersView") {
    title = "Seafarers";
    showSettings = !firstRun;
    backTarget = mainView;
  } else if (id === "seafarersBridgeView") {
    title = "Zoho BMAR Bridge";
    backTarget = bridgeViewBack || seafarersView;
  } else if (id === "formView") {
    showToggle = !firstRun;
    backTarget = formViewBack || maritimeView;
  } else if (id === "zammadSearchView") {
    title = "Vessel Search";
    backTarget = zammadSearchBack || maritimeView;
  } else if (id === "companySearchView") {
    title = "Company Search";
    backTarget = companySearchBack || maritimeView;
  } else if (id === "zammadReportsView") {
    title = "Zammad Reports";
    backTarget = zammadReportsBack || maritimeView;
  } else if (id === "settingsView") {
    title = "Settings";
    if (!firstRun) backTarget = "settingsReturn";
  }

  headerTitle.textContent = title;
  settingsBtn.classList.toggle("hidden", !showSettings);
  toggleWrap.classList.toggle("hidden", !showToggle);

  if (backTarget === "settingsReturn") {
    backBtn.classList.remove("hidden");
    backBtn.onclick = () => showView(settingsReturnView || mainView);
  } else if (backTarget) {
    backBtn.classList.remove("hidden");
    backBtn.onclick = () => showView(backTarget);
  } else {
    backBtn.classList.add("hidden");
    backBtn.onclick = null;
  }
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

    if (Env.getSetting(SETTING_COMPANIES_KEY, "")) {
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
    if (cImo) {
      body.company_imo = cImo;
      if (cNm) body.company_name = cNm;
    }

    await callApi({
      url:  `${apiBase}/api/v1/qcheck/vessel`,
      apiKey,
      body,
      onOk: (data) => renderVesselResult({ vImo, vNm, cImo, cNm, data })
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
  $("companyCopyBtn").onclick      = () => copyToClipboard(url, $("companyCopyBtn"));
  $("companyOpenBtn").onclick      = () => Env.openUrl(url);
  $("companyCopyTableBtn").onclick = () => copyAsTable(buildCompanyTable({ imo, name, data }), $("companyCopyTableBtn"));
  $("companyNewBtn").onclick       = () => showView(formView);
  const companyInsertBtn = $("companyInsertEmailBtn");
  companyInsertBtn.classList.toggle("hidden", !isComposeMode);
  companyInsertBtn.onclick = () => insertIntoEmail(buildCompanyTable({ imo, name, data }), companyInsertBtn);

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

function renderVesselResult({ vImo, vNm, cImo, cNm, data }) {
  $("vesselResultName").textContent = vNm || data.vessel_name || "Vessel";
  $("vesselResultImo").textContent  = `IMO: ${vImo}`;
  const a = data.assessment || {};
  const gBanner = $("vesselGlobalBanner");
  gBanner.textContent = a.global || "Unknown";
  gBanner.className   = "global-banner " + colorClass(a.global);
  setPill($("vesselAgePill"),     a.age);
  setPill($("vesselPscPill"),     a.psc);
  setPill($("vesselCompanyPill"), a.company);

  // ISM Manager — prefer data the user already entered in the form.
  // Fall back to async enrichment from the ships API if the form was blank.
  const ismCard = $("vesselIsmCard");
  if (cImo) {
    $("vesselIsmName").textContent = cNm || "";
    $("vesselIsmImo").textContent  = `IMO ${cImo}`;
    ismCard.classList.remove("hidden");
    // Mirror into data so buildVesselTable() picks it up via its normal path
    data.ism_manager     = cNm  || "";
    data.ism_manager_imo = cImo || "";
  } else {
    ismCard.classList.add("hidden");
  }

  const url = data.shareable_url || "";
  $("vesselShareUrl").textContent = url;
  $("vesselCopyBtn").onclick      = () => copyToClipboard(url, $("vesselCopyBtn"));
  $("vesselOpenBtn").onclick      = () => Env.openUrl(url);
  $("vesselCopyTableBtn").onclick = () => copyAsTable(buildVesselTable({ vImo, vNm: vNm || data.vessel_name || "", data }), $("vesselCopyTableBtn"));
  $("vesselNewBtn").onclick       = () => showView(formView);
  const vesselInsertBtn = $("vesselInsertEmailBtn");
  vesselInsertBtn.classList.toggle("hidden", !isComposeMode);
  vesselInsertBtn.onclick = () => insertIntoEmail(buildVesselTable({ vImo, vNm: vNm || data.vessel_name || "", data }), vesselInsertBtn);

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

  // If no ISM data was entered in the form, try enriching from the ships API
  if (!cImo) enrichVesselWithIsmData(vImo, vNm || data.vessel_name || "");
}

/**
 * After a vessel Q Check, silently fetch ISM manager data from the ship-search
 * API and populate the ISM card + lastResultData so Copy as Table includes it.
 */
async function enrichVesselWithIsmData(vImo, vNm) {
  const apiBase   = (Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE) || DEFAULT_API_BASE).replace(/\/$/, "");
  const searchKey = Env.getSetting(SETTING_COMPANIES_KEY, "") || "";
  if (!searchKey || !vNm) return;

  try {
    const resp = await fetch(`${apiBase}${SHIP_SEARCH_PATH}?name=${encodeURIComponent(vNm)}`, {
      headers: { "X-API-Key": searchKey }
    });
    if (!resp.ok) return;

    const json    = await resp.json();
    const results = json.results || [];

    // Match by IMO; fall back to first result if only one returned
    const match = results.find(r => String(r.vessel_imo) === String(vImo))
                || (results.length === 1 ? results[0] : null);

    if (!match || !match.ism_manager) return;

    const ismName = match.ism_manager     || "";
    const ismImo  = match.ism_manager_imo || "";

    // Update the visible card only if the vessel result is still showing
    if (currentView === vesselResult) {
      $("vesselIsmName").textContent = ismName;
      $("vesselIsmImo").textContent  = ismImo ? `IMO ${ismImo}` : "";
      $("vesselIsmCard").classList.remove("hidden");
    }

    // Always propagate to lastResultData so Copy as Table captures it
    if (lastResultData && lastResultData.mode === "vessel") {
      lastResultData.data.ism_manager     = ismName;
      lastResultData.data.ism_manager_imo = ismImo;
    }
  } catch {
    // ISM data is supplementary — fail silently
  }
}

function setPill(el, value) {
  el.textContent = value || "Unknown";
  el.className   = "factor-pill " + colorClass(value);
}

// ---------- Autocomplete ----------

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
  const key = Env.getSetting(SETTING_COMPANIES_KEY, "");
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
function escHtml(str) {
  return String(str ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function bannerStyle(value) {
  const c = colorClass(value);
  const bg = c === "green" ? "#16a34a" : c === "amber" ? "#d97706" : c === "red" ? "#dc2626" : "#6b7280";
  return `background:${bg};color:#fff;padding:16px;text-align:center;font-size:17px;font-weight:800;letter-spacing:0.6px;text-transform:uppercase`;
}

function pillStyle(value) {
  const c = colorClass(value);
  if (c === "green") return "background:#dcfce7;color:#15803d;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap";
  if (c === "amber") return "background:#fef3c7;color:#b45309;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap";
  if (c === "red")   return "background:#fee2e2;color:#991b1b;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap";
  return "background:#f3f4f6;color:#374151;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:700;white-space:nowrap";
}

const TABLE_HEADER_STYLE = "background:#145b76;color:#fff;padding:14px 16px";
const TABLE_ROW_STYLE    = "border-bottom:1px solid #e5e7eb";
const TABLE_CELL_STYLE   = "padding:10px 14px;font-weight:700;font-size:13px;color:#374151;text-transform:uppercase";
const TABLE_CELL_R_STYLE = "padding:10px 14px;text-align:right";

function buildCompanyTable({ imo, name, data }) {
  const url = data.shareable_url || "";
  return `<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;min-width:320px">
  <tr><td colspan="2" style="${TABLE_HEADER_STYLE}">
    <div style="font-size:16px;font-weight:700">${escHtml(name)}</div>
    <div style="font-size:13px;margin-top:2px;opacity:0.9">IMO: ${escHtml(imo)}</div>
  </td></tr>
  <tr><td colspan="2" style="${bannerStyle(data.global_performance)}">${escHtml(data.global_performance || "Unknown")}</td></tr>
  <tr><td style="padding:10px 14px;font-size:12px;color:#6b7280">Q Check Report (valid 30 days)</td>
      <td style="${TABLE_CELL_R_STYLE}">${url ? `<a href="${escHtml(url)}" style="color:#145b76;font-weight:700;font-size:13px">View Report &#8594;</a>` : ""}</td></tr>
</table>`;
}

function buildVesselTable({ vImo, vNm, data }) {
  const a       = data.assessment || {};
  const url     = data.shareable_url || "";
  const ismName = data.ism_manager     || "";
  const ismImo  = data.ism_manager_imo || "";
  // Show ISM row whenever either name or IMO is present
  const ismRow  = (ismName || ismImo)
    ? `<tr style="${TABLE_ROW_STYLE}">
    <td style="${TABLE_CELL_STYLE}">ISM Manager</td>
    <td style="${TABLE_CELL_R_STYLE}">
      ${ismName ? `<span style="font-size:13px;font-weight:600;color:#1a2332">${escHtml(ismName)}</span>` : ""}
      ${ismImo  ? `<span style="font-size:11px;color:#6b7280;display:block">IMO ${escHtml(String(ismImo))}</span>` : ""}
    </td>
  </tr>`
    : "";
  return `<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;min-width:320px">
  <tr><td colspan="2" style="${TABLE_HEADER_STYLE}">
    <div style="font-size:16px;font-weight:700">${escHtml(vNm)}</div>
    <div style="font-size:13px;margin-top:2px;opacity:0.9">IMO: ${escHtml(vImo)}</div>
  </td></tr>
  <tr><td colspan="2" style="${bannerStyle(a.global)}">${escHtml(a.global || "Unknown")}</td></tr>
  <tr style="${TABLE_ROW_STYLE}">
    <td style="${TABLE_CELL_STYLE}">Age Criteria</td>
    <td style="${TABLE_CELL_R_STYLE}"><span style="${pillStyle(a.age)}">${escHtml(a.age || "Unknown")}</span></td>
  </tr>
  <tr style="${TABLE_ROW_STYLE}">
    <td style="${TABLE_CELL_STYLE}">PSC Performance</td>
    <td style="${TABLE_CELL_R_STYLE}"><span style="${pillStyle(a.psc)}">${escHtml(a.psc || "Unknown")}</span></td>
  </tr>
  <tr style="${TABLE_ROW_STYLE}">
    <td style="${TABLE_CELL_STYLE}">Company Status</td>
    <td style="${TABLE_CELL_R_STYLE}"><span style="${pillStyle(a.company)}">${escHtml(a.company || "Unknown")}</span></td>
  </tr>
  ${ismRow}
  <tr><td style="padding:10px 14px;font-size:12px;color:#6b7280">Q Check Report (valid 30 days)</td>
      <td style="${TABLE_CELL_R_STYLE}">${url ? `<a href="${escHtml(url)}" style="color:#145b76;font-weight:700;font-size:13px">View Report &#8594;</a>` : ""}</td></tr>
</table>`;
}

function insertIntoEmail(html, btn) {
  Office.context.mailbox.item.body.setSelectedDataAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        flashBtn(btn, "Inserted!");
      } else {
        flashBtn(btn, "Failed");
      }
    }
  );
}

function copyAsTable(html, btn) {
  if (navigator.clipboard && navigator.clipboard.write) {
    const blob = new Blob([html], { type: "text/html" });
    navigator.clipboard.write([new ClipboardItem({ "text/html": blob })])
      .then(() => flashBtn(btn, "Copied!"))
      .catch(() => fallbackCopyHtml(html, btn));
  } else {
    fallbackCopyHtml(html, btn);
  }
}

function fallbackCopyHtml(html, btn) {
  const div = document.createElement("div");
  div.innerHTML = html;
  div.style.cssText = "position:fixed;opacity:0;pointer-events:none";
  document.body.appendChild(div);
  const range = document.createRange();
  range.selectNode(div);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  try { document.execCommand("copy"); flashBtn(btn, "Copied!"); }
  catch (_) { flashBtn(btn, "Copy failed"); }
  sel.removeAllRanges();
  document.body.removeChild(div);
}

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

// ---------- Zammad ticket ----------

function buildZammadTitle(d) {
  if (d.mode === "company") {
    return `Q Check Review Required : Company${d.name ? " '" + d.name : ""}' '${d.imo}'`;
  }
  return `Q Check Review Required : Vessel${d.vNm ? " '" + d.vNm : ""}' '${d.vImo}'`;
}

function buildZammadDescription(d) {
  const profile   = Env.getUserProfile();
  const userName  = profile.displayName  || "";
  const userEmail = profile.emailAddress || "";

  if (d.mode === "company") {
    const global = d.data.global_performance || "Unknown";
    const url    = d.data.shareable_url      || "";
    return [
      "Dear colleagues,",
      "",
      "Please note that the below Company requires a further Review on the Eligibility to MAR before we proceed further.",
      "",
      `**Company IMO Number** : ${d.imo}`,
      `**Company Name** : ${d.name || "Unknown"}`,
      "",
      `**Global Assessment** : ${global}`,
      "",
      `**Report Link** : ${url}`,
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
    "",
    `**Vessel IMO Number** : ${d.vImo}`,
    `**Vessel Name** : ${d.vNm || "Unknown"}`
  ];
  if (d.cImo) lines.push(`**Company IMO Number** : ${d.cImo}`);
  if (d.cNm)  lines.push(`**Company Name** : ${d.cNm}`);
  lines.push(
    "",
    `**Global Assessment** : ${global}`,
    "",
    `**Age Criteria** : ${a.age      || "Unknown"}`,
    `**PSC Performance** : ${a.psc   || "Unknown"}`,
    `**Company Status** : ${a.company || "Unknown"}`,
    "",
    `**Report Link** : ${url}`,
    "",
    "Thank you,",
    userName,
    userEmail
  );
  return lines.join("\n");
}

async function submitZammadTicket(btnEl, statusEl) {
  const zammadToken = Env.getSetting(SETTING_ZAMMAD_TOKEN, "");

  if (!zammadToken) {
    statusEl.textContent = "Zammad token not configured — open Settings.";
    statusEl.className   = "zammad-status error";
    statusEl.classList.remove("hidden");
    return;
  }

  const d         = lastResultData;
  const title     = buildZammadTitle(d);
  const body      = buildZammadDescription(d);
  const userEmail = Env.getUserProfile().emailAddress;

  const ticket = {
    title,
    state:    "new",
    group:    "Maritime Operations Department",
    customer: userEmail,
    article:  { subject: title, body, type: "note", internal: false },
    category: ["Maritime Department::Registrations / Deletions::New Registrations"]
  };

  if (d.mode === "vessel") {
    ticket.vessel_name = d.vNm || "";
    ticket.vessel_imo  = parseInt(d.vImo, 10);
  }

  btnEl.disabled    = true;
  btnEl.textContent = "Sending…";
  statusEl.classList.add("hidden");

  try {
    const resp = await fetch(`${ZAMMAD_PROXY_URL}/tickets`, {
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

// ---------- HTML escaping ----------
function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ---------- Vessel search ----------
const SHIP_SEARCH_PATH = "/api/ships/search";

function onZvlInput() {
  const q = zvlSearchInput.value.trim();
  clearTimeout(zvlSearchTimer);
  zvlDetail.classList.add("hidden");
  zvlResultsList.classList.remove("hidden");
  if (q.length < 2) {
    zvlResultsList.innerHTML = "";
    zvlStatus.textContent = q.length === 1 ? "Type at least 2 characters." : "";
    return;
  }
  zvlStatus.textContent = "Searching…";
  zvlSearchTimer = setTimeout(() => doVesselSearch(q), 300);
}

async function doVesselSearch(name) {
  const apiBase    = (Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE) || DEFAULT_API_BASE).replace(/\/$/, "");
  const searchKey  = Env.getSetting(SETTING_COMPANIES_KEY, "") || "";
  if (!searchKey) {
    zvlStatus.textContent = "Companies Search key not configured — open Settings.";
    return;
  }
  try {
    const resp = await fetch(`${apiBase}${SHIP_SEARCH_PATH}?name=${encodeURIComponent(name)}`, {
      headers: { "X-API-Key": searchKey }
    });
    if (resp.status === 401) { zvlStatus.textContent = "API key rejected (401)."; return; }
    if (!resp.ok) { zvlStatus.textContent = `Error ${resp.status}.`; return; }
    const data = await resp.json();
    zvlResults = data.results || [];
    if (!zvlResults.length) {
      zvlStatus.textContent = "No vessels found.";
      zvlResultsList.innerHTML = "";
    } else {
      zvlStatus.textContent = `${zvlResults.length} result${zvlResults.length > 1 ? "s" : ""}`;
      renderZvlResults(zvlResults);
    }
  } catch (err) {
    zvlStatus.textContent = "Network error: " + (err.message || "unknown");
  }
}

function renderZvlResults(results) {
  zvlResultsList.innerHTML = "";
  results.forEach((r) => {
    const li = document.createElement("li");
    li.className = "zvl-result-item";
    li.innerHTML = `<div class="zvl-result-name">${escHtml(r.vessel_name || "Unknown")}</div>`
                 + `<div class="zvl-result-imo">IMO ${escHtml(String(r.vessel_imo || "—"))}</div>`;
    li.addEventListener("click", () => showZvlDetail(r));
    zvlResultsList.appendChild(li);
  });
}

function showZvlDetail(r) {
  zvlResultsList.classList.add("hidden");

  const fmtNum = (v) => {
    const n = Number(String(v || "").replace(/[^\d.]/g, ""));
    return Number.isFinite(n) && n > 0 ? n.toLocaleString("en-US") : null;
  };

  const rows = [
    r.vessel_type         ? ["Type",          r.vessel_type]                                                             : null,
    r.class_society       ? ["Class Society", r.class_society]                                                           : null,
    fmtNum(r.gross_tonnage) ? ["Gross Tonnage", fmtNum(r.gross_tonnage)]                                                 : null,
    r.year_built          ? ["Year Built",    String(r.year_built)]                                                      : null,
    r.ism_manager         ? ["ISM Manager",   r.ism_manager + (r.ism_manager_imo ? ` (IMO ${r.ism_manager_imo})` : "")] : null
  ].filter(Boolean);

  zvlDetailCard.innerHTML =
    `<div class="zvl-detail-header">`
  + `<div class="zvl-detail-vessel-name">${escHtml(r.vessel_name || "Unknown")}</div>`
  + `<div class="zvl-detail-imo">IMO ${escHtml(String(r.vessel_imo || "—"))}</div>`
  + `</div><div class="zvl-detail-body">`
  + (rows.length
      ? rows.map(([k, v]) =>
          `<div class="zvl-detail-row"><span class="zvl-detail-key">${escHtml(k)}</span>`
        + `<span class="zvl-detail-val">${escHtml(v)}</span></div>`
        ).join("")
      : `<div class="zvl-detail-row"><span class="zvl-detail-key">No additional details available.</span></div>`
    )
  + `</div>`;

  if (Env.env === "extension" || new URLSearchParams(location.search).get("context") === "extension") {
    const fillBtn = document.createElement("button");
    fillBtn.className = "run-btn";
    fillBtn.style.marginTop = "12px";
    fillBtn.textContent = "Paste into Zammad Case";
    fillBtn.addEventListener("click", () => insertVesselIntoZammadTicket(r, fillBtn));
    zvlDetailCard.appendChild(fillBtn);
  }

  zvlDetail.classList.remove("hidden");
}

function insertVesselIntoZammadTicket(vessel, btn) {
  const requestId = `${Date.now()}-${Math.random()}`;
  function handleResp(event) {
    if (!event.data || event.data.type !== "zvl_fill_response" || event.data.requestId !== requestId) return;
    window.removeEventListener("message", handleResp);
    flashBtn(btn, event.data.ok ? "Filled ✓" : (event.data.error || "Fields not found"));
  }
  window.addEventListener("message", handleResp);
  window.parent.postMessage({ type: "zvl_fill", vessel, requestId }, "*");
}

// ---------- Company Search ----------

function onCsInput() {
  const q = csSearchInput.value.trim();
  clearTimeout(csSearchTimer);
  csDetail.classList.add("hidden");
  csVesselSection.classList.add("hidden");
  csResultsList.classList.remove("hidden");
  if (q.length < 2) {
    csResultsList.innerHTML = "";
    csStatus.textContent = q.length === 1 ? "Type at least 2 characters." : "";
    return;
  }
  csStatus.textContent = "Searching…";
  csSearchTimer = setTimeout(() => doCsSearch(q), 300);
}

async function doCsSearch(q) {
  const apiBase   = (Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE) || DEFAULT_API_BASE).replace(/\/$/, "");
  const searchKey = Env.getSetting(SETTING_COMPANIES_KEY, "") || "";
  if (!searchKey) {
    csStatus.textContent = "Companies Search key not configured — open Settings.";
    return;
  }
  const params = new URLSearchParams();
  if (/^\d+$/.test(q)) {
    params.set("imo", q);
  } else {
    params.set("name", q);
  }
  try {
    const resp = await fetch(`${apiBase}/api/companies/search?${params}`, {
      headers: { "X-API-Key": searchKey }
    });
    if (resp.status === 401) { csStatus.textContent = "API key rejected (401)."; return; }
    if (!resp.ok)             { csStatus.textContent = `Error ${resp.status}.`;   return; }
    const data = await resp.json();
    csResults = (data.success && Array.isArray(data.results)) ? data.results : [];
    if (!csResults.length) {
      csStatus.textContent = "No companies found.";
      csResultsList.innerHTML = "";
    } else {
      csStatus.textContent = `${csResults.length} result${csResults.length !== 1 ? "s" : ""}`;
      renderCsResults(csResults);
    }
  } catch (err) {
    csStatus.textContent = "Network error: " + (err.message || "unknown");
  }
}

function renderCsResults(results) {
  csResultsList.innerHTML = "";
  results.forEach((r) => {
    const li = document.createElement("li");
    li.className = "zvl-result-item";
    li.innerHTML = `<div class="zvl-result-name">${escHtml(r.company_name || "Unknown")}</div>`
                 + `<div class="zvl-result-imo">IMO ${escHtml(String(r.company_imo || "—"))}</div>`;
    li.addEventListener("click", () => showCsDetail(r));
    csResultsList.appendChild(li);
  });
}

function showCsDetail(company) {
  csCurrentCompany = company;
  csResultsList.classList.add("hidden");
  csVesselSection.classList.add("hidden");
  csStatus.textContent = "";

  const fleetAgeStr = company.fleet_age != null
    ? `${Number(company.fleet_age).toFixed(1)} yrs` : "—";
  const vesselQty   = company.vessels_qty != null ? String(company.vessels_qty) : "—";

  csDetailCard.innerHTML =
    `<div class="zvl-detail-header">`
  + `<div class="zvl-detail-vessel-name">${escHtml(company.company_name || "Unknown")}</div>`
  + `<div class="zvl-detail-imo">IMO ${escHtml(String(company.company_imo || "—"))}</div>`
  + `</div>`
  + `<div class="cs-stats-grid">`
  + `<div class="cs-stat-item">`
  +   `<div class="cs-stat-label">Managed Vessels</div>`
  +   `<div class="cs-stat-value">${escHtml(vesselQty)}</div>`
  + `</div>`
  + `<div class="cs-stat-item">`
  +   `<div class="cs-stat-label">Avg. Fleet Age</div>`
  +   `<div class="cs-stat-value">${escHtml(fleetAgeStr)}</div>`
  + `</div>`
  + `</div>`;

  csFleetStatus.textContent = "Loading fleet…";
  csFleetList.innerHTML     = "";
  csDetail.classList.remove("hidden");

  loadCsFleet(company.company_imo);
}

async function loadCsFleet(companyImo) {
  const apiBase   = (Env.getSetting(SETTING_API_BASE, DEFAULT_API_BASE) || DEFAULT_API_BASE).replace(/\/$/, "");
  const searchKey = Env.getSetting(SETTING_COMPANIES_KEY, "") || "";
  if (!searchKey) {
    csFleetStatus.textContent = "Companies Search key not configured.";
    return;
  }
  try {
    const resp = await fetch(
      `${apiBase}${SHIP_SEARCH_PATH}?company_imo=${encodeURIComponent(companyImo)}`,
      { headers: { "X-API-Key": searchKey } }
    );
    if (!resp.ok) {
      csFleetStatus.textContent = `Error loading fleet (${resp.status}).`;
      return;
    }
    const data    = await resp.json();
    const vessels = data.results || [];
    if (!vessels.length) {
      csFleetStatus.textContent = "No vessels found in fleet.";
      return;
    }
    csFleetStatus.textContent = `${vessels.length} vessel${vessels.length !== 1 ? "s" : ""}`;
    csFleetList.innerHTML = "";
    vessels.forEach((v) => {
      const li = document.createElement("li");
      li.className = "zvl-result-item";
      li.innerHTML = `<div class="zvl-result-name">${escHtml(v.vessel_name || "Unknown")}</div>`
                   + `<div class="zvl-result-imo">IMO ${escHtml(String(v.vessel_imo || "—"))}</div>`;
      li.addEventListener("click", () => showCsVesselDetail(v));
      csFleetList.appendChild(li);
    });
  } catch (err) {
    csFleetStatus.textContent = "Network error: " + (err.message || "unknown");
  }
}

function showCsVesselDetail(r) {
  csDetail.classList.add("hidden");

  const fmtNum = (v) => {
    const n = Number(String(v || "").replace(/[^\d.]/g, ""));
    return Number.isFinite(n) && n > 0 ? n.toLocaleString("en-US") : null;
  };

  const rows = [
    r.vessel_type             ? ["Type",          r.vessel_type]                                                             : null,
    r.class_society           ? ["Class Society", r.class_society]                                                           : null,
    fmtNum(r.gross_tonnage)   ? ["Gross Tonnage", fmtNum(r.gross_tonnage)]                                                   : null,
    r.year_built              ? ["Year Built",    String(r.year_built)]                                                      : null,
    r.ism_manager             ? ["ISM Manager",   r.ism_manager + (r.ism_manager_imo ? ` (IMO ${r.ism_manager_imo})` : "")] : null
  ].filter(Boolean);

  csVesselDetailCard.innerHTML =
    `<div class="zvl-detail-header">`
  + `<div class="zvl-detail-vessel-name">${escHtml(r.vessel_name || "Unknown")}</div>`
  + `<div class="zvl-detail-imo">IMO ${escHtml(String(r.vessel_imo || "—"))}</div>`
  + `</div><div class="zvl-detail-body">`
  + (rows.length
      ? rows.map(([k, v]) =>
          `<div class="zvl-detail-row"><span class="zvl-detail-key">${escHtml(k)}</span>`
        + `<span class="zvl-detail-val">${escHtml(v)}</span></div>`
        ).join("")
      : `<div class="zvl-detail-row"><span class="zvl-detail-key">No additional details available.</span></div>`
    )
  + `</div>`;

  csVesselSection.classList.remove("hidden");
}

// ---------- Favorites ----------

function loadFavorites() {
  try {
    const raw = Env.getSetting(SETTING_FAVORITES, "");
    favorites = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(favorites)) favorites = [];
  } catch (_) {
    favorites = [];
  }
}

async function saveFavorites() {
  try {
    await Env.setSetting(SETTING_FAVORITES, JSON.stringify(favorites));
  } catch (e) {
    console.warn("Could not persist favorites:", e);
  }
}

function toggleFavorite(toolId) {
  if (!TOOL_DEFS[toolId]) return;
  const idx = favorites.indexOf(toolId);
  if (idx === -1) {
    favorites.push(toolId);
  } else {
    favorites.splice(idx, 1);
  }
  saveFavorites();
  renderFavStars();
  renderFavoritesSection();
}

/** Update all ★/☆ star buttons already in the DOM to reflect current favorites. */
function renderFavStars() {
  document.querySelectorAll(".fav-star-btn").forEach(star => {
    const toolId = star.dataset.tool;
    const active = favorites.includes(toolId);
    star.textContent = active ? "★" : "☆"; // ★ filled / ☆ outline
    star.title = active ? "Remove from Favorites" : "Add to Favorites";
    star.classList.toggle("fav-active", active);
  });
}

/** Rebuild the Favorites section on the main page from the current favorites array. */
function renderFavoritesSection() {
  // Remove tiles generated by a previous render (keep the static label)
  favList.innerHTML = "";

  const isExt = isExtensionContext();
  const validFavs = favorites.filter(id => {
    const tool = TOOL_DEFS[id];
    if (!tool) return false;
    if (tool.extOnly && !isExt) return false;
    return true;
  });
  favoritesSection.classList.toggle("hidden", validFavs.length === 0);

  validFavs.forEach(toolId => {
    const tool = TOOL_DEFS[toolId];

    const btn = document.createElement("button");
    btn.className = "feature-tile fav-tile";
    btn.innerHTML =
      `<span class="feature-icon">${tool.icon}</span>`
    + `<div class="feature-info">`
    + `<div class="feature-name">${escHtml(tool.name)}</div>`
    + `<div class="feature-desc">${escHtml(tool.desc)}</div>`
    + `</div>`
    + `<span class="fav-star-btn fav-active" data-tool="${escHtml(toolId)}" title="Remove from Favorites">★</span>`;

    // Tile click → navigate to the tool, back button returns to mainView
    btn.addEventListener("click", () => tool.navigate(mainView));

    // Star click → remove from favorites (stop propagation so tile click isn't fired too)
    btn.querySelector(".fav-star-btn").addEventListener("click", (e) => {
      e.stopPropagation();
      toggleFavorite(toolId);
    });

    favList.appendChild(btn);
  });
}

// ---------- Seafarers — Zoho BMAR Bridge ----------

/** Show the bridge view and configure visible section based on environment. */
function showSeafarersBridgeView() {
  const isExt = isExtensionContext();
  $("sfbrExtContent").classList.toggle("hidden", !isExt);
  $("sfbrNonExtNote").classList.toggle("hidden",  isExt);
  // Reset status + button on every open
  const statusEl = $("sfbrStatus");
  if (statusEl) { statusEl.textContent = ""; statusEl.className = "sfbr-status"; }
  const injectBtn = $("sfbrInjectBtn");
  if (injectBtn) injectBtn.disabled = false;
  showView(seafarersBridgeView);
}

/** Send sfbr_inject postMessage to popup.js and wait for the response. */
function sendSfbrInject() {
  const requestId = `${Date.now()}-${Math.random()}`;
  const statusEl  = $("sfbrStatus");
  const injectBtn = $("sfbrInjectBtn");

  injectBtn.disabled    = true;
  statusEl.textContent  = "Injecting…";
  statusEl.className    = "sfbr-status info";

  let settled = false;
  function settle(ok, msg) {
    if (settled) return;
    settled = true;
    window.removeEventListener("message", handleResp);
    injectBtn.disabled   = false;
    statusEl.textContent = msg;
    statusEl.className   = "sfbr-status " + (ok ? "ok" : "error");
  }

  function handleResp(event) {
    if (!event.data || event.data.type !== "sfbr_inject_response") return;
    if (event.data.requestId !== requestId) return;
    if (event.data.ok) {
      settle(true, "✓ Bridge injected. Navigate to the Review Extracted Data step in the Seafarers Panel — the \"Send to Zoho BMAR\" button will appear there.");
    } else {
      settle(false, event.data.error || "Injection failed.");
    }
  }

  window.addEventListener("message", handleResp);
  window.parent.postMessage({ type: "sfbr_inject", requestId }, "*");

  // Timeout — if no response, probably not in extension context
  setTimeout(() => settle(false, "No response. Make sure you are using the EUROMAR Chrome Extension."), 6000);
}

// ---------- Zammad Reports ----------
const REPORT_API_URL = "https://zammad-dashboard.onrender.com/api/v1/report";

function setRptStatus(text, type) {
  rptStatus.textContent = text;
  rptStatus.className   = "rpt-status" + (type ? " " + type : "");
}

async function generateReport() {
  const dashboardKey = Env.getSetting(SETTING_DASHBOARD_KEY, "");
  if (!dashboardKey) {
    setRptStatus("Dashboard API key not configured — open Settings.", "err");
    return;
  }

  const states = ["open", "new", "With MAR/DGRM/GAMA", "With RO"];
  if (rptIncludeClosed.checked) states.push("closed", "pending close", "pending reminder");

  const filters = { state: states };
  const fromVal = rptFrom.value;
  const toVal   = rptTo.value;
  const imoVal  = rptVesselImo.value.trim();
  if (fromVal) filters.updated_from = fromVal;
  if (toVal)   filters.updated_to   = toVal;
  if (imoVal)  filters.vessel_imo   = imoVal;

  rptGenerateBtn.disabled = true;
  setRptStatus("Generating PDF…", "info");

  try {
    const resp = await fetch(REPORT_API_URL, {
      method:  "POST",
      headers: { "Authorization": `Bearer ${dashboardKey}`, "Content-Type": "application/json" },
      body:    JSON.stringify({ format: "pdf", limit: 1000, filters })
    });

    if (resp.status === 401) { setRptStatus("Authentication failed (401). Check your Zammad token.", "err"); return; }
    if (!resp.ok) {
      const txt = await resp.text().catch(() => "");
      setRptStatus(`Error ${resp.status}${txt ? ": " + txt.slice(0, 100) : ""}.`, "err");
      return;
    }

    const blob    = await resp.blob();
    const blobUrl = URL.createObjectURL(blob);

    const today = new Date().toISOString().slice(0, 10);
    let suffix = today;
    if (fromVal && toVal) suffix = `${fromVal}_to_${toVal}`;
    else if (fromVal)     suffix = `from_${fromVal}`;
    else if (toVal)       suffix = `until_${toVal}`;
    if (imoVal) suffix += `_imo${imoVal}`;

    const a = document.createElement("a");
    a.href     = blobUrl;
    a.download = `maritime-report-${suffix}.pdf`;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(blobUrl), 30000);

    setRptStatus("PDF downloaded.", "ok");
    setTimeout(() => setRptStatus("", ""), 4000);
  } catch (err) {
    setRptStatus("Network error: " + (err.message || "unknown"), "err");
  } finally {
    rptGenerateBtn.disabled = false;
  }
}
