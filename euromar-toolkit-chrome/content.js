// EUROMAR Toolkit — Zammad content script
// Receives zvl_fill messages from the side panel and fills ticket fields.

(function () {
  "use strict";

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type !== "zvl_fill") return;

    const vessel  = msg.vessel;
    const details = buildVesselDetails(vessel);

    const filledName    = fillField(["vessel_name",    "vesselname",    "vessel-name"],    vessel.vessel_name  || "");
    const filledImo     = fillField(["vessel_imo",     "vesselimo",     "vessel-imo", "imo"], String(vessel.vessel_imo || ""));
    const filledDetails = fillField(["vessel_details", "vesseldetails", "vessel-details"], details);

    sendResponse({ ok: filledName || filledImo || filledDetails });
    return true;
  });

  // ── Field locator ─────────────────────────────────────────────────────────────
  // Strategy 1: input[name] / input[data-name]
  // Strategy 2: label text match → nearest input/textarea

  function fillField(names, value) {
    if (!value) return false;

    for (const n of names) {
      const el = document.querySelector(
        `input[name="${n}"], input[data-name="${n}"], textarea[name="${n}"], textarea[data-name="${n}"]`
      );
      if (el) { setNativeValue(el, value); return true; }
    }

    const labels = document.querySelectorAll("label, .form-group label, .controls label");
    for (const lbl of labels) {
      const txt = (lbl.textContent || "").trim().toLowerCase();
      const wantsName    = names.some(n => n.includes("name"))    && txt.includes("vessel") && txt.includes("name");
      const wantsImo     = names.some(n => n.includes("imo"))     && txt.includes("imo")    && !txt.includes("manager");
      const wantsDetails = names.some(n => n.includes("details")) && txt.includes("vessel") && txt.includes("details");
      if (wantsName || wantsImo || wantsDetails) {
        const container = lbl.closest(".form-group, .controls, .row") || lbl.parentElement;
        const input = container?.querySelector("input, textarea");
        if (input) { setNativeValue(input, value); return true; }
      }
    }
    return false;
  }

  // React/Vue-friendly setter — fires native events so Zammad registers the change
  function setNativeValue(el, value) {
    const proto  = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
    setter.call(el, value);
    el.dispatchEvent(new Event("input",  { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.dispatchEvent(new Event("blur",   { bubbles: true }));
  }

  // ── Vessel details summary ────────────────────────────────────────────────────

  function buildVesselDetails(r) {
    const today = new Date().toISOString().slice(0, 10);
    const gt    = r.gross_tonnage ? formatNumber(r.gross_tonnage) : "";
    return [
      r.vessel_type   ? `Vessel Type: ${r.vessel_type}`                                                              : null,
      r.class_society ? `Class Society: ${r.class_society}`                                                          : null,
      (gt || r.year_built) ? `GT: ${gt || "—"} | Built: ${r.year_built || "—"}`                                     : null,
      r.ism_manager   ? `ISM Manager: ${r.ism_manager}${r.ism_manager_imo ? ` (IMO ${r.ism_manager_imo})` : ""}`    : null,
      `Updated: ${today}`
    ].filter(Boolean).join("\n");
  }

  function formatNumber(v) {
    const n = Number(String(v).replace(/[^\d.-]/g, ""));
    return Number.isFinite(n) ? n.toLocaleString("en-US") : String(v);
  }
})();
