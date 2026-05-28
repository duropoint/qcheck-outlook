// sra-zoho.js — Read SRA data from URL hash and fill the Zoho registration form.
// Injected into the Zoho form tab by the EUROMAR Toolkit via exec-on-tab.

(function () {
  var hashMatch = window.location.hash.match(/[#&]sraData=([^&]*)/);
  if (!hashMatch) return;

  var data;
  try { data = JSON.parse(atob(hashMatch[1])); }
  catch (e) { console.error("[SRA] Could not parse sraData from hash"); return; }

  function findFormDoc() {
    var iframes = document.querySelectorAll("iframe");
    for (var i = 0; i < iframes.length; i++) {
      try {
        var doc = iframes[i].contentDocument || iframes[i].contentWindow.document;
        if (doc.querySelector('input[name="SingleLine12"]')) return doc;
      } catch (_) {}
    }
    return document.querySelector('input[name="SingleLine12"]') ? document : null;
  }

  function fillForm(doc) {
    var win    = doc.defaultView || window;
    var proto  = win.HTMLInputElement.prototype;
    var setter = Object.getOwnPropertyDescriptor(proto, "value").set;

    function fillField(name, value) {
      if (!value) return;
      var field = doc.querySelector('input[name="' + name + '"]');
      if (!field) return;
      field.focus();
      setter.call(field, value);
      field.dispatchEvent(new Event("input",  { bubbles: true }));
      field.dispatchEvent(new Event("change", { bubbles: true }));
      field.dispatchEvent(new Event("blur",   { bubbles: true }));
    }

    var sraFields = ["SingleLine12", "SingleLine13", "SingleLine14", "SingleLine15"];
    var nums      = data.sraNumbers || [];

    nums.forEach(function (num, i) {
      if (i < sraFields.length) {
        (function (n, fn) { setTimeout(function () { fillField(fn, n); }, i * 150); })(num, sraFields[i]);
      }
    });

    var delay = nums.length * 150 + 200;
    setTimeout(function () { fillField("Date26", data.issuingDate || ""); }, delay);
    setTimeout(function () { fillField("Date27", data.validUntil  || ""); }, delay + 200);

    // Success toast
    setTimeout(function () {
      var el = document.createElement("div");
      el.style.cssText = [
        "position:fixed", "top:20px", "right:20px",
        "background:linear-gradient(135deg,#667eea,#764ba2)",
        "color:#fff", "padding:14px 22px", "border-radius:8px",
        "z-index:9999999", "font-family:-apple-system,sans-serif",
        "font-size:13px", "font-weight:600",
        "box-shadow:0 4px 16px rgba(0,0,0,.25)"
      ].join(";");
      el.textContent = "✓ SRA fields filled!";
      document.body.appendChild(el);
      setTimeout(function () { el.remove(); }, 3000);
    }, delay + 600);
  }

  function tryFill(attempts) {
    var doc = findFormDoc();
    if (doc) { fillForm(doc); return; }
    if (attempts > 0) setTimeout(function () { tryFill(attempts - 1); }, 800);
    else console.warn("[SRA] Form fields (SingleLine12) not found after retries.");
  }

  // Give the form time to render before trying to fill
  setTimeout(function () { tryFill(10); }, 1000);
})();
