const ALLOWED = [
  "seafarers.eu-registry.com",
  "seafarers-web-test.idego.io",
  "bmar.pt"
];
if (!ALLOWED.some(h => location.hostname.includes(h))) {
  throw new Error("Active tab is not a supported page. Open the BMAR form or Seafarers Panel and try again.");
}
const boxes = document.querySelectorAll(".ui-chkbox-box");
for (const box of boxes) {
  if (!box.classList.contains("ui-state-active")) {
    box.click();
  }
}
