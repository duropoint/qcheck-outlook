const SFBR_ORIGINS = [
  "https://seafarers.eu-registry.com",
  "https://seafarers-web-test.idego.io"
];
if (!SFBR_ORIGINS.includes(location.origin)) {
  throw new Error("Active tab is not the Seafarers Panel. Open seafarers.eu-registry.com and try again.");
}
const boxes = document.querySelectorAll(".ui-chkbox-box");
for (const box of boxes) {
  if (!box.querySelector(".ui-icon-check")) {
    box.click();
  }
}
