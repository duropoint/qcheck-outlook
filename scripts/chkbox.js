if (!location.hostname.includes("bmar.pt")) {
  throw new Error("Active tab is not a BMAR page. Open the BMAR application form and try again.");
}
const boxes = document.querySelectorAll(".ui-chkbox-box");
for (const box of boxes) {
  if (!box.querySelector(".ui-icon-check")) {
    box.click();
  }
}
