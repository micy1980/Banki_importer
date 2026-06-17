/* Banki UI - Drag & drop dropzone around the file input. */
(function () {
  "use strict";
  function wrap(id) {
    const input = document.getElementById(id);
    if (!input || input.dataset.dropzoneApplied) return;
    const zone = document.createElement("div");
    zone.className = "dropzone";
    zone.innerHTML = `<span>Húzd ide a fájlt, vagy kattints a tallózáshoz</span>`;
    input.parentElement.insertBefore(zone, input);
    zone.appendChild(input);
    input.dataset.dropzoneApplied = "1";
    zone.addEventListener("click", e => { if (e.target === zone || e.target.tagName === "SPAN") input.click(); });
    ["dragenter","dragover"].forEach(ev => zone.addEventListener(ev, e => {
      e.preventDefault(); zone.classList.add("dropzone--over");
    }));
    ["dragleave","drop"].forEach(ev => zone.addEventListener(ev, e => {
      e.preventDefault(); zone.classList.remove("dropzone--over");
    }));
    zone.addEventListener("drop", e => {
      const f = e.dataTransfer?.files?.[0];
      if (!f) return;
      const dt = new DataTransfer(); dt.items.add(f); input.files = dt.files;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }
  function init() { ["fileInput","accountImportFile","partnerImportFile"].forEach(wrap); }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
  document.addEventListener("click", e => {
    if (e.target?.id && /Btn$/.test(e.target.id)) setTimeout(init, 80);
  });
})();
