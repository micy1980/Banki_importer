/* Banki UI - Searchable enhancer for #companySelect.
   Native <select> stays intact (event/value contract). When >8 companies,
   a search input is inserted next to it that toggles option visibility. */
(function () {
  "use strict";
  const MIN_OPTIONS = 8;
  let wired = false;

  function enhance() {
    const sel = document.getElementById("companySelect");
    if (!sel) return;
    const count = sel.querySelectorAll("option").length;
    if (count < MIN_OPTIONS) { teardown(); return; }
    if (wired) return;
    wired = true;
    sel.classList.add("combobox-select");
    const wrap = document.createElement("div");
    wrap.className = "combobox-wrap";
    sel.parentElement.insertBefore(wrap, sel);
    const input = document.createElement("input");
    input.type = "search";
    input.className = "combobox-input";
    input.placeholder = "Cég keresése (begépelve)...";
    input.setAttribute("aria-label", "Cég keresése");
    input.setAttribute("aria-controls", "companySelect");
    wrap.appendChild(input);
    wrap.appendChild(sel);

    input.addEventListener("input", () => {
      const q = input.value.trim().toLowerCase();
      let firstMatch = null;
      sel.querySelectorAll("option").forEach(opt => {
        const hit = !q || opt.textContent.toLowerCase().includes(q);
        opt.hidden = !hit;
        opt.disabled = !hit;
        if (hit && !firstMatch) firstMatch = opt;
      });
      if (firstMatch && q) sel.value = firstMatch.value;
    });
    input.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === "ArrowDown") {
        sel.focus();
        e.preventDefault();
      }
    });
  }
  function teardown() {
    const sel = document.getElementById("companySelect");
    if (!sel) return;
    sel.querySelectorAll("option").forEach(o => { o.hidden = false; o.disabled = false; });
  }
  function init() {
    enhance();
    const sel = document.getElementById("companySelect");
    if (sel) {
      new MutationObserver(() => { wired = false; enhance(); })
        .observe(sel, { childList: true });
    }
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
