/* Banki UI - Client-side search filter for partner/account/companies lists. */
(function () {
  "use strict";
  function attach(inputId, listId) {
    const input = document.getElementById(inputId);
    const list  = document.getElementById(listId);
    if (!input || !list) return;
    input.addEventListener("input", () => {
      const q = input.value.trim().toLowerCase();
      let visible = 0;
      list.querySelectorAll("tr").forEach((tr, i) => {
        if (i === 0 && tr.querySelector("th")) return;
        const hit = !q || tr.textContent.toLowerCase().includes(q);
        tr.style.display = hit ? "" : "none";
        if (hit) visible++;
      });
      const meta = document.getElementById(inputId + "Meta");
      if (meta) meta.textContent = q ? `${visible} találat` : "";
    });
  }
  function inject(dialogId, listId, placeholder) {
    const dlg = document.getElementById(dialogId);
    if (!dlg) return;
    const body = dlg.querySelector(".dialog-body");
    if (!body || body.querySelector(`[data-filter-for="${listId}"]`)) return;
    const list = document.getElementById(listId); if (!list) return;
    const wrap = document.createElement("div");
    wrap.className = "list-filter"; wrap.dataset.filterFor = listId;
    const inputId = listId + "Filter";
    wrap.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      <input id="${inputId}" type="search" placeholder="${placeholder}" autocomplete="off">
      <span class="list-filter__meta" id="${inputId}Meta"></span>`;
    const anchor = list.closest("table") || list;
    anchor.parentElement.insertBefore(wrap, anchor);
    attach(inputId, listId);
  }
  function tryInject() {
    inject("partnersDialog", "partnersList", "Keresés név, számla, BIC...");
    inject("accountsDialog", "accountsList", "Keresés név, számla...");
    inject("companiesDialog", "companiesList", "Cég keresése...");
  }
  function init() {
    tryInject();
    document.querySelectorAll("dialog").forEach(d => {
      new MutationObserver(tryInject).observe(d, { attributes: true, attributeFilter: ["open"] });
    });
    document.addEventListener("click", e => {
      if (e.target.closest("#partnersBtn,#accountsBtn,#companiesBtn")) setTimeout(tryInject, 250);
    });
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
