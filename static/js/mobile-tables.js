/* Banki UI - Mobile card view: copy header text to td[data-label]. */
(function () {
  "use strict";
  function sync(table) {
    const heads = table.querySelectorAll("thead th, tr:first-child th");
    if (!heads.length) return;
    const labels = [...heads].map(th => th.textContent.trim());
    table.querySelectorAll("tr").forEach(tr => {
      if (tr.querySelector("th")) return;
      tr.querySelectorAll("td").forEach((td, i) => {
        if (labels[i] && !td.hasAttribute("data-label")) td.setAttribute("data-label", labels[i]);
      });
    });
  }
  function init() {
    ["partnersList", "accountsList", "companiesList"].forEach(id => {
      const list = document.getElementById(id);
      const table = list && list.closest("table");
      if (!table) return;
      if (!table.id) table.id = id + "Table";
      table.classList.add("table--cards-on-mobile");
      if (table.dataset.cardObserver === "1") return;
      table.dataset.cardObserver = "1";
      new MutationObserver(() => sync(table)).observe(table, { childList: true, subtree: true });
      sync(table);
    });
  }
  document.addEventListener("click", e => {
    if (e.target.closest("#partnersBtn,#accountsBtn,#companiesBtn")) setTimeout(init, 300);
  });
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
