/* Banki UI - Jump-to-row from error list to sample table. */
(function () {
  "use strict";
  function init() {
    const errArea = document.getElementById("errorArea");
    const sample = document.getElementById("sampleArea");
    if (!errArea || !sample) return;
    new MutationObserver(() => {
      errArea.querySelectorAll("li, .error-item").forEach(item => {
        if (item.dataset.jumpAttached) return;
        const m = (item.textContent || "").match(/sor[: ]?\s*(\d+)/i);
        if (!m) return;
        item.dataset.jumpAttached = "1";
        const btn = document.createElement("button");
        btn.type = "button"; btn.className = "row-jump"; btn.textContent = "Ugrás";
        btn.onclick = () => {
          const rowIdx = parseInt(m[1], 10);
          const trs = sample.querySelectorAll("tbody tr");
          const tr = trs[rowIdx - 1] || trs[rowIdx - 2];
          if (!tr) { window.bankiToast?.("Sor nincs a mintában.", "warn"); return; }
          sample.querySelectorAll(".row-highlight").forEach(r => r.classList.remove("row-highlight"));
          tr.classList.add("row-highlight");
          tr.scrollIntoView({ behavior: "smooth", block: "center" });
          setTimeout(() => tr.classList.remove("row-highlight"), 4000);
        };
        item.appendChild(document.createTextNode(" "));
        item.appendChild(btn);
      });
    }).observe(errArea, { childList: true, subtree: true });
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
