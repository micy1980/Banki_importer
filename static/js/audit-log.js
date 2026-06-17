/* Banki UI - Audit log viewer. Adds a small "Napló" link to the header. */
(function () {
  "use strict";
  function inject() {
    if (document.getElementById("auditBtn")) return;
    const host = document.querySelector(".command-actions");
    if (!host) return;
    const btn = document.createElement("button");
    btn.id = "auditBtn"; btn.type = "button"; btn.className = "ghost";
    btn.textContent = "Napló";
    btn.title = "Mentések és törlések naplója";
    host.appendChild(btn);
    btn.addEventListener("click", show);
  }
  async function show() {
    let dlg = document.getElementById("auditDialog");
    if (!dlg) {
      dlg = document.createElement("dialog");
      dlg.id = "auditDialog"; dlg.className = "preview-dialog";
      dlg.innerHTML = `
        <div class="dialog-head">
          <h2>Audit napló</h2>
          <button type="button" class="secondary" data-close>Bezárás</button>
        </div>
        <div class="dialog-body">
          <div id="auditBody" class="audit-body"><em>Betöltés…</em></div>
        </div>`;
      document.body.appendChild(dlg);
      dlg.addEventListener("click", e => { if (e.target===dlg||e.target.dataset.close!==undefined) dlg.close(); });
    }
    const body = dlg.querySelector("#auditBody");
    body.textContent = "Betöltés…";
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open","");
    try {
      const r = await fetch("/api/audit?limit=200");
      const data = await r.json();
      const rows = data.entries || [];
      if (!rows.length) { body.innerHTML = "<em>A napló üres.</em>"; return; }
      body.innerHTML = `<table class="audit-table"><thead><tr><th>Időpont</th><th>Művelet</th><th>Cég</th><th>Részlet</th></tr></thead><tbody>${
        rows.reverse().map(e => `<tr><td>${new Date(e.ts).toLocaleString("hu-HU")}</td><td>${e.action}</td><td>${e.company_id||""}</td><td><code>${(e.detail||"").replace(/[<>&]/g, c=>({ "<":"&lt;",">":"&gt;","&":"&amp;"}[c]))}</code></td></tr>`).join("")
      }</tbody></table>`;
    } catch (e) {
      body.textContent = "Napló nem elérhető: " + (e.message || e);
    }
  }
  if (document.readyState !== "loading") inject();
  else document.addEventListener("DOMContentLoaded", inject);
})();
