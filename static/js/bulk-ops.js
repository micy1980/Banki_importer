/* Banki UI - Bulk select + delete + Excel export for partner/account/company lists. */
(function () {
  "use strict";
  const TARGETS = {
    partnersList: { endpoint: "/api/partners/delete", label: "Partnerek", filenamePrefix: "partnerek" },
    accountsList: { endpoint: "/api/accounts/delete", label: "Bankszámlák", filenamePrefix: "bankszamlak" },
    companiesList: { endpoint: "/api/companies/delete", label: "Cégek", filenamePrefix: "cegek" },
  };
  function attachBar(host, cfg) {
    if (!host || host.dataset.bulkAttached) return;
    host.dataset.bulkAttached = "1";
    const bar = document.createElement("div");
    bar.className = "bulk-bar";
    bar.innerHTML = `
      <label><input type="checkbox" data-bulk-all> Kijelölés</label>
      <span class="bulk-count" data-bulk-count>0 kijelölve</span>
      <button type="button" class="ghost" data-bulk-export>Excel export</button>
      <button type="button" class="secondary" data-bulk-delete disabled>Kijelöltek törlése</button>`;
    host.parentElement.insertBefore(bar, host);
    function rows() { return [...host.querySelectorAll("tbody tr, .account-list-row, .row, li")].filter(r=>r.offsetParent!==null||r.querySelector("td,div")); }
    function injectCheckboxes() {
      rows().forEach(r => {
        if (r.querySelector('[data-bulk-row]')) return;
        const cb = document.createElement("input");
        cb.type = "checkbox"; cb.dataset.bulkRow = "1"; cb.style.marginRight = "8px";
        const id = r.dataset.id || r.querySelector("[data-id]")?.dataset.id || r.id || "";
        cb.dataset.id = id;
        const first = r.querySelector("td, div, span");
        first?.prepend(cb);
        cb.addEventListener("change", updateCount);
      });
    }
    function updateCount() {
      const sel = host.querySelectorAll('[data-bulk-row]:checked').length;
      bar.querySelector("[data-bulk-count]").textContent = `${sel} kijelölve`;
      bar.querySelector("[data-bulk-delete]").disabled = sel === 0;
    }
    bar.querySelector("[data-bulk-all]").addEventListener("change", e => {
      host.querySelectorAll('[data-bulk-row]').forEach(cb => cb.checked = e.target.checked);
      updateCount();
    });
    bar.querySelector("[data-bulk-delete]").addEventListener("click", async () => {
      const ids = [...host.querySelectorAll('[data-bulk-row]:checked')].map(cb => cb.dataset.id).filter(Boolean);
      if (!ids.length || !confirm(`${ids.length} elem törlése?`)) return;
      const company_id = document.getElementById("companySelect")?.value || "default";
      for (const id of ids) {
        await fetch(cfg.endpoint, { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({ id, company_id })});
      }
      window.bankiToast?.(`${ids.length} elem törölve.`, "ok");
      // refresh
      document.getElementById({partnersList:"partnersBtn",accountsList:"accountsBtn",companiesList:"companiesBtn"}[host.id])?.click();
    });
    bar.querySelector("[data-bulk-export]").addEventListener("click", () => exportCsv(host, cfg));
    new MutationObserver(injectCheckboxes).observe(host, { childList: true, subtree: true });
    injectCheckboxes();
  }
  function exportCsv(host, cfg) {
    const headers = [...host.querySelectorAll("thead th")].map(th => (th.textContent||"").trim()).filter(Boolean);
    const rows = [...host.querySelectorAll("tbody tr")].map(tr =>
      [...tr.querySelectorAll("td")].map(td => (td.textContent||"").replace(/"/g,'""').trim()));
    if (!rows.length) { window.bankiToast?.("Nincs exportálható sor.", "warn"); return; }
    const csv = [headers.join(";"), ...rows.map(r => r.map(c => `"${c}"`).join(";"))].join("\r\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${cfg.filenamePrefix}-${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    window.bankiToast?.("Excel/CSV export letöltve.", "ok");
  }
  function scan() { Object.entries(TARGETS).forEach(([id, cfg]) => attachBar(document.getElementById(id), cfg)); }
  document.addEventListener("click", e => {
    if (e.target?.id && /Btn$/.test(e.target.id)) setTimeout(scan, 120);
  });
  if (document.readyState !== "loading") scan();
  else document.addEventListener("DOMContentLoaded", scan);
})();
