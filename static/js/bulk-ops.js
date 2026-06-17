/* Banki UI - bulk select, delete and CSV export for registry lists. */
(function () {
  "use strict";

  const TARGETS = {
    partnersList: { endpoint: "/api/partners/delete", label: "Partnerek", filenamePrefix: "partnerek" },
    accountsList: { endpoint: "/api/accounts/delete", label: "Bankszámlák", filenamePrefix: "bankszamlak" },
    companiesList: { endpoint: "/api/companies/delete", label: "Cégek", filenamePrefix: "cegek" },
  };

  function listRows(host) {
    return [...host.querySelectorAll("tbody tr, .account-row")]
      .filter(row => !row.classList.contains("empty-state"));
  }

  function rowId(row) {
    return row.dataset.id
      || row.querySelector("[data-id]")?.dataset.id
      || row.querySelector("[data-edit-account]")?.dataset.editAccount
      || row.querySelector("[data-delete-account]")?.dataset.deleteAccount
      || row.querySelector("[data-edit-partner]")?.dataset.editPartner
      || row.querySelector("[data-delete-partner]")?.dataset.deletePartner
      || row.querySelector("[data-edit-company]")?.dataset.editCompany
      || row.id
      || "";
  }

  function rowValues(row) {
    if (row.matches("tr")) {
      return [...row.querySelectorAll("td")].map(cell => (cell.textContent || "").trim());
    }
    return [...row.querySelectorAll(":scope > div")]
      .map(cell => (cell.textContent || "").replace(/\s+/g, " ").trim())
      .filter(Boolean);
  }

  function attachBar(host, cfg) {
    if (!host || host.dataset.bulkAttached) return;
    host.dataset.bulkAttached = "1";

    const bar = document.createElement("div");
    bar.className = "bulk-bar";
    bar.innerHTML = `
      <label class="bulk-check"><input type="checkbox" data-bulk-all> <span>Kijelölés</span></label>
      <span class="bulk-count" data-bulk-count>0 kijelölve</span>
      <button type="button" class="ghost" data-bulk-export>Excel export</button>
      <button type="button" class="secondary" data-bulk-delete disabled>Kijelöltek törlése</button>`;
    host.parentElement.insertBefore(bar, host);

    const master = bar.querySelector("[data-bulk-all]");
    const counter = bar.querySelector("[data-bulk-count]");
    const deleteButton = bar.querySelector("[data-bulk-delete]");

    function updateCount() {
      const checkboxes = [...host.querySelectorAll("[data-bulk-row]")];
      const selected = checkboxes.filter(checkbox => checkbox.checked);
      counter.textContent = `${selected.length} kijelölve`;
      deleteButton.disabled = selected.length === 0;
      master.checked = checkboxes.length > 0 && selected.length === checkboxes.length;
      master.indeterminate = selected.length > 0 && selected.length < checkboxes.length;
    }

    function injectCheckboxes() {
      for (const row of listRows(host)) {
        if (row.querySelector("[data-bulk-row]")) continue;
        const id = rowId(row);
        if (!id) continue;
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.dataset.bulkRow = "1";
        checkbox.dataset.id = id;
        checkbox.setAttribute("aria-label", "Sor kijelölése");
        checkbox.addEventListener("change", updateCount);

        if (row.matches("tr")) {
          const cell = row.querySelector("td");
          cell?.prepend(checkbox);
        } else {
          row.prepend(checkbox);
        }
      }
      updateCount();
    }

    master.addEventListener("change", event => {
      host.querySelectorAll("[data-bulk-row]").forEach(checkbox => {
        checkbox.checked = event.target.checked;
      });
      updateCount();
    });

    deleteButton.addEventListener("click", async () => {
      const ids = [...host.querySelectorAll("[data-bulk-row]:checked")]
        .map(checkbox => checkbox.dataset.id)
        .filter(Boolean);
      if (!ids.length || !confirm(`${ids.length} elem törlése?`)) return;
      const company_id = document.getElementById("companySelect")?.value || "default";
      for (const id of ids) {
        await fetch(cfg.endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id, company_id }),
        });
      }
      window.bankiToast?.(`${ids.length} elem törölve.`, "ok");
      refreshList(host.id);
    });

    bar.querySelector("[data-bulk-export]").addEventListener("click", () => exportCsv(host, cfg));
    new MutationObserver(injectCheckboxes).observe(host, { childList: true, subtree: true });
    injectCheckboxes();
  }

  function refreshList(hostId) {
    if (hostId === "accountsList" && window.loadAccounts) window.loadAccounts();
    else if (hostId === "partnersList" && window.loadPartners) window.loadPartners();
    else if (hostId === "companiesList" && window.loadCompanies) window.loadCompanies();
  }

  function exportCsv(host, cfg) {
    const tableHeaders = [...host.querySelectorAll("thead th")]
      .map(th => (th.textContent || "").trim())
      .filter(Boolean);
    const rows = listRows(host).map(rowValues).filter(row => row.length);
    if (!rows.length) {
      window.bankiToast?.("Nincs exportálható sor.", "warn");
      return;
    }
    const width = Math.max(...rows.map(row => row.length), tableHeaders.length);
    const headers = tableHeaders.length ? tableHeaders : Array.from({ length: width }, (_, index) => `Oszlop ${index + 1}`);
    const csv = [headers, ...rows].map(row =>
      row.map(value => `"${String(value || "").replace(/"/g, '""')}"`).join(";")
    ).join("\r\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${cfg.filenamePrefix}-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
    window.bankiToast?.(`${cfg.label} export letöltve.`, "ok");
  }

  function scan() {
    Object.entries(TARGETS).forEach(([id, cfg]) => attachBar(document.getElementById(id), cfg));
  }

  document.addEventListener("click", event => {
    if (event.target?.id && /Btn$/.test(event.target.id)) setTimeout(scan, 120);
  });
  if (document.readyState !== "loading") scan();
  else document.addEventListener("DOMContentLoaded", scan);
})();
