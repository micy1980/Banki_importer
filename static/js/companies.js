/* companies.js - multi-company registry module. */
var editingCompanyId = "";

function activeCompanyId() {
  return el("companySelect")?.value || "default";
}

function activeCompanyName() {
  const found = (window.companiesState || []).find(company => company.id === activeCompanyId());
  return found?.name || "Alap cég";
}

async function loadCompanies() {
  if (el("companiesList")) el("companiesList").innerHTML = loadingRows("Cégek betöltése...");
  const data = await fetchJson("/api/companies");
  window.companiesState = data.companies || [];
  const select = el("companySelect");
  const current = data.active_company_id || window.currentSettings?.active_company_id || activeCompanyId();
  select.innerHTML = "";
  for (const company of window.companiesState) {
    const option = document.createElement("option");
    option.value = company.id;
    option.textContent = company.name;
    select.appendChild(option);
  }
  if ([...select.options].some(option => option.value === current)) select.value = current;
  else if (select.options.length) select.selectedIndex = 0;
  renderCompanies();
}

function renderCompanies() {
  if (!el("companiesList")) return;
  renderListState("companiesList", window.companiesState || [], company => `
    <div class="account-row" data-id="${escapeHtml(company.id)}">
      <div><strong>${escapeHtml(company.name)}</strong><span>${company.id === activeCompanyId() ? "Aktív" : "Cég"}</span></div>
      <div><strong>${escapeHtml(company.id)}</strong><span>Azonosító</span></div>
      <div class="account-actions" style="margin:0;">
        <button class="secondary" type="button" data-edit-company="${escapeHtml(company.id)}">Szerkesztés</button>
      </div>
    </div>
  `, "Nincs rögzített cég", "Adj hozzá céget, hogy cégenként külön számlákat és partnereket kezelj.");
}

async function saveCompany() {
  const name = el("companyName").value.trim();
  if (!name) { setCompanyStatus("Adj meg cégnevet.", "bad"); return; }
  setButtonLoading("saveCompanyBtn", true, "Mentés...");
  try {
    const data = await fetchJson("/api/companies", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name }),
    });
    window.companiesState = data.companies || [];
    el("companyName").value = "";
    await loadCompanies();
    el("companySelect").value = data.company.id;
    await changeCompany(data.company.id);
    setCompanyStatus("Cég mentve.", "ok");
  } finally {
    setButtonLoading("saveCompanyBtn", false);
  }
}

function setCompanyStatus(text, kind = "") {
  const box = el("companyStatus");
  if (!box) return;
  box.textContent = text;
  box.className = `status ${kind}`;
}

function setCompanyEditStatus(text, kind = "") {
  const box = el("companyEditStatus");
  if (!box) return;
  box.textContent = text;
  box.className = `status ${kind}`;
}

function editCompany(companyId) {
  const company = (window.companiesState || []).find(row => row.id === companyId);
  if (!company) return;
  editingCompanyId = company.id;
  el("editCompanyName").value = company.name || "";
  setCompanyEditStatus("");
  openDialog("companyEditDialog", document.querySelector(`[data-edit-company="${CSS.escape(companyId)}"]`));
}

async function saveEditedCompany() {
  const name = el("editCompanyName").value.trim();
  if (!editingCompanyId) { setCompanyEditStatus("Nincs kiválasztott cég.", "bad"); return; }
  if (!name) { setCompanyEditStatus("Adj meg cégnevet.", "bad"); return; }
  setButtonLoading("saveCompanyEditBtn", true, "Mentés...");
  try {
    const data = await fetchJson("/api/companies", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: editingCompanyId, name }),
    });
    window.companiesState = data.companies || [];
    await loadCompanies();
    el("companySelect").value = data.company.id;
    await changeCompany(data.company.id);
    editingCompanyId = "";
    closeDialog("companyEditDialog");
    setCompanyStatus("Cég mentve.", "ok");
  } finally {
    setButtonLoading("saveCompanyEditBtn", false);
  }
}

async function changeCompany(companyId) {
  window.currentSettings.active_company_id = companyId || activeCompanyId();
  await saveSettings();
  renderCompanies();
  if (el("accountsDialog")?.open) await loadAccounts();
  if (el("partnersDialog")?.open) await loadPartners();
  setStatus(`Aktív cég: ${activeCompanyName()}`, "ok");
}

window.activeCompanyId = activeCompanyId;
window.activeCompanyName = activeCompanyName;
window.loadCompanies = loadCompanies;
window.renderCompanies = renderCompanies;
window.saveCompany = saveCompany;
window.setCompanyStatus = setCompanyStatus;
window.setCompanyEditStatus = setCompanyEditStatus;
window.editCompany = editCompany;
window.saveEditedCompany = saveEditedCompany;
window.changeCompany = changeCompany;
