/* partners.js — partner adatok (CRUD, import, edit, BIC autofill). */

function setPartnerStatus(text, kind = "") {
  const box = el("partnerStatus");
  box.textContent = text;
  box.className = `status ${kind}`;
}

function clearPartnerForm() {
  for (const id of ["partnerCode", "partnerName", "partnerAccount", "partnerIban", "partnerSwift", "partnerAddress", "partnerBankName", "partnerBankAddress"]) {
    el(id).value = "";
  }
  el("partnerCountry").value = "HU";
}

function autoFillPartnerBankFromHuAccount() {
  const prefix = accountPrefixFromInput(el("partnerAccount").value || el("partnerIban").value);
  if (prefix.length !== 8) return;
  const found = registryRows.find(row => String(row.prefix || "") === prefix);
  if (found?.bank_name && !el("partnerBankName").value) {
    el("partnerBankName").value = found.bank_name;
  }
}

function renderPartners(partners) {
  partnersState = partners || [];
  renderListState("partnersList", partnersState, partner => `
    <div class="account-row">
      <div>
        <strong>${escapeHtml(partner.name || "Nincs név")}</strong>
        <span>${escapeHtml(partner.partner_code || "")}</span>
      </div>
      <div>
        <strong>${escapeHtml(partner.account_number || partner.iban || "")}</strong>
        <span>${escapeHtml(partner.swift_bic || partner.bank_name || "")}</span>
      </div>
      <div class="account-actions" style="margin:0;">
        <button class="secondary" type="button" data-edit-partner="${escapeHtml(partner.id)}">Szerkesztés</button>
        <button class="ghost" type="button" data-delete-partner="${escapeHtml(partner.id)}">Törlés</button>
      </div>
    </div>
  `, "Nincs partner", "Az aktív cég partnerlistája üres. Rögzíts új partnert vagy importáld Excelből.");
}

async function loadPartners() {
  el("partnersList").innerHTML = loadingRows("Partnerek betöltése...");
  const data = await fetchJson(`/api/partners?company_id=${encodeURIComponent(activeCompanyId())}`);
  renderPartners(data.partners || []);
}

function collectPartnerPayload() {
  return {
    company_id: activeCompanyId(),
    partner_code: el("partnerCode").value,
    name: el("partnerName").value,
    account_number: el("partnerAccount").value,
    iban: el("partnerIban").value,
    swift_bic: el("partnerSwift").value,
    country: el("partnerCountry").value,
    address: el("partnerAddress").value,
    bank_name: el("partnerBankName").value,
    bank_address: el("partnerBankAddress").value
  };
}

async function savePartner() {
  setPartnerStatus("Mentés...");
  setButtonLoading("savePartnerBtn", true, "Mentés...");
  try {
    const data = await fetchJson(`/api/partners?company_id=${encodeURIComponent(activeCompanyId())}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(collectPartnerPayload())
    });
    renderPartners(data.partners || []);
    clearPartnerForm();
    setPartnerStatus("Partner mentve.", "ok");
  } finally {
    setButtonLoading("savePartnerBtn", false);
  }
}

async function deletePartner(id) {
  const res = await fetch(`/api/partners/delete?company_id=${encodeURIComponent(activeCompanyId())}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id, company_id: activeCompanyId() })
  });
  const data = await res.json();
  renderPartners(data.partners || []);
  setPartnerStatus("Partner törölve.", "ok");
}

async function importPartners() {
  const file = el("partnerImportFile").files[0];
  if (!file) {
    setPartnerStatus("Válassz ki import fájlt.", "warn");
    return;
  }
  const form = new FormData();
  form.append("file", file);
  setPartnerStatus("Importálás...");
  const res = await fetch(`/api/partners/import?company_id=${encodeURIComponent(activeCompanyId())}`, { method: "POST", body: form });
  const data = await res.json();
  if (!res.ok) {
    setPartnerStatus(data.error || "Nem sikerült importálni.", "bad");
    return;
  }
  renderPartners(data.partners || []);
  const errorText = (data.errors || []).length ? ` Hibák: ${(data.errors || []).slice(0, 3).join("; ")}` : "";
  setPartnerStatus(`${data.added || 0} partner importálva.${errorText}`, errorText ? "warn" : "ok");
}

function setPartnerEditStatus(text, kind = "") {
  const box = el("partnerEditStatus");
  box.textContent = text;
  box.className = `status ${kind}`;
}

function editPartner(id) {
  const partner = partnersState.find(item => item.id === id);
  if (!partner) return;
  editingPartnerId = id;
  el("editPartnerCode").value = partner.partner_code || "";
  el("editPartnerName").value = partner.name || "";
  el("editPartnerAccount").value = partner.account_number || "";
  el("editPartnerIban").value = partner.iban || "";
  el("editPartnerSwift").value = partner.swift_bic || "";
  el("editPartnerCountry").value = partner.country || "HU";
  el("editPartnerAddress").value = partner.address || "";
  el("editPartnerBankName").value = partner.bank_name || "";
  el("editPartnerBankAddress").value = partner.bank_address || "";
  setPartnerEditStatus("Szerkesztésre megnyitva.");
  openDialog("partnerEditDialog");
}

function collectEditedPartnerPayload() {
  return {
    id: editingPartnerId,
    company_id: activeCompanyId(),
    partner_code: el("editPartnerCode").value,
    name: el("editPartnerName").value,
    account_number: el("editPartnerAccount").value,
    iban: el("editPartnerIban").value,
    swift_bic: el("editPartnerSwift").value,
    country: el("editPartnerCountry").value,
    address: el("editPartnerAddress").value,
    bank_name: el("editPartnerBankName").value,
    bank_address: el("editPartnerBankAddress").value
  };
}

async function saveEditedPartner() {
  if (!editingPartnerId) {
    setPartnerEditStatus("Nincs szerkesztésre kiválasztott partner.", "bad");
    return;
  }
  setPartnerEditStatus("Mentés...");
  setButtonLoading("savePartnerEditBtn", true, "Mentés...");
  try {
    const data = await fetchJson(`/api/partners?company_id=${encodeURIComponent(activeCompanyId())}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(collectEditedPartnerPayload())
    });
    renderPartners(data.partners || []);
    editingPartnerId = "";
    el("partnerEditDialog").close();
    setPartnerStatus("Partner módosítva.", "ok");
  } finally {
    setButtonLoading("savePartnerEditBtn", false);
  }
}
