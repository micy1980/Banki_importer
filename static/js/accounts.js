/* accounts.js — saját bankszámlák kezelése (CRUD, import, validáció). */

function clientFormatHuAccount(value) {
  const compact = String(value || "").replace(/[^0-9]/g, "");
  const groups = compact.slice(0, 24).match(/.{1,8}/g) || [];
  return groups.join("-");
}

function formatAccountNumberElement(input) {
  const value = input.value;
  const compact = value.replace(/[\s-]+/g, "").toUpperCase();
  if (/^[A-Z]/.test(compact)) {
    input.value = compact.replace(/[^A-Z0-9]/g, "").slice(0, 34).match(/.{1,4}/g)?.join(" ") || "";
  } else {
    input.value = clientFormatHuAccount(value);
  }
}

function formatAccountNumberInput() {
  formatAccountNumberElement(el("ownAccountNumber"));
}

function accountPrefixFromInput(value) {
  const compact = String(value || "").replace(/[\s-]+/g, "").toUpperCase();
  if (compact.startsWith("HU") && compact.length >= 12) {
    return compact.slice(4, 12).replace(/\D/g, "");
  }
  if (/^[A-Z]/.test(compact)) {
    return "";
  }
  return compact.replace(/\D/g, "").slice(0, 8);
}

function autoFillBankNameFor(accountInput, bankInput, countryValue, lastValue) {
  if (countryValue !== "HU") return lastValue;
  const prefix = accountPrefixFromInput(accountInput.value);
  if (prefix.length !== 8) {
    if (bankInput.value === lastValue) bankInput.value = "";
    return "";
  }
  const found = registryRows.find(row => String(row.prefix || "") === prefix);
  if (!found?.bank_name) return lastValue;
  if (!bankInput.value || bankInput.value === lastValue) {
    bankInput.value = found.bank_name;
    return found.bank_name;
  }
  return lastValue;
}

function autoFillBankNameFromAccount() {
  lastAutoBankName = autoFillBankNameFor(el("ownAccountNumber"), el("ownBankName"), el("ownBankCountry").value, lastAutoBankName);
}

function autoFillBankNameFromEditAccount() {
  lastEditAutoBankName = autoFillBankNameFor(el("editAccountNumber"), el("editBankName"), el("editBankCountry").value, lastEditAutoBankName);
}

function updateAccountValidationHint() {
  const value = el("ownAccountNumber").value.trim();
  const compact = value.replace(/[\s-]+/g, "").toUpperCase();
  if (!value || /^[A-Z]{2}/.test(compact)) {
    return;
  }
  el("ownAccountNumber").value = clientFormatHuAccount(value);
}

function updateEditAccountFormatting() {
  const value = el("editAccountNumber").value.trim();
  const compact = value.replace(/[\s-]+/g, "").toUpperCase();
  if (!value || /^[A-Z]{2}/.test(compact)) {
    return;
  }
  el("editAccountNumber").value = clientFormatHuAccount(value);
}

function setAccountStatus(text, kind = "") {
  const box = el("accountStatus");
  box.textContent = text;
  box.className = `status ${kind}`;
}

function setAccountEditStatus(text, kind = "") {
  const box = el("accountEditStatus");
  box.textContent = text;
  box.className = `status ${kind}`;
}

function clearAccountForm() {
  lastAutoBankName = "";
  el("ownBankCountry").value = "HU";
  el("ownBankName").value = "";
  el("ownCurrency").value = "HUF";
  el("ownAccountNumber").value = "";
  updateAccountValidationHint();
  setAccountStatus("Nincs kiválasztott bankszámla.");
}

function renderAccounts(accounts) {
  accountsState = accounts || [];
  el("accountsCompanyName").textContent = activeCompanyName();
  renderListState("accountsList", accountsState, account => `
    <div class="account-row" data-id="${escapeHtml(account.id)}">
      <div>
        <strong>${escapeHtml(account.bank_name || "Nincs banknév")}</strong>
        <span>${escapeHtml(account.bank_country || "HU")}</span>
      </div>
      <div>
        <strong>${escapeHtml(account.account_number || "")}</strong>
        <span>${escapeHtml(account.currency || "HUF")}</span>
      </div>
      <div class="account-actions" style="margin:0;">
        <button class="secondary" type="button" data-edit-account="${escapeHtml(account.id)}">Szerkesztés</button>
        <button class="ghost" type="button" data-delete-account="${escapeHtml(account.id)}">Törlés</button>
      </div>
    </div>
  `, "Nincs saját bankszámla", "Az aktív céghez még nincs rögzített bankszámla. Rögzíts kézzel vagy importálj sablonból.");
}

async function loadAccounts() {
  el("accountsList").innerHTML = loadingRows("Saját bankszámlák betöltése...");
  const data = await fetchJson(`/api/accounts?company_id=${encodeURIComponent(activeCompanyId())}`);
  renderAccounts(data.accounts || []);
}

async function saveAccount() {
  const payload = {
    id: "",
    company_id: activeCompanyId(),
    bank_country: el("ownBankCountry").value,
    bank_name: el("ownBankName").value,
    currency: el("ownCurrency").value,
    account_number: el("ownAccountNumber").value
  };
  setAccountStatus("Mentés...");
  setButtonLoading("saveAccountBtn", true, "Mentés...");
  try {
    const data = await fetchJson(`/api/accounts?company_id=${encodeURIComponent(activeCompanyId())}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    renderAccounts(data.accounts || []);
    clearAccountForm();
    setAccountStatus("Bankszámla mentve és validálva.", "ok");
  } finally {
    setButtonLoading("saveAccountBtn", false);
  }
}

async function deleteAccount(id) {
  const res = await fetch(`/api/accounts/delete?company_id=${encodeURIComponent(activeCompanyId())}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ id })
  });
  const data = await res.json();
  renderAccounts(data.accounts || []);
  setAccountStatus("Bankszámla törölve.", "ok");
}

function editAccount(id) {
  const account = accountsState.find(item => item.id === id);
  if (!account) return;
  editingAccountId = id;
  el("editBankCountry").value = account.bank_country || "HU";
  el("editBankName").value = account.bank_name || "";
  el("editCurrency").value = account.currency || "HUF";
  el("editAccountNumber").value = account.account_number || "";
  lastEditAutoBankName = account.bank_name || "";
  updateEditAccountFormatting();
  setAccountEditStatus("Szerkesztésre megnyitva.");
  openDialog("accountEditDialog");
}

async function saveEditedAccount() {
  if (!editingAccountId) {
    setAccountEditStatus("Nincs szerkesztésre kiválasztott bankszámla.", "bad");
    return;
  }
  const payload = {
    id: editingAccountId,
    company_id: activeCompanyId(),
    bank_country: el("editBankCountry").value,
    bank_name: el("editBankName").value,
    currency: el("editCurrency").value,
    account_number: el("editAccountNumber").value
  };
  setAccountEditStatus("Mentés...");
  const res = await fetch(`/api/accounts?company_id=${encodeURIComponent(activeCompanyId())}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await res.json();
  if (!res.ok) {
    setAccountEditStatus(data.error || "Nem sikerült menteni.", "bad");
    return;
  }
  renderAccounts(data.accounts || []);
  editingAccountId = "";
  lastEditAutoBankName = "";
  el("accountEditDialog").close();
  setAccountStatus("Bankszámla módosítva.", "ok");
}

async function importAccounts() {
  const file = el("accountImportFile").files[0];
  if (!file) {
    setAccountStatus("Válassz ki egy import fájlt.", "warn");
    return;
  }
  const form = new FormData();
  form.append("file", file);
  setAccountStatus("Importálás...");
  const res = await fetch(`/api/accounts/import?company_id=${encodeURIComponent(activeCompanyId())}`, { method: "POST", body: form });
  const data = await res.json();
  if (!res.ok) {
    setAccountStatus(data.error || "Nem sikerült importálni.", "bad");
    return;
  }
  renderAccounts(data.accounts || []);
  const errorText = (data.errors || []).length ? ` Hibák: ${(data.errors || []).join("; ")}` : "";
  setAccountStatus(`${data.added || 0} bankszámla importálva.${errorText}`, errorText ? "warn" : "ok");
}
