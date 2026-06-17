/* convert.js — config gyűjtés és PAYORD TXT generálás. */

function collectConfig() {
  const mapping = {};
  const defaults = {};
  document.querySelectorAll("[data-map]").forEach(sel => mapping[sel.dataset.map] = sel.value);
  document.querySelectorAll("[data-default]").forEach(inp => defaults[inp.dataset.default] = inp.value);
  return {
    company_id: activeCompanyId(),
    bank: el("bankSelect").value || "",
    format: el("formatSelect").value || "",
    encoding: el("encoding").value,
    identifier_date: el("identifierDate").value,
    mapping,
    defaults
  };
}

async function convertFile() {
  const file = el("fileInput").files[0];
  if (!file || !currentInspect) {
    setStatus("Előbb olvasd be a fájlt.", "warn");
    return;
  }
  const form = new FormData();
  form.append("file", file);
  form.append("config", JSON.stringify(collectConfig()));
  setStatus("Konvertálás...");
  setButtonLoading("convertBtn", true, "TXT készül...");
  try {
    const res = await fetch("/api/convert", { method: "POST", body: form });
    const contentType = res.headers.get("content-type") || "";
    if (!res.ok) {
      const data = contentType.includes("application/json") ? await res.json() : {error: await res.text()};
      setStatus(data.error || "Nem sikerült konvertálni.", "bad");
      renderErrors(data.errors || data.error || []);
      return;
    }
    const buffer = await res.arrayBuffer();
    const blob = new Blob([buffer], { type: contentType || "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = res.headers.get("x-filename") || "payord_import.txt";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);

    const decoderLabel = el("encoding").value === "cp1250" ? "windows-1250" : (el("encoding").value === "cp852" ? "ibm852" : "utf-8");
    let text = "";
    try { text = new TextDecoder(decoderLabel).decode(buffer); }
    catch { text = new TextDecoder("utf-8").decode(buffer); }
    el("previewBox").textContent = text.slice(0, 700).replace(/\r/g, "\\r").replace(/\n/g, "\\n\n");
    const rows = res.headers.get("x-record-count") || "?";
    setStatus(`${rows} rekord elkészült és letöltődött.`, "ok");
    renderErrors([]);
    saveSettingsDebounced();
  } finally {
    setButtonLoading("convertBtn", false);
    updateConvertAction();
  }
}


// Event listeners — DOMContentLoaded gondoskodik róla, hogy minden globális (más modulokból is) készen álljon.
document.addEventListener("DOMContentLoaded", () => {
  el("inspectBtn").addEventListener("click", () => inspectFile().catch(err => setStatus(err.message, "bad")));
  el("convertBtn").addEventListener("click", () => convertFile().catch(err => setStatus(err.message, "bad")));
  el("bankSelect").addEventListener("change", populateFormats);
  el("formatSelect").addEventListener("change", applySelectedFormat);
  el("fileInput").addEventListener("change", () => {
    currentInspect = null;
    renderSample();
    renderResultSummary();
    renderErrors([]);
    updateConvertAction();
  });
  el("sampleArea").addEventListener("click", event => {
    if (event.target?.closest("[data-open-import]")) {
      openDialog("importDialog", event.target.closest("[data-open-import]"));
    }
  });
  el("openImportBtn").addEventListener("click", event => openDialog("importDialog", event.currentTarget));
  el("companiesBtn").addEventListener("click", async event => {
    openDialog("companiesDialog", event.currentTarget);
    await loadCompanies();
  });
  el("partnersBtn").addEventListener("click", async event => {
    openDialog("partnersDialog", event.currentTarget);
    await Promise.all([loadPartners(), loadRegistry()]);
  });
  el("accountsBtn").addEventListener("click", async event => {
    openDialog("accountsDialog", event.currentTarget);
    await Promise.all([loadAccounts(), loadRegistry()]);
  });
  el("helpBtn").addEventListener("click", event => openDialog("helpDialog", event.currentTarget));
  el("closeImportBtn").addEventListener("click", () => closeDialog("importDialog"));
  el("closeAccountsBtn").addEventListener("click", () => closeDialog("accountsDialog"));
  el("closeCompaniesBtn").addEventListener("click", () => closeDialog("companiesDialog"));
  el("closePartnersBtn").addEventListener("click", () => closeDialog("partnersDialog"));
  el("closePartnerEditBtn").addEventListener("click", () => {
    editingPartnerId = "";
    closeDialog("partnerEditDialog");
  });
  el("cancelPartnerEditBtn").addEventListener("click", () => {
    editingPartnerId = "";
    closeDialog("partnerEditDialog");
  });
  el("closeAccountEditBtn").addEventListener("click", () => {
    editingAccountId = "";
    closeDialog("accountEditDialog");
  });
  el("cancelAccountEditBtn").addEventListener("click", () => {
    editingAccountId = "";
    closeDialog("accountEditDialog");
  });
  el("closeHelpBtn").addEventListener("click", () => closeDialog("helpDialog"));
  el("saveAccountBtn").addEventListener("click", () => saveAccount().catch(err => setAccountStatus(err.message, "bad")));
  el("saveAccountEditBtn").addEventListener("click", () => saveEditedAccount().catch(err => setAccountEditStatus(err.message, "bad")));
  el("importAccountsBtn").addEventListener("click", () => importAccounts().catch(err => setAccountStatus(err.message, "bad")));
  el("saveCompanyBtn").addEventListener("click", () => saveCompany().catch(err => setCompanyStatus(err.message, "bad")));
  el("savePartnerBtn").addEventListener("click", () => savePartner().catch(err => setPartnerStatus(err.message, "bad")));
  el("savePartnerEditBtn").addEventListener("click", () => saveEditedPartner().catch(err => setPartnerEditStatus(err.message, "bad")));
  el("importPartnersBtn").addEventListener("click", () => importPartners().catch(err => setPartnerStatus(err.message, "bad")));
  el("lookupPartnerBankBtn").addEventListener("click", () => lookupPartnerBank().catch(err => setPartnerStatus(err.message, "bad")));
  el("companySelect").addEventListener("change", () => changeCompany(activeCompanyId()).catch(err => setStatus(err.message, "bad")));
  el("ownAccountNumber").addEventListener("input", () => {
    formatAccountNumberInput();
    updateAccountValidationHint();
    autoFillBankNameFromAccount();
  });
  el("ownBankName").addEventListener("input", () => {
    if (el("ownBankName").value !== lastAutoBankName) lastAutoBankName = "";
  });
  el("editAccountNumber").addEventListener("input", () => {
    formatAccountNumberElement(el("editAccountNumber"));
    updateEditAccountFormatting();
    autoFillBankNameFromEditAccount();
  });
  el("editBankName").addEventListener("input", () => {
    if (el("editBankName").value !== lastEditAutoBankName) lastEditAutoBankName = "";
  });
  el("accountsList").addEventListener("click", event => {
    const editId = event.target?.dataset?.editAccount;
    const deleteId = event.target?.dataset?.deleteAccount;
    if (editId) editAccount(editId);
    if (deleteId) deleteAccount(deleteId).catch(err => setAccountStatus(err.message, "bad"));
  });
  el("companiesList").addEventListener("click", event => {
    const companyId = event.target?.dataset?.setCompany;
    if (companyId) {
      el("companySelect").value = companyId;
      changeCompany(companyId).catch(err => setCompanyStatus(err.message, "bad"));
    }
  });
  el("partnersList").addEventListener("click", event => {
    const editId = event.target?.dataset?.editPartner;
    const deleteId = event.target?.dataset?.deletePartner;
    if (editId) editPartner(editId);
    if (deleteId) deletePartner(deleteId).catch(err => setPartnerStatus(err.message, "bad"));
  });
  el("partnerAccount").addEventListener("input", () => {
    formatAccountNumberElement(el("partnerAccount"));
    autoFillPartnerBankFromHuAccount();
  });
  el("partnerIban").addEventListener("input", autoFillPartnerBankFromHuAccount);
  el("partnerSwift").addEventListener("blur", () => {
    if (el("partnerSwift").value.trim()) lookupPartnerBank().catch(() => {});
  });
  el("editPartnerAccount").addEventListener("input", () => formatAccountNumberElement(el("editPartnerAccount")));
});
