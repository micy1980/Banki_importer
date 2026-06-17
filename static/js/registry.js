/* registry.js — MNB bank-regisztrációs tábla + BIC-lookup. */

async function lookupPartnerBank() {
  const params = new URLSearchParams({ bic: el("partnerSwift").value, iban: el("partnerIban").value });
  autoFillPartnerBankFromHuAccount();
  if (!params.get("bic") && !params.get("iban")) {
    setPartnerStatus("Adj meg SWIFT/BIC kódot vagy IBAN-t.", "warn");
    return;
  }
  setPartnerStatus("Bankadat keresés...");
  const res = await fetch(`/api/bic-lookup?${params.toString()}`);
  const data = await res.json();
  if (!res.ok || !data.found) {
    setPartnerStatus(data.error || "Nem találtam online bankadatot. Kézzel megadható.", "warn");
    return;
  }
  if (data.bank_name) el("partnerBankName").value = data.bank_name;
  if (data.bank_address) el("partnerBankAddress").value = data.bank_address;
  if (data.bic) el("partnerSwift").value = data.bic;
  setPartnerStatus(data.source_message || "Bankadat kitöltve.", "ok");
}

function renderRegistry(registry) {
  registryRows = registry.rows || [];
  const meta = [];
  const count = registry.row_count || (registry.rows || []).length || 0;
  if (registry.startup_message) meta.push(registry.startup_message);
  else if (registry.updated_at) meta.push(`Hitelesítő tábla betöltve: ${count} prefix, frissítve: ${registry.updated_at}.`);
  else meta.push("Hitelesítő tábla még nincs betöltve.");
  const box = el("registryMeta");
  box.textContent = meta.join(" ");
  box.className = `status ${registry.startup_ok === false ? "warn" : "ok"}`;
  const pill = el("registryPill");
  if (pill) {
    pill.textContent = registry.startup_ok === false
      ? "MNB tábla: helyi adat"
      : `MNB tábla: ${count || 0} prefix`;
    pill.className = `status-pill ${registry.startup_ok === false ? "warn" : "ok"}`;
    pill.title = meta.join(" ");
  }
}

async function loadRegistry() {
  const res = await fetch("/api/bank-registry");
  const data = await res.json();
  renderRegistry(data);
  autoFillBankNameFromAccount();
  autoFillBankNameFromEditAccount();
}
