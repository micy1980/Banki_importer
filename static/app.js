/* app.js — app bootstrap és a maradék mapping / fájl-inspect logika.
 * A modulok (core-dom, companies, samples, accounts, partners, registry,
 * convert, dialogs) a saját moduljaikban élnek és window.* exportokon át
 * érhetők el.
 */
const BANKS = window.__BANKS;
const FORMATS = window.__FORMATS;
const FIELDS = window.__FIELDS;
const CURRENCIES = window.__CURRENCIES;
var currentInspect = null;
var currentSettings = { active_bank: "erste", active_format: "erste_huf_payord", formats: {} };
var saveTimer = null;
var editingAccountId = "";
var editingPartnerId = "";
var accountsState = [];
var companiesState = [];
var partnersState = [];
var registryRows = [];
var lastAutoBankName = "";

const today = new Date();
el("identifierDate").value = today.toISOString().slice(0, 10);

function populateBanks() {
  const bankSelect = el("bankSelect");
  bankSelect.innerHTML = "";
  for (const [id, bank] of Object.entries(BANKS)) {
    const option = document.createElement("option");
    option.value = id;
    option.textContent = bank.label;
    bankSelect.appendChild(option);
  }
}

function populateCurrencySelect(select, selected = "HUF") {
  select.innerHTML = "";
  for (const code of CURRENCIES) {
    const option = document.createElement("option");
    option.value = code;
    option.textContent = code;
    select.appendChild(option);
  }
  select.value = CURRENCIES.includes(selected) ? selected : "HUF";
}

function populateCurrencySelects() {
  populateCurrencySelect(el("ownCurrency"), "HUF");
  populateCurrencySelect(el("editCurrency"), "HUF");
}

function populateFormats() {
  const bankId = el("bankSelect").value;
  const desired = currentSettings.active_format || el("formatSelect").value;
  const formatSelect = el("formatSelect");
  formatSelect.innerHTML = "";
  for (const [id, format] of Object.entries(FORMATS)) {
    if (format.bank !== bankId) continue;
    const option = document.createElement("option");
    option.value = id;
    option.textContent = format.label;
    formatSelect.appendChild(option);
  }
  if ([...formatSelect.options].some(o => o.value === desired)) {
    formatSelect.value = desired;
  } else if (formatSelect.options.length) {
    formatSelect.selectedIndex = 0;
  }
  applySelectedFormat();
}

function selectedFormat() {
  return FORMATS[el("formatSelect").value] || Object.values(FORMATS)[0];
}

function currentFields() {
  return selectedFormat()?.fields || FIELDS;
}

function applySelectedFormat() {
  const format = selectedFormat();
  if (!format) return;
  el("formatBadge").textContent = format.badge || format.short_label || format.label;
  el("commandFormatName").textContent = format.short_label || format.label;
  el("formatHelp").textContent = format.description;
  el("templateLink").href = `/template.xlsx?format=${encodeURIComponent(el("formatSelect").value)}`;
  el("templateLinkInDialog").href = `/template.xlsx?format=${encodeURIComponent(el("formatSelect").value)}`;
  const settings = formatSettings();
  el("encoding").value = settings.encoding || format.default_encoding || "cp1250";
  if (settings.identifier_date) el("identifierDate").value = settings.identifier_date;
  buildMappings();
  renderResultSummary(currentInspect);
  updateConvertAction();
  saveSettingsDebounced();
}

function formatSettings(formatId = el("formatSelect").value) {
  currentSettings.formats ||= {};
  currentSettings.formats[formatId] ||= { mapping: {}, defaults: {} };
  currentSettings.formats[formatId].mapping ||= {};
  currentSettings.formats[formatId].defaults ||= {};
  return currentSettings.formats[formatId];
}

async function loadSettings() {
  try {
    const res = await fetch("/api/settings");
    if (res.ok) currentSettings = await res.json();
  } catch {}
}

function saveSettingsDebounced() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveSettings, 350);
}

async function saveSettings() {
  const formatId = el("formatSelect").value || "erste_huf_payord";
  currentSettings.active_company_id = activeCompanyId();
  currentSettings.active_bank = el("bankSelect").value || "erste";
  currentSettings.active_format = formatId;
  currentSettings.formats ||= {};
  const existing = formatSettings(formatId);
  const collected = collectConfig();
  if (!document.querySelector("[data-map]")) {
    collected.mapping = existing.mapping || {};
    collected.defaults = existing.defaults || {};
  }
  currentSettings.formats[formatId] = collected;
  try {
    await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(currentSettings)
    });
  } catch {
    setStatus("A beállítás mentése nem sikerült.", "warn");
  }
}

function setStatus(text, kind = "") {
  const box = el("statusBox");
  box.textContent = text;
  box.dataset.kind = kind;
}

function setDisabledReason(buttonId, reason = "") {
  const button = el(buttonId);
  const reasonBox = el(`${buttonId.replace("Btn", "")}DisabledReason`);
  if (!button) return;
  const disabled = Boolean(reason);
  button.disabled = disabled;
  button.title = reason || "";
  button.setAttribute("aria-disabled", disabled ? "true" : "false");
  if (reasonBox) reasonBox.textContent = reason || "";
}

function getConvertDisabledReason() {
  if (!el("fileInput").files[0]) return "Tölts fel egy Excel vagy CSV fájlt az Import panelen.";
  if (!currentInspect) return "Olvasd be a fájlt, hogy ellenőrizni lehessen a fejlécet és a sorokat.";
  if (requiredMappingMissing()) return "Ellenőrizd a kötelező Excel oszlop-hozzárendeléseket a Haladó részben.";
  return "";
}

function updateConvertAction() {
  setDisabledReason("convertBtn", getConvertDisabledReason());
}

function buildMappings() {
  const headers = currentInspect?.headers || [];
  const guesses = currentInspect?.guesses || {};
  const fields = currentFields();
  const settings = formatSettings();
  const area = el("mappingArea");
  area.innerHTML = "";
  for (const [key, info] of Object.entries(fields)) {
    const row = document.createElement("div");
    row.className = "mapping";
    const required = info.required ? `<span class="req">*</span>` : "";
    const guessed = settings.mapping[key] || guesses[key] || "";
    const defaultValue = settings.defaults[key] ?? defaultFor(key);
    row.innerHTML = `
      <div>
        <strong>${escapeHtml(info.label)} ${required}</strong>
        <div class="map-help">${escapeHtml(info.help || "Excel oszlop hozzárendelése ehhez a fix banki mezőhöz.")}</div>
      </div>
      <select data-map="${escapeHtml(key)}">${optionList(headers, guessed)}</select>
      <input class="default-input" data-default="${escapeHtml(key)}" type="text"
        value="${escapeHtml(defaultValue)}" placeholder="Fix érték, ha nincs Excel oszlop">
    `;
    area.appendChild(row);
  }
}

function requiredMappingMissing() {
  const headers = currentInspect?.headers || [];
  const guesses = currentInspect?.guesses || {};
  const settings = formatSettings();
  return Object.entries(currentFields()).some(([key, info]) => {
    if (!info.required) return false;
    const mapped = settings.mapping[key] || guesses[key] || "";
    const fallback = settings.defaults[key] ?? defaultFor(key);
    return !mapped && !fallback && headers.length > 0;
  });
}

function defaultFor(key) {
  if (el("formatSelect")?.value === "erste_sepa_payord" && ["currency", "payout_currency"].includes(key)) return "EUR";
  if (el("formatSelect")?.value === "erste_sepa_payord" && key === "decimals") return "2";
  if (el("formatSelect")?.value === "unicredit_fx_ccy" && key === "cost_bearer") return "SHA";
  if (el("formatSelect")?.value === "unicredit_fx_ccy" && key === "legal_code") return "000";
  const defaults = {
    sender_currency: "HUF", currency: "HUF", payout_currency: "EUR", decimals: "0", status: "",
    beneficiary_country: "HU", beneficiary_bank_country: "", sender_account_type: "0",
    beneficiary_account_type: "0", swift_copy: "N", custom_rate_use: "N", urgent_use: "N",
    urgent_execution: "N", process_mode: "", group_transfer: "N", hold_flag: "N", chqb_flag: "N",
    deal_ticket_flag: "N", cost_bearer: "1", commission_bearer: "0", other_fee_bearer: "0",
    amount_currency_mode: " ", payment_method: " ", priority: " ", item_type: " ", iban_flag: " ",
    document_no: "", legal_code: "", debtor_name: "", external_ref: "", internal_note: ""
  };
  return defaults[key] ?? "";
}

async function inspectFile() {
  const file = el("fileInput").files[0];
  if (!file) {
    setStatus("Előbb válassz ki egy fájlt.", "warn");
    updateConvertAction();
    return;
  }
  const form = new FormData();
  form.append("file", file);
  setStatus("Beolvasás...");
  setButtonLoading("inspectBtn", true, "Beolvasás...");
  try {
    const data = await fetchJson("/api/inspect", { method: "POST", body: form });
    currentInspect = data;
    buildMappings();
    el("mappingDetails").open = requiredMappingMissing();
    renderSample();
    updateConvertAction();
    renderResultSummary(data);
    renderErrors([]);
    const bank = BANKS[el("bankSelect").value]?.label || el("bankSelect").value;
    const format = selectedFormat()?.label || selectedFormat()?.short_label || el("formatSelect").value;
    setStatus(`${bank} - ${format}: ${data.headers.length} fejléc beolvasva, ${data.data_rows} adatsor észlelve.`, "ok");
    saveSettingsDebounced();
  } catch (err) {
    currentInspect = null;
    updateConvertAction();
    setStatus(err.message || "Nem sikerült beolvasni a fájlt.", "bad");
    renderErrors([err.message || "Nem sikerült beolvasni a fájlt."]);
  } finally {
    setButtonLoading("inspectBtn", false);
    updateConvertAction();
  }
}

el("useGuessesBtn").addEventListener("click", () => {
  const settings = formatSettings();
  settings.mapping = {...(currentInspect?.guesses || {})};
  buildMappings();
  saveSettingsDebounced();
});
el("mappingArea").addEventListener("change", saveSettingsDebounced);
el("mappingArea").addEventListener("change", updateConvertAction);
el("mappingArea").addEventListener("input", () => {
  saveSettingsDebounced();
  updateConvertAction();
});
el("encoding").addEventListener("change", saveSettingsDebounced);
el("identifierDate").addEventListener("change", saveSettingsDebounced);

(async function initApp() {
  populateBanks();
  populateCurrencySelects();
  await loadSettings();
  await loadCompanies();
  if (currentSettings.active_company_id && [...el("companySelect").options].some(o => o.value === currentSettings.active_company_id)) {
    el("companySelect").value = currentSettings.active_company_id;
  }
  if (BANKS[currentSettings.active_bank]) el("bankSelect").value = currentSettings.active_bank;
  populateFormats();
  if (FORMATS[currentSettings.active_format]) el("formatSelect").value = currentSettings.active_format;
  applySelectedFormat();
  renderResultSummary();
  renderSample();
  updateConvertAction();
  setupDialogA11y();
  makeDialogsDraggable();
  loadRegistry().catch(() => {
    const pill = el("registryPill");
    if (pill) { pill.textContent = "MNB tábla: nem elérhető"; pill.className = "status-pill warn"; }
  });
})();
