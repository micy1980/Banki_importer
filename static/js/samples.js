/* samples.js - import sample table, result summary and error list rendering. */
function importEmptyState() {
  return `
    <div class="empty-state empty-state--compact" role="status">
      <strong>Még nincs beolvasott adat</strong>
      <span>Nyisd meg az Import ablakot, válassz bankot, formátumot és fájlt.</span>
      <div class="empty-state-actions">
        <button class="secondary" type="button" data-open-import>Import megnyitása</button>
        <a class="button-link ghost" href="/template.xlsx" download>Sablon letöltése</a>
      </div>
    </div>
  `;
}

function renderSample() {
  const headers = window.currentInspect?.headers || [];
  const rows = window.currentInspect?.sample || [];
  if (!headers.length) {
    el("sampleArea").innerHTML = importEmptyState();
    return;
  }
  let html = "<table><thead><tr>";
  for (const h of headers) html += `<th>${escapeHtml(h)}</th>`;
  html += "</tr></thead><tbody>";
  for (const row of rows) {
    html += "<tr>";
    for (let index = 0; index < headers.length; index++) {
      html += `<td data-label="${escapeHtml(headers[index])}">${escapeHtml(row[index] ?? "")}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table>";
  el("sampleArea").innerHTML = html;
}

function renderResultSummary(data = null) {
  const format = selectedFormat();
  const bank = BANKS[el("bankSelect")?.value] || {};
  const rows = [
    [data?.headers?.length ?? "-", "Fejléc"],
    [data?.data_rows ?? "-", "Adatsor"],
    [bank.label || "-", "Bank"],
    [format?.label || format?.short_label || "-", "Formátum"],
  ];
  el("resultSummary").innerHTML = rows.map(([value, label]) =>
    `<div class="metric"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>`
  ).join("");
}

function renderErrors(errors) {
  const list = Array.isArray(errors) ? errors : String(errors || "").split("\n").filter(Boolean);
  if (!list.length) { el("errorArea").innerHTML = ""; return; }
  el("errorArea").innerHTML = `
    <div class="error-summary" role="alert" tabindex="-1">
      <strong>${list.length} javítandó hiba</strong>
      <ul>${list.map(error => `<li>${escapeHtml(error)}</li>`).join("")}</ul>
    </div>
  `;
  el("errorArea").querySelector(".error-summary")?.focus();
}

window.importEmptyState = importEmptyState;
window.renderSample = renderSample;
window.renderResultSummary = renderResultSummary;
window.renderErrors = renderErrors;
