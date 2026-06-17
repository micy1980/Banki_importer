/* samples.js — import minta-tábla, eredmény összefoglaló és hiba lista renderelése.
 * Globálisak: renderSample, renderResultSummary, renderErrors, importEmptyState.
 */
function importEmptyState() {
  return `
    <div class="empty-state" role="status">
      <strong>Még nincs beolvasott adat</strong>
      <span>A konvertálás három rövid lépésből áll. Először válaszd ki az aktív céget, utána tölts fel egy Excel vagy CSV fájlt az Import panelen.</span>
      <ol>
        <li>Cég kiválasztása vagy létrehozása.</li>
        <li>Import panel megnyitása, bank és formátum ellenőrzése.</li>
        <li>Fájl beolvasása, majd TXT letöltése.</li>
      </ol>
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
  for (const r of rows) {
    html += "<tr>";
    for (let i = 0; i < headers.length; i++) {
      html += `<td data-label="${escapeHtml(headers[i])}">${escapeHtml(r[i] ?? "")}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table>";
  el("sampleArea").innerHTML = html;
}

function renderResultSummary(data = null) {
  const format = selectedFormat();
  const rows = [
    [data?.headers?.length ?? "-", "Fejléc"],
    [data?.data_rows ?? "-", "Adatsor"],
    [data?.errors?.length ?? 0, "Hibás sor"],
    [format?.short_label ?? "-", "Formátum"]
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
      <ul>${list.map(e => `<li>${escapeHtml(e)}</li>`).join("")}</ul>
    </div>
  `;
  el("errorArea").querySelector(".error-summary")?.focus();
}

window.importEmptyState = importEmptyState;
window.renderSample = renderSample;
window.renderResultSummary = renderResultSummary;
window.renderErrors = renderErrors;
