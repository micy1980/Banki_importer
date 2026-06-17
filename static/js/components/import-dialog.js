/* <import-dialog id="importDialog"> */
class ImportDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
    <dialog aria-labelledby="importDialogTitle">
      <div class="dialog-head">
        <h2 id="importDialogTitle">Import</h2>
        <button id="closeImportBtn" class="secondary" type="button" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <div class="import-grid">
          <div><label for="bankSelect">Bank</label><select id="bankSelect"></select></div>
          <div>
            <label for="formatSelect">Formátum</label>
            <select id="formatSelect"></select>
            <div id="formatHelp" class="hint">A bank által elvárt TXT rekordforma.</div>
          </div>
          <div class="full-row">
            <label for="fileInput">Fájl</label>
            <input id="fileInput" type="file" accept=".xlsx,.xlsm,.csv">
          </div>
          <div>
            <label for="encoding">TXT kódolás</label>
            <select id="encoding">
              <option value="cp1250" selected>windows-1250</option>
              <option value="cp852">IBM CP852</option>
              <option value="utf-8">UTF-8</option>
            </select>
          </div>
          <div>
            <label for="identifierDate">Azonosító dátuma</label>
            <input id="identifierDate" type="date">
          </div>
        </div>
        <div class="import-actions">
          <button id="inspectBtn" class="primary">Beolvasás</button>
          <a id="templateLinkInDialog" class="button-link ghost" href="/template.xlsx" download>Excel sablon</a>
        </div>
        <details id="mappingDetails" class="advanced-mapping">
          <summary>Haladó: Excel oszlopmegfeleltetés</summary>
          <div class="import-section-title">
            <div>
              <h3>Excel oszlopok hozzárendelése</h3>
              <div class="muted-small">Nem a banki szabványt módosítja, csak azt állítja be, melyik Excel oszlop melyik fix mezőbe kerüljön.</div>
            </div>
            <button id="useGuessesBtn" class="secondary" type="button">Automatikus felismerés</button>
          </div>
          <div id="mappingArea" class="stack"><p>Még nincs beolvasott fejléc.</p></div>
        </details>
      </div>
    </dialog>`;
  }
}
if (!customElements.get("import-dialog")) customElements.define("import-dialog", ImportDialog);
