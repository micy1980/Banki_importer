/* <partners-dialog id="partnersDialog"> */
class PartnersDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
    <dialog aria-labelledby="partnersDialogTitle">
      <div class="dialog-head">
        <h2 id="partnersDialogTitle">Partnerlista</h2>
        <button id="closePartnersBtn" class="secondary" type="button" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <div class="account-grid">
          <section class="account-card">
            <h3>Kézi rögzítés</h3>
            <div class="grid-2">
              <div><label for="partnerCode">Partner kód</label><input id="partnerCode" type="text"></div>
              <div><label for="partnerName">Név</label><input id="partnerName" type="text"></div>
              <div><label for="partnerAccount">Magyar számlaszám</label><input id="partnerAccount" type="text" placeholder="12345678-12345678-12345678"></div>
              <div><label for="partnerIban">IBAN</label><input id="partnerIban" type="text"></div>
              <div><label for="partnerSwift">SWIFT/BIC</label><input id="partnerSwift" type="text"></div>
              <div><label for="partnerCountry">Ország</label><input id="partnerCountry" type="text" value="HU"></div>
            </div>
            <label for="partnerAddress">Partner címe</label>
            <input id="partnerAddress" type="text">
            <div class="grid-2" style="margin-top:10px;">
              <div><label for="partnerBankName">Bank neve</label><input id="partnerBankName" type="text"></div>
              <div><label for="partnerBankAddress">Bank címe</label><input id="partnerBankAddress" type="text"></div>
            </div>
            <div class="account-actions">
              <button id="savePartnerBtn" class="primary" type="button">Mentés</button>
              <button id="lookupPartnerBankBtn" class="secondary" type="button">Bankadat keresés</button>
            </div>
            <div id="partnerStatus" class="status" role="status" aria-live="polite">Magyar számlánál az első 8 számjegyből tölt banknevet.</div>
          </section>
          <section class="account-card">
            <h3>Import</h3>
            <label for="partnerImportFile">Partner import</label>
            <input id="partnerImportFile" type="file" accept=".xlsx,.xlsm,.csv">
            <div class="account-actions">
              <button id="importPartnersBtn" class="secondary" type="button">Importálás</button>
              <a class="button-link ghost" href="/partners-template.xlsx" download>Import sablon</a>
            </div>
          </section>
        </div>
        <section class="account-card" style="margin-top:14px;">
          <h3>Rögzített partnerek</h3>
          <div id="partnersList" class="account-list"></div>
        </section>
      </div>
    </dialog>`;
  }
}
if (!customElements.get("partners-dialog")) customElements.define("partners-dialog", PartnersDialog);
