/* <accounts-dialog id="accountsDialog"> */
class AccountsDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
    <dialog aria-labelledby="accountsDialogTitle">
      <div class="dialog-head">
        <h2 id="accountsDialogTitle">Saját bankszámlák listája - <span id="accountsCompanyName"></span></h2>
        <button id="closeAccountsBtn" class="secondary" type="button" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <div class="account-grid">
          <section class="account-card">
            <h3>Kézi rögzítés</h3>
            <div class="grid-2">
              <div>
                <label for="ownBankCountry">Bank országa</label>
                <select id="ownBankCountry"><option value="HU" selected>HU - Magyarország</option></select>
              </div>
              <div>
                <label for="ownBankName">Bank neve</label>
                <input id="ownBankName" type="text" placeholder="Automatikus is lehet a hitelesítő táblából">
              </div>
              <div><label for="ownCurrency">Deviza</label><select id="ownCurrency"></select></div>
            </div>
            <div style="margin-top:10px;">
              <label for="ownAccountNumber">Számlaszám vagy HU IBAN</label>
              <input id="ownAccountNumber" type="text" placeholder="12345678-12345678-12345678 vagy HU...">
            </div>
            <div class="account-actions">
              <button id="saveAccountBtn" class="primary" type="button">Mentés</button>
            </div>
            <div id="accountStatus" class="status" role="status" aria-live="polite">Nincs kiválasztott bankszámla.</div>
          </section>
          <section class="account-card">
            <h3>Import</h3>
            <label for="accountImportFile">Bankszámla import</label>
            <input id="accountImportFile" type="file" accept=".xlsx,.xlsm,.csv">
            <div class="account-actions">
              <button id="importAccountsBtn" class="secondary" type="button">Importálás</button>
              <a class="button-link ghost" href="/accounts-template.xlsx" download>Import sablon</a>
            </div>
            <div id="registryMeta" class="status" role="status" aria-live="polite">Hitelesítő tábla állapota betöltés alatt.</div>
          </section>
        </div>
        <section class="account-card" style="margin-top:14px;">
          <h3>Rögzített bankszámlák</h3>
          <div id="accountsList" class="account-list"></div>
        </section>
      </div>
    </dialog>`;
  }
}
if (!customElements.get("accounts-dialog")) customElements.define("accounts-dialog", AccountsDialog);
