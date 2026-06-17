/* <account-edit-dialog id="accountEditDialog"> */
class AccountEditDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
    <dialog aria-labelledby="accountEditDialogTitle">
      <div class="dialog-head">
        <h2 id="accountEditDialogTitle">Bankszámla szerkesztése</h2>
        <button id="closeAccountEditBtn" class="secondary" type="button" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <div class="account-card">
          <div class="grid-2">
            <div>
              <label for="editBankCountry">Bank országa</label>
              <select id="editBankCountry"><option value="HU" selected>HU - Magyarország</option></select>
            </div>
            <div><label for="editBankName">Bank neve</label><input id="editBankName" type="text"></div>
            <div><label for="editCurrency">Deviza</label><select id="editCurrency"></select></div>
          </div>
          <div style="margin-top:10px;">
            <label for="editAccountNumber">Számlaszám vagy HU IBAN</label>
            <input id="editAccountNumber" type="text">
          </div>
          <div class="account-actions">
            <button id="saveAccountEditBtn" class="primary" type="button">Mentés</button>
            <button id="cancelAccountEditBtn" class="secondary" type="button" data-close>Mégse</button>
          </div>
          <div id="accountEditStatus" class="status" role="status" aria-live="polite">Szerkesztésre megnyitva.</div>
        </div>
      </div>
    </dialog>`;
  }
}
if (!customElements.get("account-edit-dialog")) customElements.define("account-edit-dialog", AccountEditDialog);
