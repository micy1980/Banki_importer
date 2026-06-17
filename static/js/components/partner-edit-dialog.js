/* <partner-edit-dialog id="partnerEditDialog"> — Web Component verzió.
 * Az input-okat light DOM-ban tartja, így getElementById("editPartnerCode") stb.
 * a partners.js wiringja változatlanul működik.
 */
class PartnerEditDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
      <dialog aria-labelledby="partnerEditDialogTitle">
        <div class="dialog-head">
          <h2 id="partnerEditDialogTitle">Partner szerkesztése</h2>
          <button id="closePartnerEditBtn" class="secondary" type="button" data-close>Bezárás</button>
        </div>
        <div class="dialog-body">
          <div class="account-card">
            <div class="grid-2">
              <div><label for="editPartnerCode">Partner kód</label><input id="editPartnerCode" type="text"></div>
              <div><label for="editPartnerName">Név</label><input id="editPartnerName" type="text"></div>
              <div><label for="editPartnerAccount">Magyar számlaszám</label><input id="editPartnerAccount" type="text"></div>
              <div><label for="editPartnerIban">IBAN</label><input id="editPartnerIban" type="text"></div>
              <div><label for="editPartnerSwift">SWIFT/BIC</label><input id="editPartnerSwift" type="text"></div>
              <div><label for="editPartnerCountry">Ország</label><input id="editPartnerCountry" type="text"></div>
            </div>
            <label for="editPartnerAddress">Partner címe</label>
            <input id="editPartnerAddress" type="text">
            <div class="grid-2" style="margin-top:10px;">
              <div><label for="editPartnerBankName">Bank neve</label><input id="editPartnerBankName" type="text"></div>
              <div><label for="editPartnerBankAddress">Bank címe</label><input id="editPartnerBankAddress" type="text"></div>
            </div>
            <div class="account-actions">
              <button id="savePartnerEditBtn" class="primary" type="button">Mentés</button>
              <button id="cancelPartnerEditBtn" class="secondary" type="button" data-close>Mégse</button>
            </div>
            <div id="partnerEditStatus" class="status" role="status" aria-live="polite">Szerkesztésre megnyitva.</div>
          </div>
        </div>
      </dialog>
    `;
  }
}
if (!customElements.get("partner-edit-dialog")) {
  customElements.define("partner-edit-dialog", PartnerEditDialog);
}
