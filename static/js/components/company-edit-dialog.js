/* <company-edit-dialog id="companyEditDialog"> */
class CompanyEditDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
    <dialog aria-labelledby="companyEditDialogTitle">
      <div class="dialog-head">
        <h2 id="companyEditDialogTitle">Cég szerkesztése</h2>
        <button id="closeCompanyEditBtn" class="secondary" type="button" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <section class="account-card form-card">
          <label for="editCompanyName">Cég neve</label>
          <input id="editCompanyName" type="text" autocomplete="organization">
          <div class="account-actions">
            <button id="saveCompanyEditBtn" class="primary" type="button">Mentés</button>
            <button id="cancelCompanyEditBtn" class="secondary" type="button">Mégse</button>
          </div>
          <div id="companyEditStatus" class="status" role="status" aria-live="polite"></div>
        </section>
      </div>
    </dialog>`;
  }
}

if (!customElements.get("company-edit-dialog")) {
  customElements.define("company-edit-dialog", CompanyEditDialog);
}
