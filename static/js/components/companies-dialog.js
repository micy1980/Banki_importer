/* <companies-dialog id="companiesDialog"> */
class CompaniesDialog extends (window.DialogBase || HTMLElement) {
  get template() {
    return `
    <dialog aria-labelledby="companiesDialogTitle">
      <div class="dialog-head">
        <h2 id="companiesDialogTitle">Cégek</h2>
        <button id="closeCompaniesBtn" class="secondary" type="button" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <div class="account-grid">
          <section class="account-card">
            <h3>Új cég</h3>
            <label for="companyName">Cég neve</label>
            <input id="companyName" type="text" placeholder="pl. Minta Kft">
            <div class="account-actions">
              <button id="saveCompanyBtn" class="primary" type="button">Mentés</button>
            </div>
            <div id="companyStatus" class="status" role="status" aria-live="polite">A cégválasztó az import célcégét adja meg.</div>
          </section>
          <section class="account-card">
            <h3>Rögzített cégek</h3>
            <div id="companiesList" class="account-list"></div>
          </section>
        </div>
      </div>
    </dialog>`;
  }
}
if (!customElements.get("companies-dialog")) customElements.define("companies-dialog", CompaniesDialog);
