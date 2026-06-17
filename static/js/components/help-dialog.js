/* <help-dialog id="helpDialog"> */
class HelpDialog extends HTMLElement {
  connectedCallback() {
    if (this._ready) return;
    this._ready = true;
    this.innerHTML = `
      <dialog aria-labelledby="helpDialogTitle" class="help-dialog">
        <form method="dialog" class="dialog-form">
          <header class="dialog-header">
            <div>
              <h2 id="helpDialogTitle">Súgó</h2>
              <p>Import beolvasása, ellenőrzése és banki TXT export készítése.</p>
            </div>
            <button id="closeHelpBtn" class="secondary icon-only" type="button" data-close aria-label="Bezárás">×</button>
          </header>
          <div class="dialog-body">
            <div class="help-grid">
              <section>
                <h3>Alap folyamat</h3>
                <ol>
                  <li>Válaszd ki az aktív céget.</li>
                  <li>Import ablakban válassz bankot, formátumot és fájlt.</li>
                  <li>Beolvasás után a TXT letöltése gomb készíti el az importfájlt.</li>
                </ol>
              </section>
              <section>
                <h3>Mentés és napló</h3>
                <p>A Mentés (ZIP) a helyi törzsadatokat és naplót csomagolja. Az audit napló a Napló gombbal nyitható meg.</p>
              </section>
            </div>
          </div>
          <footer class="dialog-footer">
            <button class="primary" type="button" data-close>Rendben</button>
          </footer>
        </form>
      </dialog>
    `;
    const dialog = this.querySelector("dialog");
    this.querySelectorAll("[data-close]").forEach(button =>
      button.addEventListener("click", () => dialog.close()),
    );
  }
  showModal() { this.querySelector("dialog")?.showModal(); }
  close() { this.querySelector("dialog")?.close(); }
}

if (!customElements.get("help-dialog")) {
  customElements.define("help-dialog", HelpDialog);
}
