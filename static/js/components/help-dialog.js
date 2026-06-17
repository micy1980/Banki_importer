/* Sample WebComponent: <help-dialog>
 * Bizonyítja a mintát — a többi dialog átírása fokozatosan jöhet.
 * Ha él, a meglévő helpDialog markup eltávolítható az app.py-ból,
 * és helyette: <help-dialog id="helpDialog"></help-dialog>
 */
class HelpDialog extends HTMLElement {
  connectedCallback() {
    if (this._ready) return;
    this._ready = true;
    this.innerHTML = `
      <dialog aria-labelledby="helpDialogTitle">
        <form method="dialog" class="dialog-form">
          <header class="dialog-header">
            <h2 id="helpDialogTitle">Súgó</h2>
            <button id="closeHelpBtn" class="ghost" type="button" data-close aria-label="Bezárás">×</button>
          </header>
          <div class="dialog-body">
            <p><strong>Mit csinál ez az eszköz?</strong> Banki Excel/CSV → PAYORD TXT konverter.</p>
            <ul>
              <li>Válassz aktív céget.</li>
              <li>Import → bank + formátum → fájl beolvasása.</li>
              <li>TXT letöltése.</li>
            </ul>
            <p>Backup: <code>/api/backup</code> (vagy a fejléc „Mentés (ZIP)” gombja).</p>
            <p>Audit log: <code>/api/audit?limit=200</code>.</p>
            <p>Health: <code>/healthz</code>.</p>
          </div>
          <footer class="dialog-footer">
            <button class="primary" type="button" data-close>Rendben</button>
          </footer>
        </form>
      </dialog>
    `;
    const dialog = this.querySelector("dialog");
    this.querySelectorAll("[data-close]").forEach(b =>
      b.addEventListener("click", () => dialog.close()),
    );
  }
  showModal() { this.querySelector("dialog")?.showModal(); }
  close() { this.querySelector("dialog")?.close(); }
}

if (!customElements.get("help-dialog")) {
  customElements.define("help-dialog", HelpDialog);
}
