/* DialogBase — közös ős a custom element dialogokhoz.
 * Használat:
 *   class FooDialog extends DialogBase {
 *     get template() { return `<dialog>...</dialog>`; }
 *   }
 *   customElements.define("foo-dialog", FooDialog);
 *
 * A komponens belül `<dialog>`-ot tart, így az openDialog("xxx") /
 * closeDialog("xxx") globális API változatlanul működik, ha az id-t
 * a host elemre tesszük: <foo-dialog id="xxx"></foo-dialog>
 * (a globális el(xxx).showModal a host showModal()-jét fogja hívni.)
 */
class DialogBase extends HTMLElement {
  connectedCallback() {
    if (this._ready) return;
    this._ready = true;
    this.innerHTML = this.template;
    this._dialog = this.querySelector("dialog");
    this.querySelectorAll("[data-close]").forEach(b =>
      b.addEventListener("click", () => this._dialog?.close()),
    );
    if (typeof this.onReady === "function") this.onReady();
  }
  get template() { return "<dialog><p>Override get template().</p></dialog>"; }
  showModal() { this._dialog?.showModal(); }
  close() { this._dialog?.close(); }
  get open() { return !!this._dialog?.open; }
}
window.DialogBase = DialogBase;
